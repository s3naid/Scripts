import os
import shutil
import pandas as pd
import re
from sympy import Segment,Point
from mailmerge import MailMerge
from datetime import date
###########################################################

class Horizont:
    """ Klasa horizont predstavlja pojedine horizonte na kojima se nalazi više radilišta
        Hardkoded je dio koji uzima orginalni fajl (3D model radilišta je snimljen u .str formatu koji je sličan CSVu)
        i kopira ga u radni direktorij skripte
        Svaki fajl se očisti i time postaje pravi CSV fajl, u kojem se nalaze koordinate XYZ, opis i segment
        CSV fajl se zatim dijeli na više txt fajlvoa prema segmentima a od kojih je potreban samo segment 5
        CSV segmenta 5 se koristi za raunanje dužine koji predstavlja dužinu napredovanja

    """
    def __init__(self, horizont,br_string=5, stari=False):
        self.horizont = horizont.lower()
        self.br_string=br_string.lower()
        self.stari=stari
        pass
    def dct(self):
        script_dir = os.getcwd()
        dest_dir=os.path.join(script_dir,str(self.horizont)) 
        try:
            os.makedirs(dest_dir) 
        except OSError: 
            pass
        return dest_dir
    
    def kopiranje(self):
        #kopiranje orginalno fajla u radni direktorij
        dest_dir=self.dct()
        if self.stari:
            shutil.copy((r'Staro\\'+str(self.horizont)+'_radni.str'),dest_dir+'\\'+str(self.horizont)+'.str')
            return (os.path.join(dest_dir,str(self.horizont)+'.str'))
        else:
            shutil.copy((r'Azurno\\'+str(self.horizont)+'_radni.str'),dest_dir+'\\'+str(self.horizont)+'.str')
            return (os.path.join(dest_dir,str(self.horizont)+'.str'))
 
    def ucitavanje_str(self):
        #čišćenje fajlova kako bi se dobio pravi CSV
        file=self.kopiranje()
        sFile=pd.read_csv(file, sep=',', names=["string","X","Y","Z","opis"], dtype=str,skipinitialspace=True)
        del sFile['opis']
        start=sFile.where(sFile['string']==self.br_string).first_valid_index()
        end=(sFile.where(sFile['string']==self.br_string).last_valid_index())+2
        df=sFile.iloc[start:end]
        df.reset_index(drop=True, inplace=True)
        #df.to_csv(os.path.join(dest_dir,str(self.horizont)+'_'+str(self.br_string)+'.txt'))     #aktivirati samo ako treba upisivati u novi fajl
        return df
    
    def segmentiranje(self):
        #dijeljenje fajla prema segmentima i snimanje istih
        dest_dir=self.dct()
        df=self.ucitavanje_str()
        dest_dir_seg=os.path.join(dest_dir, 'Segmenti_stringa_'+str(self.br_string))
        try: 
            os.makedirs(dest_dir_seg) 
        except OSError: 
            pass

        tacke_segmenta=[]
        ime_file=[]
        break_index = 0  # trenutni prekid
        chunk_index = 0  # index linije
        for index, row in df.iterrows():
            if row.string == "0":  # ako je prekid
                df[break_index:index].to_csv(os.path.join(dest_dir_seg,'file_'+ str(chunk_index)+ '.txt')) 
                tacke_segmenta.append(index-break_index)
                ime_file.append(str(self.horizont)+'_file_'+ str(chunk_index))
                break_index = index+1  
                chunk_index += 1 
        if self.br_string=='5':
            seg=pd.DataFrame(list(zip(tacke_segmenta,ime_file)), columns=['Br_tacaka','Ime'])
            seg['horizont']=self.horizont
            return seg
        
    def sorted_aphanumeric(self,data):
        convert = lambda text: int(text) if text.isdigit() else text.lower()
        alphanum_key = lambda key: [ convert(c) for c in re.split('([0-9]+)', key)] 
        return sorted(data, key=alphanum_key)

    def duzina_segmenta(self,df):
        #računanje dužine segmenta
        nula=0
        red=[]
        for index, row in df.iterrows():
            if (index!=nula):
                previous=Point(df['Y'][nula],df['X'][nula], evaluate = False)
                current=Point(df['Y'][index],df['X'][index], evaluate = False)
                duzina=Segment(previous,current).length
                nula+=1
                red.append(duzina)
        return red
    
    def duzina_stringa(self, x=None):
        #raunanje dužine stringa (jedan string može sadržavati više segmenata a gdje svaki segment predstavlja radilište)
        current=self.segmentiranje()
        direktorij=self.dct()
        directory = os.path.join(direktorij,'Segmenti_stringa_5') 
        files=[os.path.join(directory, f) 
                for f in self.sorted_aphanumeric(os.listdir(directory))] 
        suma=[]
        if x==None:
            for file in files:
                df=pd.read_csv(file,header=0)
                red=self.duzina_segmenta(df)
                suma.append(sum(red))    
            current['Duz']=pd.Series(suma)
            #current.to_csv(os.path.join(direktorij,'current.txt'))             
        else:
            i=0
            x.sort()      
            for file in files:
                try:
                    if file.endswith('file_'+str(x[i])+'.txt'):
                        df=pd.read_csv(file,header=0)
                        red=self.duzina_segmenta(df)
                        suma.append(sum(red))    
                        i+=1
                        current=pd.DataFrame(suma, columns=['Duz'])
                except: pass
        return current

class KomparatorHorizonta:
    """ Clasa kojom se vrši poređenje pojedinih Horizonata kako bi se mogle pronaći promjene dužine, odnosno
        dobivamo informaciju na kojem se radilištu radilo i koliko se napredovalo
    """
    _registry = []
    def __init__(self, novo, staro):
      self.novo = novo
      self.staro = staro
      self._registry=KomparatorHorizonta._registry
    
    def razlike_duzina(self):
      df1 = self.novo
      df2 = self.staro

      df=pd.merge(df1,df2[['Ime','Duz']], on='Ime',how='left',suffixes=('_current','_previous'))
      df=df.fillna(0)
      df['Razlika']=df['Duz_current']-df['Duz_previous']
      return df
      
    def radilista_napredovanje(self):
      df=self.razlike_duzina()
      index=pd.DataFrame()
      for i,row in df.iterrows():
          if (row.Razlika!=0.0):
              index=index.append(row)
      self._registry.append(index)
      return index

class Radilista:
    """ Priprema podataka za izradu izvještaja
        Popunjavanje izvještaja
    """
    def __init__(self, lista):
      self.lista = lista
      
    def grupisanjeDF(self):
        df = pd.concat(self.lista)
        script_dir= os.getcwd()
        radilista=pd.read_csv(script_dir+'\\'+'Radilista.txt',index_col=[0],header=0)
        big_df=pd.merge(radilista,df, how='inner', on='Ime', left_index=True)
        return big_df
                
    def izvjestaj(self,template):
        ddatum=date.today()
        ddatum=ddatum.strftime("%d/%b/%y")
        document = MailMerge(template)
        df=self.grupisanjeDF()
        df['masa'] = df.apply(lambda row: row.Razlika*12.5
                                    if row.profil=='3.5x3.5'
                                    else row.Razlika*9, axis=1)
        df['tone'] = df['masa'].apply(lambda x: x*1.6)
        df=df.round({'Duz_current':3,'Duz_previous':3,'Razlika':3,'masa':3,'tone':3}).applymap(str)
        dct_table=df.to_dict('record')
        document.merge(datum=ddatum)
        document.merge_rows('Duz_current', dct_table)
        document.merge_rows('tone', dct_table)
        document.write('test.docx')
        history=pd.DataFrame(list(zip(df['radiliste'],df['Duz_current'])), columns=['radiliste',ddatum])
        return history

###########################################################     

#Main program
h1=Horizont('H710','5')
h2=Horizont('H710','5',True)
h3=Horizont('GTR','5')
h4=Horizont('GTR','5',True)

ds1=h1.duzina_stringa()
ds2=h2.duzina_stringa()
kr1=KomparatorHorizonta(ds1,ds2).radilista_napredovanje()
#
ds3=h3.duzina_stringa()
ds4=h4.duzina_stringa()
kr2=KomparatorHorizonta(ds3,ds4).radilista_napredovanje()
#lista=h4.lista()

#x=Radilista(KomparatorHorizonta._registry)  
big_df=Radilista(KomparatorHorizonta._registry).grupisanjeDF()
h=Radilista(KomparatorHorizonta._registry).izvjestaj('Weekly_report_Olovo_SD.docx')