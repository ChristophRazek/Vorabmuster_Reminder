import pandas as pd
import warnings
from tkinter import messagebox
from datetime import datetime
import os

warnings.filterwarnings('ignore')

link = r'C:\Users\ChristophRazek\Emea\06_Qualitymanagement - Dokumente\01_QS\04_MPS\Vorabmuster.xlsx'

#Zeitstempel Letzte Änderung!
c_time = os.path.getmtime(link)
dt_c = datetime.fromtimestamp(c_time)

def update():
    #Kopieren von Vorabmusterfile File zu Laufwerk
    samples = pd.read_excel(link)

    samples[['VorabM_Pflicht', 'PE14_MassProdRel','FIXPOSNR','BELEGART','BELEGNR']] = samples[['VorabM_Pflicht','PE14_MassProdRel','FIXPOSNR','BELEGART','BELEGNR']].fillna(0).astype('int64')
    samples.to_csv(r'L:\Q\Vorabmuster.csv', sep=';', index=False)

    #Log File
    with open(r'S:\EMEA\Kontrollabfragen\Vorabmuster.txt', 'w') as f:
        f.write(f'Last Vorabmuster copied at: {dt_c}')
        f.close()
    return samples
