import win32com.client as win32
import warnings
import pyodbc
import pandas as pd
import SQL as s
from datetime import date
import Receiver as r

from datetime import date

warnings.filterwarnings('ignore')
today = date.today()


def send_mail(email_contacts):
    receivers = email_contacts
    cc = ['christoph.razek@emea-cosmetics.com', 'dzanana.dautefendic@emea-cosmetics.com']

    # creating an win32 object/mail object
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)

    # mail.Font.Name = 'Georgia'
    mail.To = ";".join(receivers)
    mail.CC = ";".join(cc)
    mail.Subject = f'QA Samples Reminder'
    mail.HTMLBody = """<font face='Calibri, Calibri, monospace'>
    Good Day, <br><br>
    
    Please send us the samples for the orders in the list as soon as possible.<br>
    In case there are problems, please inform us as soon as possible.<br>
    If you have any questions please feel free to contact kirstin.past@emea-cosmetics.com or kirsten.prohaska@emea-cosmetics.com.<br><br>
    
    <b>Please follow the sample instructions in the excel thoroughly!</b>
    <br><br>
    Thank you and kind regards.<br>
    <br>
    Linnea Jonasson<br> 
    Personal Assistant to CEO & Product Development<br>

    <br>emea Handelsgesellschaft mbH<br>
    Brucknerstraße 8/5<br>
    A-1040 Wien<br>
    Tel.:    +43 1 535 10 01 - 210<br>
    Fax:    +43 1 535 10 01 - 900<br>
    </font>"""
    mail.Attachments.Add(rf'S:\EMEA\Kontrollabfragen\VorabM_Reminder\QA-Sample-Reminder_{i}_{today}.xlsx')

    mail.Display()
    mail.Save()
    mail.Send()

#Datenbankverbindung
connx_string = r'DRIVER={SQL Server}; server=172.19.128.2\emeadb; database=emea_enventa_live45; UID=usr_razek; PWD=wB382^%H3INJ'
conx = pyodbc.connect(connx_string)

#Reading SQL for open Vorabmuster
df = pd.read_sql_query(s.offene_Vmuster, conx)

#Adjusting File
df[['PO','LIEFERANTENNR']] = df[['PO','LIEFERANTENNR']].astype('int64')

receiver = r.contacts

for i in receiver:

    df_mail = df[df['LIEFERANTENNR'] == i]
    df_mail.drop('LIEFERANTENNR', axis=1, inplace=True)
    df_mail.to_excel(rf'S:\EMEA\Kontrollabfragen\VorabM_Reminder\QA-Sample-Reminder_{i}_{today}.xlsx', index= False)
    if df_mail.shape[0] != 0:
        send_mail(receiver[i])

#Log

with open(r'S:\EMEA\Kontrollabfragen\Vorabmuster_Reminder.txt','a') as file:
    file.write(f'\nVorabMuster Reminder wurde zuletzt am {today} verschickt!')
    file.close()