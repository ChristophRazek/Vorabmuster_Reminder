import win32com.client as win32
import warnings
import Update as u
import pandas as pd

from datetime import date

warnings.filterwarnings('ignore')
today = date.today()

#Führe Update aus
df_samples = u.update()



def send_reminder(df_samples):

    receivers = {'NUCO': 'yian.su@emea-cosmetics.com','ANCOROTTI': 'yian.su@emea-cosmetics.com','ART': 'yian.su@emea-cosmetics.com'}
    cc = ['christoph.razek@emea-cosmetics.com','dzanana.dautefendic@emea-cosmetics.com']

    #Reminder für alle Artikel für die Samples benötigt werden und noch nicht eingetroffen sind
    df_reminder = df_samples[(df_samples['VorabM_Pflicht'] == 1) & (df_samples['PE14_SampleReceived'] != 'nan')]
    companies = set(df_reminder['SUCHNAME'].tolist())

    #Reminder für jede Firma einzeln
    for c in companies:

        if c in receivers:
            df_attachmment = df_reminder[df_reminder['SUCHNAME']== c].drop(['FIXPOSNR','BELEGART','VorabM_Pflicht','PE14_SampleReceived','PE14_MassProdRel'], axis=1)

            #Zwischenspeichern für Attachment
            df_attachmment.to_excel(rf'S:\EMEA\Kontrollabfragen\{c}_Sample_Reminder.xlsx', index=False)


            # creating an win32 object/mail object
            outlook = win32.Dispatch('outlook.application')
            mail = outlook.CreateItem(0)

            mail.To = receivers[c]
            mail.CC = ";".join(cc)
            mail.Subject = f'Reminder for Sample,{c}'
            mail.HTMLBody = """<font face='Calibri, Calibri, monospace'>
            Good Day, <br><br>
            Please send us the Production Samples for the Articles in the list attached as the initial delivery dates will soon be reached.<br>
            In case there are problems, please inform us as soon as possible.<br>
            If you have any questions please feel free to contact me (yian.su@emea-cosmetics.com).<br><br>
            Thank you and kind regards.<br>
            <br>
            Yian
            </font>"""
            mail.Attachments.Add(rf'S:\EMEA\Kontrollabfragen\{c}_Sample_Reminder.xlsx')

            mail.Display()
            mail.Save()
            #mail.Send()
        else:
            # creating an win32 object/mail object
            outlook = win32.Dispatch('outlook.application')
            mail = outlook.CreateItem(0)

            mail.To = 'christoph.razek@emea-cosmetics.com'
            mail.Subject = f'Fehlende Email Adresse {c} für Vorabmuster'


            mail.Display()
            mail.Save()
            mail.Send()




send_reminder(df_samples)