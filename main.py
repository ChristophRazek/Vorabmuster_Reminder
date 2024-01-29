import win32com.client as win32
import warnings
import Update as u
import pandas as pd
import numpy as np

from datetime import date

warnings.filterwarnings('ignore')
today = date.today()

#Führe Update aus
df_samples = u.update()



def send_reminder(df_samples):

    receivers = {'NUCO': 'yian.su@emea-cosmetics.com','ANCOROTTI': 'yian.su@emea-cosmetics.com','ART': 'yian.su@emea-cosmetics.com'}
    cc = ['christoph.razek@emea-cosmetics.com','dzanana.dautefendic@emea-cosmetics.com']

    #print(df_samples.to_markdown())


    #Reminder für alle Artikel für die Samples benötigt werden und noch nicht eingetroffen sind
    df_samples['PE14_SampleReceived'] = df_samples['PE14_SampleReceived'].fillna('0000-00-00 00:00:00')
    df_samples['Today'] = today
    df_samples['Today'] = pd.to_datetime(df_samples['Today'])
    df_samples['diff_days'] = (df_samples['LIEFERDATUM'] - df_samples['Today']) / np.timedelta64(1, 'D')

    # ZEITDIFFERENZ NOCH ZU ERMITTELN!!!!!!
    #Prüfung und EMAIL wenn KEINE Reminder

    df_reminder = df_samples[(df_samples['VorabM_Pflicht'] == 1) & (df_samples['PE14_SampleReceived'] == '0000-00-00 00:00:00') & (df_samples['diff_days'] < 14)]
    companies = set(df_reminder['SUCHNAME'].tolist())

    #Reminder für jede Firma einzeln
    for c in companies:

        if len(companies) == 0:
            # creating an win32 object/mail object
            outlook = win32.Dispatch('outlook.application')
            mail = outlook.CreateItem(0)

            mail.To = ['christoph.razek@emea-cosmetics.com','yian.su@emea-cosmetics.com']
            mail.Subject = f'Kein Reminder am {today} versendet, da keine Liefertermine anstehen'

            mail.Display()
            mail.Save()
            # mail.Send()

        elif c in receivers:
            df_attachmment = df_reminder[df_reminder['SUCHNAME']== c].drop(['FIXPOSNR','BELEGART','VorabM_Pflicht',
                                                                    'PE14_SampleReceived', 'Today', 'diff_days'], axis=1)
            df_attachmment.rename(columns={'BELEGNR':'PO','SUCHNAME':'SUPPLIER', 'ARTIKELNR':'ARTICLE',
                                   'BEZEICHNUNG':'DESCRIPTION', 'LIEFERDATUM':'DELIVERY-DATE' }, inplace=True)

            #Zwischenspeichern für Attachment
            df_attachmment.to_excel(rf'S:\EMEA\Kontrollabfragen\VorabM_Reminder\{c}_Sample_Reminder.xlsx', index=False)


            # creating an win32 object/mail object
            outlook = win32.Dispatch('outlook.application')
            mail = outlook.CreateItem(0)

            mail.To = receivers[c]
            mail.CC = ";".join(cc)
            mail.Subject = f'Reminder for Sample: {c}'
            mail.HTMLBody = """<font face='Calibri, Calibri, monospace'>
            Good Day, <br><br>
            Please send us the Production Samples for the Articles in the list attached as the initial delivery dates will soon be reached.<br>
            In case there are problems, please inform us as soon as possible.<br>
            If you have any questions please feel free to contact me (yian.su@emea-cosmetics.com).<br><br>
            Thank you and kind regards.<br>
            <br>
            Yian
            </font>"""
            mail.Attachments.Add(rf'S:\EMEA\Kontrollabfragen\VorabM_Reminder\{c}_Sample_Reminder.xlsx')

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
            #mail.Send()




send_reminder(df_samples)