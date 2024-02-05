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

    receivers = {'NUCO': ['a.adamczewska@nuco.pl', 'k.jeziorska@nuco.pl'],
                 'ANCOROTTI': ['vbonazza@ancorotticosmetics.com', 'abrambini@ancorotticosmetics.com', 'smonticelli@ancorotticosmetics.com'],
                 'ART': ['chiara.bonacina@artcosmetics.it', 'valentina.ghillini@artcosmetics.it'],
                 'CHROMAVIS':'magdalena.skupinska@chromavis.com',
                 'CONFALON':'sonia@confaloniericosmetica.com',
                 'COSMETEC':['balzarienrica@cosmetec.it', 'giuliastrada@cosmetec.it'],
                 'COSMONDE':['jana.wirt@cosmonde.cz', 'eliska.kockova@cosmonde.cz'],
                 'DELIA':['skupinskak@delia.pl', 'klaudia.stoinska@delia.pl'],
                 'DONAUKANOL':['eldina.takmak@donau-kanol.com','sarah.prentner@donau-kanol.com'],
                 'FABER-CAST':'karin.puerner@fc-cosmetics.com',
                 'FIABILA':['stephanie.hardouin-letellier@fiabila.com','cindy.berthelot@fiabila.com'],
                 'ICC': ['m.lodolo@icc-italy.com','f.bellavita@icc-italy.com'],
                 'ITIT':['s.marzorati@ititcosmetics.it', 's.puce@ititcosmetics.it'],
                 'JOVI': ['tkostova@jovi.es', 'cscos@jovi.es'],
                 'PHARMA COS': ['michela.bruno@pharmacos.it', 'isabelle.diliddo@pharmacos.it'],
                 'PHARMACF':['joanna.kostrz@pharmacf.com.pl', 'olga.wesolowska@pharmacf.com.pl'],
                 'R&D COLOR': ['alessandrabertagna@redcolor.it', 'robertabattarola@redcolor.it', 'jessicacioccolini@redcolor.it'],
                 'STEP': 'customercare@stepcosmetici.com',
                 'OXY':['eekinci@oxygendevelopment.com', 'emea-specifications-de@oxygendevelopment.com', 'ablunck@oxygendevelopment.com']}

    cc = ['christoph.razek@emea-cosmetics.com','yian.su@emea-cosmetics.com','dzanana.dautefendic@emea-cosmetics.com']




    #Auffüllen der Zellen ohne Info ob erhalten
    df_samples['PE14_SampleReceived'] = df_samples['PE14_SampleReceived'].fillna('0000-00-00 00:00:00')
    df_samples['Today'] = today
    df_samples['Today'] = pd.to_datetime(df_samples['Today'])
    #Ermittlung Differenz von Lieferdatum zu heute
    df_samples['diff_days'] = (df_samples['LIEFERDATUM'] - df_samples['Today']) / np.timedelta64(1, 'D')

    #Vorabmusterpflicht, keine Erhalt-Info und Liefertermin in 14 fällig
    df_reminder = df_samples[(df_samples['VorabM_Pflicht'] == 1) & (df_samples['PE14_SampleReceived'] == '0000-00-00 00:00:00') & (df_samples['diff_days'] < 14)]
    companies = set(df_reminder['SUCHNAME'].tolist())

    #print(len(companies))
    #breakpoint()
    # Wenn kein Reminder notwendig: info an User
    if len(companies) == 0:
        # creating an win32 object/mail object
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)

        mail.To = ";".join(cc)
        mail.Subject = f'Kein Reminder am {today} versendet, da keine Liefertermine anstehen'

        mail.Display()
        mail.Save()
        # mail.Send()
    else:
        #Reminder für jede Firma einzeln
        for c in companies:


            if c in receivers:
                df_attachmment = df_reminder[df_reminder['SUCHNAME']== c].drop(['FIXPOSNR','BELEGART','VorabM_Pflicht',
                                                                        'PE14_SampleReceived', 'Today', 'diff_days'], axis=1)
                df_attachmment.rename(columns={'BELEGNR':'PO','SUCHNAME':'SUPPLIER', 'ARTIKELNR':'ARTICLE',
                                       'BEZEICHNUNG':'DESCRIPTION', 'LIEFERDATUM':'DELIVERY-DATE' }, inplace=True)

                #Zwischenspeichern für Attachment
                df_attachmment.to_excel(rf'S:\EMEA\Kontrollabfragen\VorabM_Reminder\{c}_Sample_Reminder.xlsx', index=False)


                # creating an win32 object/mail object
                outlook = win32.Dispatch('outlook.application')
                mail = outlook.CreateItem(0)

                if type(receivers[c])!= list:
                    mail.To = receivers[c]
                else:
                    mail.To = ";".join(receivers[c])
                mail.CC = ";".join(cc)
                mail.Subject = f'Reminder for Sample: {c}'
                mail.HTMLBody = """<font face='Calibri, Calibri, monospace'>
                Good Day, <br><br>
                Please send us the Production Samples for the Articles in the list attached as the initial delivery dates will soon be reached.<br>
                In case there are problems, please inform us as soon as possible.<br>
                If you have any questions please feel free to contact me (yian.su@emea-cosmetics.com).<br><br>
                Thank you and kind regards.<br>
                <br>
                Yian<br>
                <img src='S:/EMEA/Kontrollabfragen/VorabM_Reminder/Logo.png' width='200' height='100'>
                <br>
                emea Handelsgesellschaft mbH<br>
                Brucknerstraße 8/5<br>
                A-1040 Wien<br>
                Tel.:    +43 1 535 10 01 - 232<br>
                
                </font>
                """
                mail.Attachments.Add(rf'S:\EMEA\Kontrollabfragen\VorabM_Reminder\{c}_Sample_Reminder.xlsx')

                mail.Display()
                mail.Save()
                #mail.Send()
            else:
                # Wenn keine Email hinterlegt ist, info an Entwickler
                outlook = win32.Dispatch('outlook.application')
                mail = outlook.CreateItem(0)

                mail.To = 'christoph.razek@emea-cosmetics.com'
                mail.Subject = f'Fehlende Email Adresse {c} für Vorabmuster'


                mail.Display()
                mail.Save()
                #mail.Send()




send_reminder(df_samples)