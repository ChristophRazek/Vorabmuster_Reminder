offene_Vmuster = """SELECT 
FIXPOSNR, BELEGART,[BELEGNR] as 'PO',[ARTIKELNR] as 'ARTICLE',[BEZEICHNUNG] as 'DESCRIPTION',[MENGE_BESTELLT] as 'QTY',[LIEFERDATUM] as 'DELIVERY-DATE',[LIEFERANTENNR]

--,[PE14_WE_Note4],[PE14_SampleReceived]
  
  FROM [emea_enventa_live].[dbo].[BESTELLPOS]
  where BELEGART not in (2,191) and STATUS = 1 
  and [PE14_WE_Note4] = '1' and PE14_SampleReceived is null
  order by LIEFERANTENNR"""