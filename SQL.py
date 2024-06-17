offene_Vmuster = """

with cte_bgr as (
Select Min(Belegnr) as 'Belegnr', Artikelnr
from emea_enventa_live.dbo.BESTELLPOS
where BELEGART not in (2,191) and STATUS <> 6 and BELEGNR > 23000000
Group by ARTIKELNR
),

cte_bp_ohne_storno as (select belegnr, artikelnr, min(fixposnr) as 'fixposnr'
from [emea_enventa_live].[dbo].[BESTELLPOS]
where BranchKey = 110 and STATUS <> 6
group by belegnr, artikelnr
)

select --bp.Fixposnr, bp.Belegart, 
cte_bgr.Belegnr as 'PO', cte_bgr.Artikelnr as 'ARTICLE',  bp.Bezeichnung as 'DESCRIPTION',
l.LIEFERANTENNR, 
bp.Lieferdatum as 'DELIVERY_DATE',
/*case when bp.Status in (3,4) then 'geliefert'
	 when bp.Status in (1,2) then 'offen'
	 else 'fehler'
End as 'Status'
 ,case when bp.[PE14_WE_Note4] = 1 then 'Ja'
 else 'Nein'
 End as 'Vorabmuster',

 bp.[PE14_SampleReceived] as 'Muster-erhalten',
 
 case when [PE14_MassProdRel] = 1 then 'Ja'
 else 'Nein'
 End as 'Freigabe',
 */
 bp.[PE14_WE_Vessel] 'REMARKS'
 --, DATEDIFF(DAY,  dateadd(day,14,SYSDATETIME()),bp.LIEFERDATUM) as 'Datediff'
-- ,dateadd(day,14,SYSDATETIME()) as 'heute + 14'

 

from cte_bgr
left join cte_bp_ohne_storno
on cte_bgr.Belegnr = cte_bp_ohne_storno.BELEGNR and cte_bgr.ARTIKELNR = cte_bp_ohne_storno.ARTIKELNR

left join emea_enventa_live.dbo.BESTELLPOS as bp
on cte_bgr.Belegnr = bp.BELEGNR and cte_bgr.Artikelnr = bp.ARTIKELNR and cte_bp_ohne_storno.fixposnr = bp.FIXPOSNR

left join emea_enventa_live.dbo.LIEFERANTEN as l
on bp.LIEFERANTENNR = l.LIEFERANTENNR

where DATEDIFF(DAY, bp.LIEFERDATUM, SYSDATETIME()) < 90 -- filtert Lieferdati raus, die 3 Monate in der Vergangenheit liegen
and bp.STATUS = 1 and bp.[PE14_WE_Note4] = 1 and bp.[PE14_SampleReceived] is null and DATEDIFF(DAY,  dateadd(day,14,SYSDATETIME()),bp.LIEFERDATUM) <= 0
order by bp.LIEFERDATUM

"""