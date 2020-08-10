set define off;

  CREATE OR REPLACE PROCEDURE "GHIA"."HIGHPOLLUTION" (
minKod in number, maxKod in number, startDate in date, endDate in date, c1 out SYS_REFCURSOR) IS  
BEGIN
open c1 for

select 
kitnp.NAME, fn.bas, ing.NAME, TO_CHAR(fn.�����_����������_��), TO_CHAR((decode(fn.cod,65,fn.����_nt,67,fn.����_nt,round(fn.����_nt/ing.pdk,0))), 'TM','nls_numeric_characters=''. ''') as ����,
TO_CHAR(fn.dt,'dd.mm'),'' as ist, obl.NAZOBL

from 
------------------�������� ������---------------------------------------------------------------------
(select distinct dt, ma.bas, ma.nt, ma.cod, ma.�����_����������_��, ma.����_nt,ma.b_order  
from ghia.prob_body a, ghia.prob_zag pz, 

(select distinct ak.nt,
ak.cod,
sum(ak.�����_����������_��) as �����_����������_��,
max(ak.����) as ����_nt, 
ak.bas,
ak.b_order
from 
--------------------------------------------------------------------------------
(select distinct c.kn, c.cod,
count(*) �����_����������_��,
max(c.znach) as ����,
e.g,
e.nb,
nt.nt,
ba.NAME as bas,
ba.num as b_order
from
--------------�������� unic � ��������------------------------
(select distinct a.unic, a.cod, a.znach, pz.kn, pz.ku 
from ghia.prob_body a, ghia.ing b, ghia.prob_zag pz 
where a.cod = b.cod and a.znach / b.pdk >= b.vz and a.znach / b.pdk<b.EVZ and a.cod <> 35
and a.unic = pz.unic and dt >= startDate and dt <= endDate) c,
---------------------------------------------------------------
ghia.ntkn nt, KPH2012.KPH_NEW e, BASS_WITHNUM ba, novoch.kit1_9_new ki
where c.kn = e.kn  and e.KVO >= minKod and e.KVO<=maxKod and  nt.kn = c.kn and c.ku = e.nu and c.ku = nt.ku and ba.kb = e.nb and ba.kggr = e.g 
and e.pr1<>'t' and nt.nt=ki.nt
group by c.kn, c.cod, nt.nt, e.nb,e.g, ba.name, ba.NUM) ak
----------------------------------------------------------------
group by ak.nt, ak.cod, ak.g, ak.nb, ak.bas, ak.b_order
order by ak.nt, ak.cod) ma 

where ma.cod = a.cod and a.znach = ma.����_nt 
and pz.KN in (select distinct kn from ghia.ntkn nt
              where ma.nt = nt) and a.unic = pz.unic and dt >= startDate and dt <= endDate) fn,
--------------------------------------------------------------------------------------------------------------
ghia.ing ing, NOVOCH.KIT1_9_NEW kit, ghia.OBLASTY_1 obl, KIT_NPP_UPDATE kitnp

where ing.cod = fn.cod and kit.NT = fn.nt and kit.KO = obl.KO and fn.cod=ing.cod and kitnp.nt=fn.nt

order by fn.b_order, substr(kit.nt,1,1), kit.npp, fn.cod;
END;

/
