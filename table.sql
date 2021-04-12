-- Таблица №1. Непокрытые внутроссийские аккредитивы.
create table RU_LC
(
GOSB_ID  VARCHAR2(100),
DATE_OF_CONCLUSION  DATE  not null,
PRODUCT VARCHAR2(100),
NAME_OF_THE_CLIENT VARCHAR2(100),
CURRENCY VARCHAR2(5),
VOLUME_NKD NUMBER(30) not null,
MANAGER VARCHAR2(100)
)
/

-- Таблица №2. Международные аккредитивы.
create table int_LC
(
GOSB_ID  VARCHAR2(100),
DATE_OF_CONCLUSION  DATE  not null,
PRODUCT VARCHAR2(100),
NAME_OF_THE_CLIENT VARCHAR2(100),
CURRENCY VARCHAR2(5),
VOLUME_NKD NUMBER(30) not null,
MANAGER VARCHAR2(100)
)
/

-- Таблица №3. Покрытые аккредитивы ВРА
create table cove_LC
(
GOSB_ID  VARCHAR2(100),
DATE_OF_CONCLUSION  DATE  not null,
PRODUCT VARCHAR2(100),
NAME_OF_THE_CLIENT VARCHAR2(100),
CURRENCY VARCHAR2(5),
VOLUME_NKD NUMBER(30) not null,
MANAGER VARCHAR2(100)
)

-- Подсчет рейтинга

select manager, sum(summ) as summ
from (
select manager,case when CURRENCY='EUR' then VOLUME_NKD*90 when CURRENCY='USD' then VOLUME_NKD*75
else VOLUME_NKD end as summ
from int_LC
union all
select manager,case when CURRENCY='EUR' then VOLUME_NKD*90 when CURRENCY='USD' then VOLUME_NKD*75
else VOLUME_NKD end as summ
from cove_LC
union all
select manager,case when CURRENCY='EUR' then VOLUME_NKD*90 when CURRENCY='USD' then VOLUME_NKD*75
else VOLUME_NKD end as summ
from RU_LC)
group by manager
order by summ desc


select manager, sum(summ) as summ
                from (
                select manager,case when CURRENCY='EUR' then VOLUME_NKD*90 when CURRENCY='USD' then VOLUME_NKD*75
                end as summ
                from int_LC
                where CURRENCY in ('EUR','USD'))
                group by manager
                order by summ desc

select manager, sum(VOLUME_NKD) as summ
                from (
                select manager,VOLUME_NKD
                from cove_LC 
                where CURRENCY='RUR'
                union all
                select manager,VOLUME_NKD
                from RU_LC
                where CURRENCY='RUR')
                group by manager
                order by summ desc
