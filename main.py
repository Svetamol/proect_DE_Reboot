# -*- coding: utf-8 -*-
"""
Created on Sat Apr 10 15:58:31 2021
@author: Администратор
"""


from tkinter import *
from tkinter.ttk import Combobox 
from sqlalchemy import *
from sqlalchemy_utils import *
from tkinter import Text


import sqlalchemy as adb
import cx_Oracle as ora
from sqlalchemy import MetaData, Table 
import datetime 
import pandas as pd
import openpyxl,xlsxwriter,xlrd

# Экспорт данных из таблицы в файл
def export_xls(namefile):
    try:
        text = ent.get('1.0', END)
        flag_ = True
        name = ''
        arr_name = []
        sum = ''
        arr_sum = []

        ch = True
        ch_z = False
        for buk in text:
            if buk != ' ' and buk != '\n':
                if flag_:
                    if ch or ch_z:
                        name += buk
                    else:
                        name += ' ' + buk
                        ch_z = True
                else:
                    sum += buk
            elif buk == ' ':
                if ch == True:
                    ch = False
                else:
                    flag_ = False
            elif buk == '\n':
                if not flag_:
                    arr_name.append(name)
                    name = ''
                    arr_sum.append(sum)
                    sum = ''
                    flag_ = True
                    ch = True
                    ch_z = False
           

        df = pd.DataFrame({'Manager': arr_name, 'Summa':arr_sum})
        df.to_excel('./{}.xlsx'.format(namefile))
        print('Данные экспортированы в файл')
        lbl_exl_2.grid_remove()
        lbl_exl.grid(column=2, row=22, padx=3, pady=3)
    except:
        lbl_exl.grid_remove()
        lbl_exl_2.grid(column=2, row=22, padx=3, pady=3)


# Вставка данных в БД из формы интерфейса

def rating(l_conn_ora):
    sql = """select manager, sum(summ) as summ
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
            order by summ desc"""

    # Подсчет общего рейтинга
    l_sql_exec = l_conn_ora.connect()
    l_pre_res = l_sql_exec.execute(sql)

    ent.delete('1.0', END)
    for data in l_pre_res.fetchall():
        ent.insert(END, '\n{} {}'.format(data[0], data[1]))

# Подсчет итогового рейтинга     
def rating2(l_conn_ora):
    if var2.get()==0:
        
        sql = """select manager, sum(VOLUME_NKD) as summ
                from (
                select manager,VOLUME_NKD
                from cove_LC 
                where CURRENCY='RUR'
                union all
                select manager,VOLUME_NKD
                from RU_LC
                where CURRENCY='RUR')
                group by manager
                order by summ desc"""
                
    else: 
        sql = """select manager, sum(summ) as summ
                from (
                select manager,case when CURRENCY='EUR' then VOLUME_NKD*90 when CURRENCY='USD' then VOLUME_NKD*75
                end as summ
                from int_LC
                where CURRENCY in ('EUR','USD'))
                group by manager
                order by summ desc"""

    l_sql_exec = l_conn_ora.connect()
    l_pre_res = l_sql_exec.execute(sql)

    ent.delete('1.0', END)
    for data in l_pre_res.fetchall():
        ent.insert(END, '\n{} {}'.format(data[0], data[1]))
        


def insertbd(l_conn_ora):
    flag = False # флаг ошибки - если он хоть раз поднят, значит какие данные введены не корректно

    territory = str(combo2.get())
    if len(territory) == 0: # проверка что данных введены
        lbl_none1.grid(column=1, row=3)
        flag = True
    else:
        lbl_none1.grid_remove()

    if len(combo3.get()) == 0:
        lbl_none2.grid(column=1, row=5)
        flag = True
    else:
        if combo3.get() == "Генеральное соглашение на ВРА":
            table = 'RU_LC'
        elif combo3.get() == "Покрытый ВРА":
            table = 'cove_LC'
        else:
            table = 'int_LC'
        lbl_none2.grid_remove()
        product = str(combo3.get())

    date = str(txt.get())
    if len(date) == 0:
        lbl_none3.grid(column=1, row=7)
        flag = True
    else:
        try: # проверка на соответствие дате
            date = datetime.datetime.strptime(date, '%Y-%m-%d').date()
            lbl_none3.grid_remove()
        except:
            lbl_none3.grid(column=1, row=7)
            flag = True

    if var.get() == 0:
        currency = "RUR"
    elif var.get() == 1:
        currency = "EUR"
    else:
        currency = "USD"
        
    if  currency == "RUR" and table=='int_LC':
        lbl_none8.grid(column=1, row=8)
        flag=True
    else:
        lbl_none8.grid_remove()
        
    client_name = str(txt1.get())
    if len(client_name) == 0:
        lbl_none4.grid(column=1, row=13)
        flag = True
    else:
        lbl_none4.grid_remove()

    # проверка на то, что данные есть и они только цифры
    if len(txt2.get()) == 0 or txt2.get().isalpha():
        lbl_none5.grid(column=1, row=15)
        flag = True
    else:
        cash = int(txt2.get())
        lbl_none5.grid_remove()

    manager_name = str(combo6.get())
    if len(manager_name) == 0:
        lbl_none6.grid(column=1, row=17)
        flag = True
    else:
        lbl_none6.grid_remove()

    if flag == False:
        l_sql_exec = l_conn_ora.connect()
        sql = "insert into %s values ('%s',to_date ('%s','yyyy-mm-dd'),'%s','%s','%s',%d,'%s')" % \
              (table, territory, date, product, client_name, currency, cash, manager_name)

        l_sql_exec.execute(sql)
        print('Данные вставлены')
        lbl_none7 = Label(window, text=" Данные  занесены   ", font=("Calibri", 10, "bold"))
        lbl_none7.grid(column=1, row=18)
    else:
        lbl_none7 = Label(window, text="Данные не занесены", font=("Calibri", 10, "bold"))
        lbl_none7.grid(column=1, row=18)

# Интерфейс
window = Tk()
window.geometry('1000x500')
window.title("Рейтинг результативности менеджеров")

lbl = Label(window, text="Для подсчета рейтинга, необходимо заполнить поля:", font=("Calibri", 12, "bold"))
lbl.grid(column=0, row=1, sticky=E + W)

lbl2 = Label(window, text="Выберите территорию", font=("Calibri", 10, "bold"))
lbl2.grid(column=0, row=2)
combo2 = Combobox(window, width=30)
combo2['values'] = ("Головной офис", "Екатеринбург", "Челябинск", "Башкирия",
                    "Курган", "Тюмень", "Новый Уренгой")
combo2.grid(column=0, row=3)

lbl3 = Label(window, text="Выберите сделку", font=("Calibri", 10, "bold"))
lbl3.grid(column=0, row=4)
combo3 = Combobox(window, width=30)
combo3['values'] = ("Генеральное соглашение на ВРА", "Мультивалютное генеральное соглашение",
                    "Покрытый ВРА", "Покрытый ИА за счет с/с", "Экспортный аккредитив")
combo3.grid(column=0, row=5)

lbl7 = Label(window, text="Введите дату сделки в формате YYYY-MM-DD", font=("Calibri", 10, "bold"))
lbl7.grid(column=0, row=6)
txt = Entry(window, width=33)
txt.grid(column=0, row=7)

lbl8 = Label(window, text="Выберите валюту", font=("Calibri", 10, "bold"))
var = IntVar()
var.set(0) # установка по-умолчанию
rad1 = Radiobutton(window, text='RUR', value=0, variable=var)
rad2 = Radiobutton(window, text='EUR', value=1, variable=var)
rad3 = Radiobutton(window, text='USD', value=2, variable=var)
lbl8.grid(column=0, row=8)
rad1.grid(column=0, row=9)
rad2.grid(column=0, row=10)
rad3.grid(column=0, row=11)

lbl4 = Label(window, text="Введите наименование клиента", font=("Calibri", 10, "bold"))
lbl4.grid(column=0, row=12)
txt1 = Entry(window, width=33)
txt1.grid(column=0, row=13)

lbl5 = Label(window, text="Введите объем дохода в тыс. руб", font=("Calibri", 10, "bold"))
lbl5.grid(column=0, row=14)
txt2 = Entry(window, width=33)
txt2.grid(column=0, row=15)

lbl6 = Label(window, text="Выберите продуктового менеджера", font=("Calibri", 10, "bold"))
lbl6.grid(column=0, row=16)
combo6 = Combobox(window, width=30)
combo6['values'] = ("Пупкин И.А.", "Сергеев А.А",
                    "Покрышкин Е.Н", "Лимонова С.Р", "Крутой И.О")
combo6.grid(column=0, row=17)

# Текст при ошибочном вводе
lbl_none1 = Label(window, text="Выберите территорию", font=("Calibri", 10, "bold"))
lbl_none2 = Label(window, text="Выберите сделку", font=("Calibri", 10, "bold"))
lbl_none3 = Label(window, text="Введите корректную дату", font=("Calibri", 10, "bold"))
lbl_none4 = Label(window, text="Введите наименование клиента", font=("Calibri", 10, "bold"))
lbl_none5 = Label(window, text="Введите общий доход", font=("Calibri", 10, "bold"))
lbl_none6 = Label(window, text="Выберите менеджера", font=("Calibri", 10, "bold"))
lbl_none8 = Label(window, text="Выберите корректную валюту", font=("Calibri", 10, "bold"))

# Текст рейтинга
lbl4 = Label(window, text="Рейтинг менеджеров", font=("Calibri", 10, "bold"))
lbl4.grid(column=0, row=22)
ent = Text(window, width=33)
ent.grid(row=24, column=0, rowspan=10, sticky=E + W + N)

# Подключение к БД
l_user = ''
l_pass = ''
l_tns = ora.makedsn('13.95.167.129', 1521, service_name='pdb1')
l_conn_ora = ''
l_conn_ora = adb.create_engine(r'oracle://{p_user}:{p_pass}@{p_tns}'
    .format(p_user=l_user, p_pass=l_pass, p_tns=l_tns))
# print(l_conn_ora)

# Кнопки
btn = Button(window, text="Отправить", width=20, command=lambda: insertbd(l_conn_ora))
btn.grid(column=0, row=18, padx=3, pady=3) # padx - рамка
btn1 = Button(window, text="Посмотреть общий рейтинг", width=25, command=lambda: rating(l_conn_ora))
btn1.grid(column=0, row=19, padx=3, pady=3)
btn = Button(window, text="Посмотреть итоговый рейтинг", width=25, command=lambda: rating2(l_conn_ora))
btn.grid(column=1, row=22, padx=3, pady=3) # padx - рамка
lbl_cur = Label(window, text="Выберите вид сделки для итогового рейтинга", font=("Calibri", 10, "bold"))
lbl_cur.grid(column=1, row=19)
var2 = IntVar()
var2.set(0) # установка по-умолчанию
rad_3 = Radiobutton(window, text='Внутрироссийские аккредитивы', value=0, variable=var2)
rad_4 = Radiobutton(window, text='Международные аккредитивы', value=1, variable=var2)
rad_3.grid(column=1, row=20)
rad_4.grid(column=1, row=21)

btn2 = Button(window, text="Экспорт в Excel", width=20, command=lambda: export_xls(txt_exl.get()))
btn2.grid(column=2, row=21, padx=3, pady=3)
txt_exl = Entry(window, width=33)
txt_exl.insert(0, 'Введите название файла')
txt_exl.grid(column=3, row=21, padx=3, pady=3)
lbl_exl = Label(window, text="Данные сохранены в excel-файле", font=("Calibri", 10, "bold"))
lbl_exl_2 = Label(window, text="Данные НЕ сохранены в excel-файле", font=("Calibri", 10, "bold"))

window.mainloop()
