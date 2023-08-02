import win32com.client as win32
import os
from stat import S_IREAD, S_IRGRP, S_IROTH, S_IWUSR
import time

print(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | Обновление отчётов для работы, пожалуйста, подождите...")

xlapp  = win32.DispatchEx('Excel.Application')
xlapp.DisplayAlerts = False
xlapp.Visible = False

l = ['Выгрузка syn vpo', 'Выгрузка сайт  synergy', 
         'Лиды-14 vpo', 'Сделки-14 vpo',
         'Отчет по карточкам Ленд VPO по ОП', 'Отчет по карточкам сайта synergy.ru по ОП',
        'MED','2.1,2 MED']


y = S_IREAD|S_IRGRP|S_IROTH
n = S_IWUSR|S_IREAD

def chek(first):
    chekpoint = first
    for index in l:
        filename = f'\\\\synergy.local\\Documents\\11.Коммерческий департамент\\01. Аналитика КД\\06. Общая аналитика\\Факультеты\\Отчеты для работы\\{index}.xlsx'
        os.chmod(filename, chekpoint)

chek(n)


for i in l:
    xlbook = xlapp.Workbooks.open(fr'\\synergy.local\\Documents\\11.Коммерческий департамент\\01. Аналитика КД\\06. Общая аналитика\\Факультеты\\Отчеты для работы\\{i}.xlsx')
    xlbook.RefreshAll()   
    time.sleep(600)
    xlbook.Save()
    xlbook.Close()
    xlapp.Quit()
    
del xlbook
del xlapp

y = S_IREAD|S_IRGRP|S_IROTH
n = S_IWUSR|S_IREAD

def chek(first):
    chekpoint = first
    for index in l:
        filename = f'\\\\synergy.local\\Documents\\11.Коммерческий департамент\\01. Аналитика КД\\06. Общая аналитика\\Факультеты\\Отчеты для работы\\{index}.xlsx'
        os.chmod(filename, chekpoint)

chek(y)

print(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | Готово!")
print("")