import win32com.client as win32
import os
from stat import S_IREAD, S_IRGRP, S_IROTH, S_IWUSR
import time
import keyboard

print(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | Обновление отчётов по госконтрактам, пожалуйста, подождите...")

xlapp  = win32.DispatchEx('Excel.Application')
xlapp.DisplayAlerts = False
xlapp.Visible = False

l = ['1.6 РОМИ Госконтракты SA', '1.6 РОМИ Госконтракты', 
         'Report NEW', 'Report']

for i in l:
    xlbook = xlapp.Workbooks.open(fr'\\synergy.local\\Documents\\11.Коммерческий департамент\\01. Аналитика КД\\06. Общая аналитика\\ГосКонтракты\\{i}.xlsx')
    xlbook.RefreshAll()
    time.sleep(180)
    xlbook.Save()
    xlbook.Close()
    xlapp.Quit()
    
del xlbook
del xlapp

print(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | Готово!")
print("")