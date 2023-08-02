from datetime import datetime
from __future__ import print_function
import os
import sys
import time
import keyboard

#Это запись в текстовый файл
#-----------------------------------------
#rep = open(r"C:\Users\ADavydovskiy\Desktop\report.txt", "w")
#print("test", file = rep)
#rep.close()
#-----------------------------------------

darkblue = "\033[1;34m"
red = "\033[1;31m"
end = "\033[0;0m"

program_starts_1 = '08:00:00'
program_starts_2 = ['11:40:00', '14:40:00', '17:40:00']
program_starts_week = [0, 1, 2, 3, 4]

puk = 1
path = r"\\synergy.local\Documents\19.Группа мониторинга и сопровождения сделок\01.Отчеты\База ЧВ_бд18 0,25.xlsx"
print(f"|------------------------------------{datetime.now().strftime('%d.%m.%Y')}------------------------------------|")
print("")
print(" 1) Протягивание лидов: " + darkblue + "(alt + shift + 1)" + end)
print(" 2) Обновление БАЗЫ 2016: " + darkblue + "(alt + shift + 2)" + end)
print(" 3) Товары: " + darkblue + "(alt + shift + 3)" + end)
print(" 4) Маркетинг: " + darkblue + "(alt + shift + 4)" + end)
print(" 5) Выгрузка запланированных: " + darkblue + "(alt + shift + 5)"+ end)
print(" 6) Любой скрипт через input(): " + darkblue + "(alt + shift + 6)" + end)
print("")

while(True):
#-----------------------------------------
    try:
        ti_m = os.path.getmtime(path)
    except:
        time.sleep(10)
    m_ti = time.ctime(ti_m)
    t_obj = time.strptime(m_ti)
    T_stamp = time.strftime("%Y-%m-%d", t_obj)
#-----------------------------------------
    if datetime.now().strftime('%H:%M:%S') == program_starts_1 and datetime.today().weekday() in program_starts_week:
        try:
            execfile(r"\\synergy.local\Documents\19.Группа мониторинга и сопровождения сделок\01.Отчеты\Автоматизация\1.2. Конвертации в 1.py")
        except Exception as e: 
            print(red + f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | Ошибка: " + end, e)
            print("")
        #-----------------------------------------
        #Распределение лидов на МП (Сейчас не используется)
        #try:
        #    execfile(r"\\synergy.local\Documents\19.Группа мониторинга и сопровождения сделок\01.Отчеты\Автоматизация\1.3. Распределение на МП.py")
        #except Exception as e: 
        #    print(red + f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | Ошибка: " + end + e)
        #    print("")
        #-----------------------------------------
        try:
            execfile(r"\\synergy.local\Documents\19.Группа мониторинга и сопровождения сделок\01.Отчеты\Автоматизация\1.4. Запланированные оплаты.py")
        except Exception as e: 
            print(red + f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | Ошибка: " + end, e)
            print("")
        try:
            execfile(r"\\synergy.local\Documents\19.Группа мониторинга и сопровождения сделок\01.Отчеты\Автоматизация\1.5. Отчёт по зачислению.py")
        except Exception as e: 
            print(red + f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | Ошибка: " + end, e)
            print("")
        try:
            execfile(r"\\synergy.local\Documents\19.Группа мониторинга и сопровождения сделок\01.Отчеты\Автоматизация\Отчёт по ЛВ.py")
        except Exception as e: 
            print(red + f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | Ошибка: " + end, e)
            print("")        
#-----------------------------------------
        try:
            execfile(r"\\synergy.local\Documents\19.Группа мониторинга и сопровождения сделок\01.Отчеты\Автоматизация\Запланированные.py")
        except:
            print(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | Попытка номер 2...")
            print("")
            try:
                execfile(r"\\synergy.local\Documents\19.Группа мониторинга и сопровождения сделок\01.Отчеты\Автоматизация\Запланированные.py")
            except:
                print(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | Попытка номер 3...")
                print("")
                try:
                    execfile(r"\\synergy.local\Documents\19.Группа мониторинга и сопровождения сделок\01.Отчеты\Автоматизация\Запланированные.py")
                except:
                    print(red + f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | Увы, обновить запланированные не удалось..." + end)
                    print("")
#-----------------------------------------    
    elif datetime.now().strftime('%H:%M:%S') in program_starts_2:
        try:
            execfile(r"\\synergy.local\Documents\19.Группа мониторинга и сопровождения сделок\01.Отчеты\Автоматизация\Планирование на день РОП.py")
        except:
            print(red + f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | Увы, выгрузить планирование не удалось..." + end)
            print("")
#-----------------------------------------    
    elif puk != 1 and T_stamp == datetime.today().strftime("%Y-%m-%d"):
        try:
            execfile(r"\\synergy.local\Documents\19.Группа мониторинга и сопровождения сделок\01.Отчеты\Автоматизация\2.2. ДОП.py")
        except Exception as e: 
            print(red + f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | Ошибка: " + end, e)
            print("")
        try:
            execfile(r"\\synergy.local\Documents\19.Группа мониторинга и сопровождения сделок\01.Отчеты\Автоматизация\2.3. Отчёты Баз Карточки.py")
        except Exception as e: 
            print(red + f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | Ошибка: " + end, e)
            print("")
        puk = 1
#-----------------------------------------    
    elif keyboard.is_pressed('alt+shift+1'):
        try:
            execfile(r"\\synergy.local\Documents\19.Группа мониторинга и сопровождения сделок\01.Отчеты\Автоматизация\1.1. Протягивание лидов.py")
        except Exception as e: 
            print(red + f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | Ошибка: " + end, e)
            print("")
#-----------------------------------------    
    elif keyboard.is_pressed('alt+shift+2'):
        try:
            execfile(r"\\synergy.local\Documents\19.Группа мониторинга и сопровождения сделок\01.Отчеты\Автоматизация\2.1. База 2016.py")
        except Exception as e: 
            print(red + f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | Ошибка: " + end, e)
            print("")
#-----------------------------------------    
    elif keyboard.is_pressed('alt+shift+3'):
        try:
            execfile(r"\\synergy.local\Documents\19.Группа мониторинга и сопровождения сделок\01.Отчеты\Автоматизация\Товары.py")
        except Exception as e: 
            print(red + f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | Ошибка: " + end, e)
            print("")
#-----------------------------------------    
    elif keyboard.is_pressed('alt+shift+4'):
        try:
            execfile(r"\\synergy.local\Documents\19.Группа мониторинга и сопровождения сделок\01.Отчеты\Автоматизация\Маркетинг.py")
        except Exception as e: 
            print(red + f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | Ошибка: " + end, e)
            print("")
#-----------------------------------------    
    elif keyboard.is_pressed('alt+shift+5'):
        try:
            execfile(r"\\synergy.local\Documents\19.Группа мониторинга и сопровождения сделок\01.Отчеты\Автоматизация\Запланированные.py")
        except Exception as e: 
            print(red + f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | Увы, обновить запланированные не удалось..." + end)
            print("")
#-----------------------------------------    
    elif keyboard.is_pressed('alt+shift+6'):
        try:
            x=input()
            execfile(r"\\synergy.local\Documents\19.Группа мониторинга и сопровождения сделок\01.Отчеты\Автоматизация\{}".format(x))
        except Exception as e: 
            print(red + f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | Ошибка: " + end, e)
            print("")
#-----------------------------------------            
    elif datetime.now().strftime('%H:%M:%S') == '00:00:05':
        puk = 0
        time.sleep(2)
        print(f"|------------------------------------{datetime.now().strftime('%d.%m.%Y')}------------------------------------|")
        print("")
        print(" 1) Протягивание лидов: " + darkblue + "(alt + shift + 1)" + end)
        print(" 2) Обновление БАЗЫ 2016: " + darkblue + "(alt + shift + 2)" + end)
        print(" 3) Товары: " + darkblue + "(alt + shift + 3)" + end)
        print(" 4) Маркетинг: " + darkblue + "(alt + shift + 4)" + end)
        print(" 5) Выгрузка запланированных: " + darkblue + "(alt + shift + 5)"+ end)
        print(" 6) Любой скрипт через input(): " + darkblue + "(alt + shift + 6)" + end)
        print("")