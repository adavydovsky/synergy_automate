import pandas as pd
import numpy as np
import time
from datetime import datetime, timedelta
import re
import os
import shutil

from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains


import warnings
warnings.filterwarnings("ignore")

log = os.getlogin()

day = datetime.today()
d = day.day - 2

#---------------------------------------------
#NOT CORRECT
#---------------------------------------------
driver = webdriver.Chrome(executable_path = f'C:\\Users\\{log}\\scripts\\YandexDriver\\yandexdriver.exe')
x = driver.get('http://msk1-bidb02/reports/report/%D0%9E%D1%82%D1%87%D0%B5%D1%82%D1%8B%20%D0%9A%D0%94/NotCorrectDealsByDepartmentAndDate')
time.sleep(10)
ActionChains(driver) \
    .move_by_offset(333, 158).click() \
    .move_by_offset(-207, 42).click() \
    .send_keys(Keys.ENTER) \
    .move_by_offset(0, 35).click() \
    .send_keys("01.01.2022", Keys.ENTER) \
    .move_by_offset(0, 45).click() \
    .send_keys(f"{d}.06.2023", Keys.ENTER) \
    .perform()

time.sleep(10)

ActionChains(driver) \
    .move_by_offset(1000, -115).click() \
    .perform()

time.sleep(120)

ActionChains(driver) \
    .move_by_offset(-485, 180).click() \
    .move_by_offset(0, 90).click() \
    .perform()

time.sleep(120)

src_nc = 'C:\\Users\\ADavydovskiy\\Downloads\\NotCorrectDealsByDepartmentAndDate.xlsx'
dest_nc = '\\\\synergy.local\\Documents\\19.Группа мониторинга и сопровождения сделок\\01.Отчеты\\Харитуня\\Шаблоны\\STAGE\\Сделки\\NotCorrectDealsByDepartmentAndDate.xlsx'

shutil.move(src_nc, dest_nc)

print("")
print("Ура, Готово!")