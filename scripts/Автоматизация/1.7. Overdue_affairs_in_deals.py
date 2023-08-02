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
#OVERDUE
#---------------------------------------------
driver = webdriver.Chrome(executable_path = f'C:\\Users\\{log}\\scripts\\YandexDriver\\yandexdriver.exe')
x = driver.get('http://msk1-bidb02/reports/report/Отчеты%20КД/Overdue_affairs_in_deals')
time.sleep(5)

ActionChains(driver) \
    .send_keys(Keys.TAB,
               "01.01.2022",
               Keys.TAB * 2,
               f"{d}.06.2023",
               Keys.TAB * 3,
               Keys.ENTER,
               Keys.TAB * 4,
               Keys.SPACE,
               Keys.ENTER,
               Keys.TAB * 5) \
    .perform()

time.sleep(2)

ActionChains(driver) \
    .send_keys(Keys.ENTER,
               Keys.TAB * 5,
               Keys.ENTER) \
    .perform()

time.sleep(300)

ActionChains(driver) \
    .move_by_offset(635, 265).click() \
    .move_by_offset(0, 90).click() \
    .perform()

time.sleep(60)

src_ov = 'C:\\Users\\ADavydovskiy\\Downloads\\Overdue_affairs_in_deals.xlsx'
dest_ov = '\\\\synergy.local\\Documents\\19.Группа мониторинга и сопровождения сделок\\01.Отчеты\\Харитуня\\Шаблоны\\STAGE\\Сделки\\Overdue_affairs_in_deals.xlsx'

shutil.move(src_ov, dest_ov)

print("")
print("Готово!")