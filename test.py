#-*- coding: utf-8 -^-
#-*- encoding: utf8 -*-   # 한글경로 읽고 쓰기 되면 좋겠다
import pyautogui as pag
import time as t
import keyboard
from datetime import *
import calendar
import openpyxl # Openpyxl 라이브러리 불러오기
import xlwings as xw
import pandas as pd
import pyperclip
import ctypes
import pyautogui as pag
import time as t
import keyboard
from datetime import *
import calendar
import numpy as np
import xlwings as xw
user32 = ctypes.windll.user32
from PIL import ImageGrab, Image

# left_top = (74, 85)
# right_bottom = (131, 98)
# # 74, 85, 58, 14
# left_top_x = left_top[0]
# right_bottom_y = right_bottom[1]
# capture_width = right_bottom[0] - left_top[0]
# capture_height = right_bottom[1] - left_top[1]
#
#
# for i in range(0, 12, 1):
#     path = r"C:\work\crm_"+str(i+1)+".png"
#     pag.screenshot(path, region=(left_top[0], left_top[1], capture_width, capture_height))
#     t.sleep(1)
#     pag.click(184, 89)
#     t.sleep(1)

# adm = pag.locateCenterOnScreen('h_adm.PNG', confidence=0.98, region=(515, 504, 69, 30))
# not_adm = pag.locateCenterOnScreen('h_not_adm.PNG', confidence=0.98, region=(515, 504, 69, 30))
# if(adm):
#     print('입원환자')
# elif(not_adm):
#     print('외래환자')
# else:
#     print('에러예요!')


# left_top = (515,504)
# right_bottom = (69, 30)
# 74, 85, 58, 14
# left_top_x = left_top[0]
# right_bottom_y = right_bottom[1]
# capture_width = right_bottom[0] - left_top[0]
# capture_height = right_bottom[1] - left_top[1]


# for i in range(0, 12, 1):
# path = r"C:\work\crm_adm.png"
# pag.screenshot(path, region=(left_top[0], left_top[1], right_bottom[0], right_bottom[1]))
# pag.screenshot(path, region=(left_top[0], left_top[1], capture_width, capture_height))
# t.sleep(1)
# pag.click(184, 89)
# t.sleep(1)



#
# yy=2022
# mm=4
# dd=29
#
# def crm_month_check():
#     if keyboard.is_pressed('END'):
#         return
#
#     crm_1 = pag.locateCenterOnScreen('crm_1.PNG', confidence=0.95, region=(74, 85, 57, 13))
#     crm_2 = pag.locateCenterOnScreen('crm_2.PNG', confidence=0.95, region=(74, 85, 57, 13))
#     crm_3 = pag.locateCenterOnScreen('crm_3.PNG', confidence=0.95, region=(74, 85, 57, 13))
#     crm_4 = pag.locateCenterOnScreen('crm_4.PNG', confidence=0.95, region=(74, 85, 57, 13))
#     crm_5 = pag.locateCenterOnScreen('crm_5.PNG', confidence=0.95, region=(74, 85, 57, 13))
#     crm_6 = pag.locateCenterOnScreen('crm_6.PNG', confidence=0.95, region=(74, 85, 57, 13))
#     crm_7 = pag.locateCenterOnScreen('crm_7.PNG', confidence=0.95, region=(74, 85, 57, 13))
#     crm_8 = pag.locateCenterOnScreen('crm_8.PNG', confidence=0.95, region=(74, 85, 57, 13))
#     crm_9 = pag.locateCenterOnScreen('crm_9.PNG', confidence=0.95, region=(74, 85, 57, 13))
#     crm_10 = pag.locateCenterOnScreen('crm_10.PNG', confidence=0.95, region=(74, 85, 57, 13))
#     crm_11 = pag.locateCenterOnScreen('crm_11.PNG', confidence=0.95, region=(74, 85, 57, 13))
#     crm_12 = pag.locateCenterOnScreen('crm_12.PNG', confidence=0.95, region=(74, 85, 57, 13))
#     if crm_1:
#         return 1
#     elif crm_2:
#         return 2
#     elif crm_3:
#         return 3
#     elif crm_4:
#         return 4
#     elif crm_5:
#         return 5
#     elif crm_6:
#         return 6
#     elif crm_7:
#         return 7
#     elif crm_8:
#         return 8
#     elif crm_9:
#         return 9
#     elif crm_10:
#         return 10
#     elif crm_11:
#         return 11
#     elif crm_12:
#         return 12
#     else:
#         return 0
#
# def crm_click_month(month, check_month):
#     if keyboard.is_pressed('END'):
#         return
#
#     to_go_month = month
#     cur_month = check_month
#
#     for i in range(0, 12, 1):
#         if(to_go_month > cur_month):      # 목표 월이 더 크면?
#             pag.click(184, 89)   # 다음달로 보냅시다
#             t.sleep(1)
#             if(to_go_month == crm_month_check()):
#                 return
#         else:                            # 목표 월이 더 작으면?
#             pag.click(24, 92)    # 이전달로 보냅시다
#             t.sleep(1)
#             if (to_go_month == crm_month_check()):
#                 return
#
# if (mm == crm_month_check()):
#     print('달:', mm, 'check:', crm_month_check())
# elif (mm > crm_month_check()):
#     print('달:', mm, 'check:', crm_month_check())
#     crm_click_month(mm, crm_month_check())
# elif (mm < crm_month_check()):
#     print('달:', mm, 'check:', crm_month_check())
#     crm_click_month(mm, crm_month_check())



#
# def get_date(y, m, d):
#     '''y: year(4 digits)
#      m: month(2 digits)
#      d: day(2 digits'''
#     s = f'{y:04d}-{m:02d}-{d:02d}'
#     return datetime.strptime(s, '%Y-%m-%d')
# def getMonthRange(year, month):
#     date = datetime.(year=year, month=month, day=1).date()
#     monthrange = calendar.monthrange(date.year, date.month)
#     first_day = calendar.monthrange(date.year, date.month)[0]
#     last_day = calendar.monthrange(date.year, date.month)[1]
#
# date = get_date(yy, mm, dd)
# firstday = calendar.monthrange(date.year, date.month)[0]
# print(date, firstday)
# date = get_date(yy, mm - 1, dd)
# lastday = calendar.monthrange(date.year, date.month)[1]
# print(date, lastday)
#
# time1 = datetime(yy,mm,dd)
# print(time1.strftime('%Y-%m-%d'))
# time2 = time1 + timedelta(days=-3)
# print(time2.strftime('%Y-%m-%d'))
# print(time2.year)
# print(time2.month)
# print(time2.day)
# print(type(time2.day))


s5117 = pag.locateCenterOnScreen('s5117.png', confidence=0.92, region=(508, 76, 782, 854))
s5119 = pag.locateCenterOnScreen('s5119.png', confidence=0.92, region=(508, 76, 782, 854))
print('s5117:', s5117)
print('s5119:', s5119)
x = s5117[0] + 495
print('x:', x, '?', s5117[1])