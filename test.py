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
#
#
# s5117 = pag.locateCenterOnScreen('s5117.png', confidence=0.92, region=(508, 76, 782, 854))
# s5119 = pag.locateCenterOnScreen('s5119.png', confidence=0.92, region=(508, 76, 782, 854))
# print('s5117:', s5117)
# print('s5119:', s5119)
# x = s5117[0] + 495
# print('x:', x, '?', s5117[1])


# path = r"C:\work\crm_adm.png"
# pag.screenshot(path, region=(left_top[0], left_top[1], right_bottom[0], right_bottom[1]))
# pag.screenshot(path, region=(left_top[0], left_top[1], capture_width, capture_height))
# t.sleep(1)
# pag.click(184, 89)
# t.sleep(1)

# for i in range(0, 6, 1):
#     path = r"C:\work\hiq_year_" + str(i + 1) + ".png"
#     pag.screenshot(path, region=(1732, 91, 33,11))
#     t.sleep(1)
#     pag.click(1875,96) # 다음 해로 넘겨!
#     t.sleep(0.1)
#     pag.moveTo(1838,319)
#     t.sleep(0.5)

# for i in range(0, 12, 1):
#     path = r"C:\work\hiq_month_" + str(i + 1) + ".png"
#     pag.screenshot(path, region=(1769, 91, 33, 11))
#     t.sleep(1)
#     pag.click(1840, 96)  # 다음 달로 넘겨!
#     t.sleep(0.1)
#     pag.moveTo(1838, 319)
#     t.sleep(0.5)
#
# hiq_month_1 = pag.locateCenterOnScreen('hiq_month_1.PNG', confidence=0.95, region=(1769, 91, 33, 11))
# hiq_month_2 = pag.locateCenterOnScreen('hiq_month_2.PNG', confidence=0.95, region=(1769, 91, 33, 11))
# hiq_month_3 = pag.locateCenterOnScreen('hiq_month_3.PNG', confidence=0.95, region=(1769, 91, 33, 11))
# hiq_month_4 = pag.locateCenterOnScreen('hiq_month_4.PNG', confidence=0.95, region=(1769, 91, 33, 11))
# hiq_month_5 = pag.locateCenterOnScreen('hiq_month_5.PNG', confidence=0.95, region=(1769, 91, 33, 11))
# hiq_month_6 = pag.locateCenterOnScreen('hiq_month_6.PNG', confidence=0.95, region=(1769, 91, 33, 11))
# hiq_month_7 = pag.locateCenterOnScreen('hiq_month_7.PNG', confidence=0.95, region=(1769, 91, 33, 11))
# hiq_month_8 = pag.locateCenterOnScreen('hiq_month_8.PNG', confidence=0.95, region=(1769, 91, 33, 11))
# hiq_month_9 = pag.locateCenterOnScreen('hiq_month_9.PNG', confidence=0.95, region=(1769, 91, 33, 11))
# hiq_month_10 = pag.locateCenterOnScreen('hiq_month_10.PNG', confidence=0.95, region=(1769, 91, 33, 11))
# hiq_month_11 = pag.locateCenterOnScreen('hiq_month_11.PNG', confidence=0.95, region=(1769, 91, 33, 11))
# hiq_month_12 = pag.locateCenterOnScreen('hiq_month_12.PNG', confidence=0.95, region=(1769, 91, 33, 11))
# if hiq_month_1:
#     print(' 1')
# elif hiq_month_2:
#     print(' 2')
# elif hiq_month_3:
#     print(' 3')
# elif hiq_month_4:
#     print(' 4')
# elif hiq_month_5:
#     print(' 5')
# elif hiq_month_6:
#     print(' 6')
# elif hiq_month_7:
#     print(' 7')
# elif hiq_month_8:
#     print(' 8')
# elif hiq_month_9:
#     print(' 9')
# elif hiq_month_10:
#     print(' 10')
# elif hiq_month_11:
#     print(' 11')
# elif hiq_month_12:
#     print(' 12')
# else:
#     print(' 0')


# def hiq_ready():
#     if keyboard.is_pressed('END'):
#         return
#     while True:
#         if keyboard.is_pressed('END'):
#             return
#
#         hiq_jinryo = pag.locateCenterOnScreen('d_hiq_jinryo.PNG', confidence=0.98, region=(75, 53, 98, 30))  # 진료차트 파란불 들어왔나?
#         find_hiq = pag.locateCenterOnScreen('d_find_hiq.PNG', confidence=0.98, region=(105, 1038, 717, 42))
#         hiq_icon = pag.locateCenterOnScreen('d_hiq_off.PNG', confidence=0.98, region=(105, 1038, 717, 42))
#         hiq_finished_list_off = pag.locateCenterOnScreen('d_hiq_finished_list_off.PNG', confidence=0.98, region=(1294, 333, 493, 30))
#         hiq_finished_list_on = pag.locateCenterOnScreen('d_hiq_finished_list_on.PNG', confidence=0.98, region=(1294, 333, 493, 30))
#         hiq_jinryo_menu = pag.locateCenterOnScreen('d_hiq_jinryo_menu.PNG', confidence=0.9, region=(8, 24, 218, 28))  # 필요없을 것 같지만 '진료차트' 메뉴 선택
#
#         if hiq_jinryo and hiq_finished_list_on:
#             return print('화면 준비 완료')
#
#         if hiq_finished_list_off and not hiq_finished_list_on:    # 완료() 버튼에 파란불 아니고 회색불 들어와있다?
#             pag.click(hiq_finished_list_off)
#             t.sleep(0.5)
#             return print('화면 준비됨')
#
#         if not hiq_jinryo_menu:    # 화면에 접수등록회색바탕메뉴가 없다면?
#             hiq_soonap = pag.locateCenterOnScreen('d_hiq_soonap.PNG', confidence=0.98, region=(305, 111, 1106, 719))
#             if hiq_soonap:  # 메인메뉴에서 '접수.수납'버튼 클릭
#                 pag.click(hiq_soonap)  # 수납/접수창 띄워짐.. 좀 오래걸령
#                 print('수납/접수창 띄우는 중')
#                 for i in range(0, 11, 1):
#                     t.sleep(1)
#                     print(i+1)
#                     hiq_magam = pag.locateCenterOnScreen('d_hiq_magam.PNG', confidence=0.98)
#                     if (hiq_magam):  # 접수등록 창에 마감관리버튼 보인다!
#                         print('!')
#                         break
#             # hiq_jeopsoo = pag.locateCenterOnScreen('d_hiq_jeopsoo.PNG', confidence=0.95)
#             # if (hiq_jeopsoo):  # 접수등록탭 꺼져있음
#             #     pag.click(hiq_jeopsoo)  # 여기까지 하면 접수등록탭 파란불 들어온 화면이 되지요!
#             #     t.sleep(3)
#             else:
#                 print('수납접수가 엄서, 작업표시줄에 하이큐 아이콘 클릭해')
#                 pag.click(find_hiq)
#                 t.sleep(1)
#         else:
#             pag.click(hiq_jinryo_menu)
#             for i in range(0, 11, 1):
#                 t.sleep(1)
#                 print(i + 1)
#                 hiq_jinryo = pag.locateCenterOnScreen('d_hiq_jinryo.PNG', confidence=0.98, region=(75, 53, 98, 30))  # 진료차트 파란불 들어왔나?
#                 if hiq_jinryo:
#                     print('!')
#                     break

        # elif (find_hiq):  # 메인메뉴만 켜져있음(아이콘있음)
        #     pag.click(find_hiq)
        #     print('hiq 메인메뉴만 켜져있네요')
        #     t.sleep(1)
        # elif (find_hiq):   # 메인메뉴만 켜져있음(아이콘있음)
        #     pag.click(find_hiq)
        #     t.sleep(1)
        #     hiq_soonap = pag.locateCenterOnScreen('d_hiq_soonap.PNG', confidence=0.9, region=(305, 111, 1106, 719))
        #     if (hiq_soonap):      # 메인메뉴에서 수납버튼 클릭
        #         pag.click(hiq_soonap)     # 수납/접수창 띄워짐
        #         t.sleep(12)
        #         if (hiq_jinryo_menu):
        #             pag.click(hiq_jinryo_menu)
        #             t.sleep(1)   # 여기까지 하면 진료차트창 띄워졌다
        #     else:
        #         print('왜 수납아이콘 없지')
        #         return False
        # else:
        #     print('차트프로그램 켜야해')
        #     return False


#
# def Updown(updown):
#     if keyboard.is_pressed('end'):
#         print('end 눌러 종료합니다.')
#         return
#     if updown == 'down':
#         pag.click(1332, 905)
#         t.sleep(0.5)
#         pag.scroll(-2200) # 2200이면 18명정도 내리는듯
#         t.sleep(1)
# Updown('down')
#
# def find_num(where):
#     if keyboard.is_pressed('END'):
#         return
#     if (where == 'top'):
#         x = 1282
#         y = 427
#     if (where == 'bottom'):
#         x = 1279
#         y = 887
#     cond_num_1 = pag.locateCenterOnScreen('1.PNG', confidence=0.96, region=(x, y, 68+50, 34+30))
#     cond_num_2 = pag.locateCenterOnScreen('2.PNG', confidence=0.96, region=(x, y, 68+50, 34+30))
#     cond_num_3 = pag.locateCenterOnScreen('3.PNG', confidence=0.96, region=(x, y, 68+50, 34+30))
#     cond_num_4 = pag.locateCenterOnScreen('4.PNG', confidence=0.96, region=(x, y, 68+50, 34+30))
#     cond_num_5 = pag.locateCenterOnScreen('5.PNG', confidence=0.96, region=(x, y, 68+50, 34+30))
#     cond_num_6 = pag.locateCenterOnScreen('6.PNG', confidence=0.96, region=(x, y, 68+50, 34+30))
#     cond_num_7 = pag.locateCenterOnScreen('7.PNG', confidence=0.96, region=(x, y, 68+50, 34+30))
#     cond_num_8 = pag.locateCenterOnScreen('8.PNG', confidence=0.96, region=(x, y, 68+50, 34+30))
#     cond_num_9 = pag.locateCenterOnScreen('9.PNG', confidence=0.96, region=(x, y, 68+50, 34+30))
#     cond_num_10 = pag.locateCenterOnScreen('10.PNG', confidence=0.96, region=(x, y, 68+50, 34+30))
#     cond_num_11 = pag.locateCenterOnScreen('11.PNG', confidence=0.96, region=(x, y, 68+50, 34+30))
#     cond_num_11_1 = pag.locateCenterOnScreen('11-1.PNG', confidence=0.96, region=(x, y, 68+50, 34+30))
#     cond_num_12 = pag.locateCenterOnScreen('12.PNG', confidence=0.96, region=(x, y, 68+50, 34+30))
#     cond_num_13 = pag.locateCenterOnScreen('13.PNG', confidence=0.96, region=(x, y, 68+50, 34+30))
#     cond_num_14 = pag.locateCenterOnScreen('14.PNG', confidence=0.96, region=(x, y, 68+50, 34+30))
#     cond_num_15 = pag.locateCenterOnScreen('15.PNG', confidence=0.96, region=(x, y, 68+50, 34+30))
#     cond_num_16 = pag.locateCenterOnScreen('16.PNG', confidence=0.96, region=(x, y, 68+50, 34+30))
#     cond_num_17 = pag.locateCenterOnScreen('17.PNG', confidence=0.96, region=(x, y, 68+50, 34+30))
#     cond_num_18 = pag.locateCenterOnScreen('18.PNG', confidence=0.96, region=(x, y, 68+50, 34+30))
#     cond_num_19 = pag.locateCenterOnScreen('19.PNG', confidence=0.96, region=(x, y, 68+50, 34+30))
#     cond_num_20 = pag.locateCenterOnScreen('20.PNG', confidence=0.96, region=(x, y, 68+50, 34+30))
#     cond_num_21 = pag.locateCenterOnScreen('21.PNG', confidence=0.96, region=(x, y, 68+50, 34+30))
#     cond_num_22 = pag.locateCenterOnScreen('22.PNG', confidence=0.96, region=(x, y, 68+50, 34+30))
#     cond_num_23 = pag.locateCenterOnScreen('23.PNG', confidence=0.96, region=(x, y, 68+50, 34+30))
#     cond_num_24 = pag.locateCenterOnScreen('24.PNG', confidence=0.96, region=(x, y, 68+50, 34+30))
#     cond_num_25 = pag.locateCenterOnScreen('25.PNG', confidence=0.96, region=(x, y, 68+50, 34+30))
#     cond_num_26 = pag.locateCenterOnScreen('26.PNG', confidence=0.96, region=(x, y, 68+50, 34+30))
#     cond_num_27 = pag.locateCenterOnScreen('27.PNG', confidence=0.96, region=(x, y, 68+50, 34+30))
#     cond_num_28 = pag.locateCenterOnScreen('28.PNG', confidence=0.96, region=(x, y, 68+50, 34+30))
#     cond_num_29 = pag.locateCenterOnScreen('29.PNG', confidence=0.96, region=(x, y, 68+50, 34+30))
#     cond_num_30 = pag.locateCenterOnScreen('30.PNG', confidence=0.96, region=(x, y, 68+50, 34+30))
#     cond_num_31 = pag.locateCenterOnScreen('31.PNG', confidence=0.96, region=(x, y, 68+50, 34+30))
#     cond_num_32 = pag.locateCenterOnScreen('32.PNG', confidence=0.96, region=(x, y, 68+50, 34+30))
#     cond_num_33 = pag.locateCenterOnScreen('33.PNG', confidence=0.96, region=(x, y, 68+50, 34+30))
#     cond_num_34 = pag.locateCenterOnScreen('34.PNG', confidence=0.96, region=(x, y, 68+50, 34+30))
#     cond_num_35 = pag.locateCenterOnScreen('35.PNG', confidence=0.96, region=(x, y, 68+50, 34+30))
#     cond_num_36 = pag.locateCenterOnScreen('36.PNG', confidence=0.96, region=(x, y, 68+50, 34+30))
#     cond_num_37 = pag.locateCenterOnScreen('37.PNG', confidence=0.96, region=(x, y, 68+50, 34+30))
#     cond_num_38 = pag.locateCenterOnScreen('38.PNG', confidence=0.96, region=(x, y, 68+50, 34+30))
#     cond_num_39 = pag.locateCenterOnScreen('39.PNG', confidence=0.96, region=(x, y, 68+50, 34+30))
#     cond_num_40 = pag.locateCenterOnScreen('40.PNG', confidence=0.96, region=(x, y, 68+50, 34+30))
#     cond_num_41 = pag.locateCenterOnScreen('41.PNG', confidence=0.96, region=(x, y, 68+50, 34+30))
#     cond_num_42 = pag.locateCenterOnScreen('42.PNG', confidence=0.96, region=(x, y, 68+50, 34+30))
#     cond_num_43 = pag.locateCenterOnScreen('43.PNG', confidence=0.96, region=(x, y, 68+50, 34+30))
#     cond_num_44 = pag.locateCenterOnScreen('44.PNG', confidence=0.96, region=(x, y, 68+50, 34+30))
#     cond_num_45 = pag.locateCenterOnScreen('45.PNG', confidence=0.96, region=(x, y, 68+50, 34+30))
#     cond_num_46 = pag.locateCenterOnScreen('46.PNG', confidence=0.96, region=(x, y, 68+50, 34+30))
#     cond_num_47 = pag.locateCenterOnScreen('47.PNG', confidence=0.96, region=(x, y, 68+50, 34+30))
#     cond_num_48 = pag.locateCenterOnScreen('48.PNG', confidence=0.96, region=(x, y, 68+50, 34+30))
#     cond_num_49 = pag.locateCenterOnScreen('49.PNG', confidence=0.96, region=(x, y, 68+50, 34+30))
#     cond_num_50 = pag.locateCenterOnScreen('50.PNG', confidence=0.96, region=(x, y, 68+50, 34+30))
#     cond_num_51 = pag.locateCenterOnScreen('51.PNG', confidence=0.96, region=(x, y, 68+50, 34+30))
#     cond_num_52 = pag.locateCenterOnScreen('52.PNG', confidence=0.96, region=(x, y, 68+50, 34+30))
#     cond_num_53 = pag.locateCenterOnScreen('53.PNG', confidence=0.96, region=(x, y, 68+50, 34+30))
#     cond_num_54 = pag.locateCenterOnScreen('54.PNG', confidence=0.96, region=(x, y, 68+50, 34+30))
#     cond_num_55 = pag.locateCenterOnScreen('55.PNG', confidence=0.96, region=(x, y, 68+50, 34+30))
#     # cond_num_56 = pag.locateCenterOnScreen('56.PNG', confidence=0.92, region=(x, y, 68, 34))
#     # cond_num_57 = pag.locateCenterOnScreen('57.PNG', confidence=0.92, region=(x, y, 68, 34))
#     # cond_num_58 = pag.locateCenterOnScreen('58.PNG', confidence=0.92, region=(x, y, 68, 34))
#     # cond_num_59 = pag.locateCenterOnScreen('59.PNG', confidence=0.92, region=(x, y, 68, 34))
#     # cond_num_60 = pag.locateCenterOnScreen('60.PNG', confidence=0.92, region=(x, y, 68, 34))
#     # cond_num_61 = pag.locateCenterOnScreen('61.PNG', confidence=0.92, region=(x, y, 68, 34))
#     # cond_num_62 = pag.locateCenterOnScreen('62.PNG', confidence=0.92, region=(x, y, 68, 34))
#     # cond_num_63 = pag.locateCenterOnScreen('63.PNG', confidence=0.92, region=(x, y, 68, 34))
#     # cond_num_64 = pag.locateCenterOnScreen('64.PNG', confidence=0.92, region=(x, y, 68, 34))
#     # cond_num_65 = pag.locateCenterOnScreen('65.PNG', confidence=0.92, region=(x, y, 68, 34))
#     # cond_num_66 = pag.locateCenterOnScreen('66.PNG', confidence=0.92, region=(x, y, 68, 34))
#     # cond_num_67 = pag.locateCenterOnScreen('67.PNG', confidence=0.92, region=(x, y, 68, 34))
#     # cond_num_68 = pag.locateCenterOnScreen('68.PNG', confidence=0.92, region=(x, y, 68, 34))
#     # cond_num_69 = pag.locateCenterOnScreen('69.PNG', confidence=0.92, region=(x, y, 68, 34))
#     # cond_num_70 = pag.locateCenterOnScreen('70.PNG', confidence=0.92, region=(x, y, 68, 34))
#     # cond_num_71 = pag.locateCenterOnScreen('71.PNG', confidence=0.92, region=(x, y, 68, 34))
#     # cond_num_72 = pag.locateCenterOnScreen('72.PNG', confidence=0.92, region=(x, y, 68, 34))
#     # cond_num_73 = pag.locateCenterOnScreen('73.PNG', confidence=0.92, region=(x, y, 68, 34))
#     # cond_num_74 = pag.locateCenterOnScreen('74.PNG', confidence=0.92, region=(x, y, 68, 34))
#     # cond_num_75 = pag.locateCenterOnScreen('75.PNG', confidence=0.92, region=(x, y, 68, 34))
#     # cond_num_76 = pag.locateCenterOnScreen('76.PNG', confidence=0.92, region=(x, y, 68, 34))
#     # cond_num_77 = pag.locateCenterOnScreen('77.PNG', confidence=0.92, region=(x, y, 68, 34))
#     if (cond_num_1):
#         return 1
#     if (cond_num_2):
#         return 2
#     if (cond_num_3):
#         return 3
#     if (cond_num_4):
#         return 4
#     if (cond_num_5):
#         return 5
#     if (cond_num_6):
#         return 6
#     if (cond_num_7):
#         return 7
#     if (cond_num_8):
#         return 8
#     if (cond_num_9):
#         return 9
#     if (cond_num_10):
#         return 10
#     if (cond_num_11) or (cond_num_11_1):
#         return 11
#     if (cond_num_12):
#         return 12
#     if (cond_num_13):
#         return 13
#     if (cond_num_14):
#         return 14
#     if (cond_num_15):
#         return 15
#     if (cond_num_16):
#         return 16
#     if (cond_num_17):
#         return 17
#     if (cond_num_18):
#         return 18
#     if (cond_num_19):
#         return 19
#     if (cond_num_20):
#         return 20
#     if (cond_num_21):
#         return 21
#     if (cond_num_22):
#         return 22
#     if (cond_num_23):
#         return 23
#     if (cond_num_24):
#         return 24
#     if (cond_num_25):
#         return 25
#     if (cond_num_26):
#         return 26
#     if (cond_num_27):
#         return 27
#     if (cond_num_28):
#         return 28
#     if (cond_num_29):
#         return 29
#     if (cond_num_30):
#         return 30
#     if (cond_num_31):
#         return 31
#     if (cond_num_32):
#         return 32
#     if (cond_num_33):
#         return 33
#     if (cond_num_34):
#         return 34
#     if (cond_num_35):
#         return 35
#     if (cond_num_36):
#         return 36
#     if (cond_num_37):
#         return 37
#     if (cond_num_38):
#         return 38
#     if (cond_num_39):
#         return 39
#     if (cond_num_40):
#         return 40
#     if (cond_num_41):
#         return 41
#     if (cond_num_42):
#         return 42
#     if (cond_num_43):
#         return 43
#     if (cond_num_44):
#         return 44
#     if (cond_num_45):
#         return 45
#     if (cond_num_46):
#         return 46
#     if (cond_num_47):
#         return 47
#     if (cond_num_48):
#         return 48
#     if (cond_num_49):
#         return 49
#     if (cond_num_50):
#         return 50
#     if (cond_num_51):
#         return 51
#     if (cond_num_52):
#         return 52
#     if (cond_num_53):
#         return 53
#     if (cond_num_54):
#         return 54
#     if (cond_num_55):
#         return 55
#     # if (cond_num_56):
#     #     return 56
#     # if (cond_num_57):
#     #     return 57
#     # if (cond_num_58):
#     #     return 58
#     # if (cond_num_59):
#     #     return 59
#     # if (cond_num_60):
#     #     return 60
#     # if (cond_num_61):
#     #     return 61
#     # if (cond_num_62):
#     #     return 62
#     # if (cond_num_63):
#     #     return 63
#     # if (cond_num_64):
#     #     return 64
#     # if (cond_num_65):
#     #     return 65
#     # if (cond_num_66):
#     #     return 66
#     # if (cond_num_67):
#     #     return 67
#     # if (cond_num_68):
#     #     return 68
#     # if (cond_num_69):
#     #     return 69
#     # if (cond_num_70):
#     #     return 70
#     # if (cond_num_71):
#     #     return 71
#     # if (cond_num_72):
#     #     return 72
#     # if (cond_num_73):
#     #     return 73
#     # if (cond_num_74):
#     #     return 74
#     # if (cond_num_75):
#     #     return 75
#     # if (cond_num_76):
#     #     return 76
#     # if (cond_num_77):
#     #     return 77
#     else:
#         return 0
#




def crm_year_check():
    if keyboard.is_pressed('END'):
        return
    crm_year_2015 = pag.locateCenterOnScreen('crm_year_2015.PNG', confidence=0.95, region=(74, 85, 57, 13))
    crm_year_2016 = pag.locateCenterOnScreen('crm_year_2016.PNG', confidence=0.95, region=(74, 85, 57, 13))
    crm_year_2017 = pag.locateCenterOnScreen('crm_year_2017.PNG', confidence=0.95, region=(74, 85, 57, 13))
    crm_year_2018 = pag.locateCenterOnScreen('crm_year_2018.PNG', confidence=0.95, region=(74, 85, 57, 13))
    crm_year_2019 = pag.locateCenterOnScreen('crm_year_2019.PNG', confidence=0.95, region=(74, 85, 57, 13))
    crm_year_2020 = pag.locateCenterOnScreen('crm_year_2020.PNG', confidence=0.95, region=(74, 85, 57, 13))
    crm_year_2021 = pag.locateCenterOnScreen('crm_year_2021.PNG', confidence=0.95, region=(74, 85, 57, 13))
    crm_year_2022 = pag.locateCenterOnScreen('crm_year_2022.PNG', confidence=0.95, region=(74, 85, 57, 13))
    crm_year_2023 = pag.locateCenterOnScreen('crm_year_2023.PNG', confidence=0.95, region=(74, 85, 57, 13))
    crm_year_2024 = pag.locateCenterOnScreen('crm_year_2024.PNG', confidence=0.95, region=(74, 85, 57, 13))
    crm_year_2025 = pag.locateCenterOnScreen('crm_year_2025.PNG', confidence=0.95, region=(74, 85, 57, 13))
    crm_year_2026 = pag.locateCenterOnScreen('crm_year_2026.PNG', confidence=0.95, region=(74, 85, 57, 13))
    if crm_year_2015:
        return 2015
    elif crm_year_2016:
        return 2016
    elif crm_year_2017:
        return 2017
    elif crm_year_2018:
        return 2018
    elif crm_year_2019:
        return 2019
    elif crm_year_2020:
        return 2020
    elif crm_year_2021:
        return 2021
    elif crm_year_2022:
        return 2022
    elif crm_year_2023:
        return 2023
    elif crm_year_2024:
        return 2024
    elif crm_year_2025:
        return 2025
    elif crm_year_2026:
        return 2026
    else:
        return 0

a = crm_year_check()
print(a)



#
# for i in range(0, 145, 1):
#     if (i % 12) == 1:
#         path = r"C:\work\crm_year_" + str(i//12 + 2015) + ".png"
#         pag.screenshot(path, region=(104, 87, 25, 10))
#     t.sleep(0.7)
#     pag.click(184,92) # 다음 달로 넘겨!
#     t.sleep(0.7)
    # pag.moveTo(1838,319)
    # t.sleep(0.5)










#
# a = find_num('top')
# print(a)
#
# path = r"C:\work\ire-1.png"
# pag.screenshot(path, region=(1282, 427, 68+50, 34+30))


