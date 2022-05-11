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


def hiq_ready():
    if keyboard.is_pressed('END'):
        return
    while True:
        if keyboard.is_pressed('END'):
            return

        hiq_jinryo = pag.locateCenterOnScreen('d_hiq_jinryo.PNG', confidence=0.98, region=(75, 53, 98, 30))  # 진료차트 파란불 들어왔나?
        find_hiq = pag.locateCenterOnScreen('d_find_hiq.PNG', confidence=0.98, region=(105, 1038, 717, 42))
        hiq_icon = pag.locateCenterOnScreen('d_hiq_off.PNG', confidence=0.98, region=(105, 1038, 717, 42))
        hiq_finished_list_off = pag.locateCenterOnScreen('d_hiq_finished_list_off.PNG', confidence=0.98, region=(1294, 333, 493, 30))
        hiq_finished_list_on = pag.locateCenterOnScreen('d_hiq_finished_list_on.PNG', confidence=0.98, region=(1294, 333, 493, 30))
        hiq_jinryo_menu = pag.locateCenterOnScreen('d_hiq_jinryo_menu.PNG', confidence=0.9, region=(8, 24, 218, 28))  # 필요없을 것 같지만 '진료차트' 메뉴 선택

        if hiq_jinryo and hiq_finished_list_on:
            return print('화면 준비 완료')

        if hiq_finished_list_off and not hiq_finished_list_on:    # 완료() 버튼에 파란불 아니고 회색불 들어와있다?
            pag.click(hiq_finished_list_off)
            t.sleep(0.5)
            return print('화면 준비됨')

        if not hiq_jinryo_menu:    # 화면에 접수등록회색바탕메뉴가 없다면?
            hiq_soonap = pag.locateCenterOnScreen('d_hiq_soonap.PNG', confidence=0.98, region=(305, 111, 1106, 719))
            if hiq_soonap:  # 메인메뉴에서 '접수.수납'버튼 클릭
                pag.click(hiq_soonap)  # 수납/접수창 띄워짐.. 좀 오래걸령
                print('수납/접수창 띄우는 중')
                for i in range(0, 11, 1):
                    t.sleep(1)
                    print(i+1)
                    hiq_magam = pag.locateCenterOnScreen('d_hiq_magam.PNG', confidence=0.98)
                    if (hiq_magam):  # 접수등록 창에 마감관리버튼 보인다!
                        print('!')
                        break
            # hiq_jeopsoo = pag.locateCenterOnScreen('d_hiq_jeopsoo.PNG', confidence=0.95)
            # if (hiq_jeopsoo):  # 접수등록탭 꺼져있음
            #     pag.click(hiq_jeopsoo)  # 여기까지 하면 접수등록탭 파란불 들어온 화면이 되지요!
            #     t.sleep(3)
            else:
                print('수납접수가 엄서, 작업표시줄에 하이큐 아이콘 클릭해')
                pag.click(find_hiq)
                t.sleep(1)
        else:
            pag.click(hiq_jinryo_menu)
            for i in range(0, 11, 1):
                t.sleep(1)
                print(i + 1)
                hiq_jinryo = pag.locateCenterOnScreen('d_hiq_jinryo.PNG', confidence=0.98, region=(75, 53, 98, 30))  # 진료차트 파란불 들어왔나?
                if hiq_jinryo:
                    print('!')
                    break

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





hiq_ready()