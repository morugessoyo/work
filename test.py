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

find_hiq = pag.locateCenterOnScreen('d_find_hiq.PNG', confidence=0.95, region=(105, 1038, 717, 42))
hiq_off = pag.locateCenterOnScreen('d_hiq_off.PNG', confidence=0.95, region=(105, 1038, 717, 42))
finance_on = pag.locateCenterOnScreen('d_finance_on.PNG', confidence=0.95)
if (hiq_off):
    pag.click(hiq_off)
    t.sleep(0.5)
    finance_o = pag.locateCenterOnScreen('d_finance_o.PNG', confidence=0.95)
    t.sleep(1)
    if (finance_o):
        pag.click(finance_o)
        finance_on = pag.locateCenterOnScreen('d_finance_on.PNG', confidence=0.95)
        t.sleep(2)
        if (finance_on):
            # print('수납집계 창 켜졌어요1')
            
        else:
            print('에러예요1')
            
    else:
        hiq_magam = pag.locateCenterOnScreen('d_hiq_magam.PNG', confidence=0.85)
        t.sleep(2)
        if (hiq_magam):
            pag.click(hiq_magam)
            t.sleep(0.5)
            pag.click(730, 62)
            t.sleep(0.5)
            finance_on = pag.locateCenterOnScreen('d_finance_on.PNG', confidence=0.95)
            t.sleep(2)
            if (finance_on):
                # print('수납집계 창 켜졌어요2')
                t.sleep(1)
                
        print('에러예요2')
        
elif (find_hiq):
    pag.click(find_hiq)
    t.sleep(1)
    hiq_soonap = pag.locateCenterOnScreen('d_hiq_soonap.PNG', confidence=0.85, region=(305, 111, 1106, 719))
    if (hiq_soonap):
        pag.click(hiq_soonap)
        t.sleep(10)
        hiq_jeopsoo = pag.locateCenterOnScreen('d_hiq_jeopsoo.PNG', confidence=0.85)
        # hiq_jeopsoo_on = pag.locateCenterOnScreen('d_hiq_jeopsoo_on.PNG', confidence=0.85)
        if (hiq_jeopsoo):
            pag.click(hiq_jeopsoo)
            t.sleep(1)
        hiq_magam = pag.locateCenterOnScreen('d_hiq_magam.PNG', confidence=0.85)
        if (hiq_magam):
            pag.click(hiq_magam)
            t.sleep(0.5)
            pag.click(730, 62)
            t.sleep(2)
            if (finance_on):
                # print('수납집계 창 켜졌어요3')
                t.sleep(1)
                
    else:
        print('왜 수납아이콘 없지')