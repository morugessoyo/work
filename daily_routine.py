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

pix_status_line = (210, 210, 210)   # 회색 줄이 없으면 환자가 엄서요!



crm_password = '38tkfdl!'


def get_week_of_month(year, month, day):
    x = np.array(calendar.monthcalendar(year, month))
    week_of_month = np.where(x==day)[0][0] + 1
    return(week_of_month)

def get_date(y, m, d):
    '''y: year(4 digits)
     m: month(2 digits)
     d: day(2 digits'''
    s = f'{y:04d}-{m:02d}-{d:02d}'
    return datetime.strptime(s, '%Y-%m-%d')

def click_date(day_day, what_week):  # 화면 왼쪽 위에 펼쳐져있는 달력에서 날짜 클릭하기
    if keyboard.is_pressed('END'):
        return

    # 달력에서 날짜 클릭하기 시작
    # print('얼마인가?:',(day_day + 1) % 7, 'day_day:', day_day)
    if (((day_day + 1) % 7)==0):
        pag.click(46 + ((day_day + 1) % 7) * 20, 127 + ((what_week+1) * 16))  # 날짜 더블클릭!
    else:
        pag.click(46 + ((day_day + 1) % 7) * 20, 127 + (what_week * 16))  # 날짜 더블클릭!
    # print('day_day:', day_day + 1,' what_week:', what_week)
    # print('달력 클릭위치는:', 1631 + ((day_day + 1) % 7) * 40, 155 + (what_week * 25))
    # print('달력 클릭:', 46 + ((day_day + 1) % 7) * 20, 127 + (what_week * 16))
    t.sleep(1)
    return


def clear_screen():
    if keyboard.is_pressed('END'):
        return

    c_yes = pag.locateCenterOnScreen('c_yes.PNG', confidence=0.9)
    c_yes2 = pag.locateCenterOnScreen('c_yes2.PNG', confidence=0.9)
    c_yes1 = pag.locateCenterOnScreen('c_yes1.PNG', confidence=0.9)
    c_save = pag.locateCenterOnScreen('c_save.PNG', confidence=0.9)
    if (c_yes):
        pag.click(c_yes)
        t.sleep(1)
    if (c_yes1):
        pag.click(c_yes1)
        t.sleep(1)
    if (c_yes2):
        pag.click(c_yes2)
        t.sleep(1)
    if (c_save):
        pag.click(c_save)
        t.sleep(1)
    return

def crm_on_check():                # CRM 켜져있는지 확인하는 함수
    if keyboard.is_pressed('END'):
        return

    crm_on = pag.locateCenterOnScreen('a_crm_on.PNG', confidence=0.9, region=(105, 1038, 717, 42))
    crm_off = pag.locateCenterOnScreen('a_crm_off.PNG', confidence=0.9, region=(105, 1038, 717, 42))
    if (crm_on):
        crm_on_screen = pag.locateCenterOnScreen('a_crm_on_screen.PNG', confidence=0.9)
        if not (crm_on_screen):
            pag.click(crm_on)
        return
    # else:
    #     pag.click(crm_off)
    #     t.sleep(1)
    #     pag.click(851, 665)
    #     t.sleep(1)
    #     crm_login = pag.locateCenterOnScreen('a_crm_login.PNG', confidence=0.9)
    #     pag.click(crm_login)
    #     t.sleep(0.5)
    #     pag.write(crm_password)
    #     pag.hotkey('enter')
    #     t.sleep(2.5)
    #     crm_resize = pag.locateCenterOnScreen('a_crm_resize.PNG', confidence=0.9)
    #     pag.click(crm_resize)
    #     t.sleep(0.5)
    #     crm_schedule = pag.locateCenterOnScreen('a_crm_schedule.PNG', confidence=0.9)
    #     pag.click(crm_schedule)
    #     t.sleep(0.5)

def hiq_on_check():
    if keyboard.is_pressed('END'):
        return

    hiq_icon = pag.locateCenterOnScreen('d_hiq_off.PNG', confidence=0.95, region=(105, 1038, 717, 42))
    find_hiq = pag.locateCenterOnScreen('d_find_hiq.PNG', confidence=0.95, region=(105, 1038, 717, 42))
    if (hiq_icon):    # 접수/수납프로그램 켜져있음(아이콘있음)
        pag.click(hiq_icon)
        t.sleep(1)
        hiq_jeopsoo = pag.locateCenterOnScreen('d_hiq_jeopsoo.PNG', confidence=0.9)
        if (hiq_jeopsoo):   # 접수등록탭 꺼져있음
            pag.click(hiq_jeopsoo)     # 여기까지 하면 접수등록탭 파란불 들어온 화면이 되지요!
            t.sleep(1)

    elif (find_hiq):   # 메인메뉴만 켜져있음(아이콘있음)
        pag.click(find_hiq)
        t.sleep(1)
        hiq_soonap = pag.locateCenterOnScreen('d_hiq_soonap.PNG', confidence=0.9, region=(305, 111, 1106, 719))
        if (hiq_soonap):      # 메인메뉴에서 수납버튼 클릭
            pag.click(hiq_soonap)
            t.sleep(10)
            # hiq_jeopsoo_on = pag.locateCenterOnScreen('d_hiq_jeopsoo_on.PNG', confidence=0.9)
            hiq_jeopsoo = pag.locateCenterOnScreen('d_hiq_jeopsoo.PNG', confidence=0.9) # 필요없을 것 같지만 '접수등록'(파란바탕) 확인용
            if (hiq_jeopsoo):     # 수납/등록창 띄워졌는데 접수등록에 파란불이 안들어와있다?
                pag.click(hiq_jeopsoo)
                t.sleep(1)   # 여기까지 하면 접수등록창 띄워져있다
        else:
            print('왜 수납아이콘 없지')
            return False
    else:
        print('차트프로그램 켜야해')
        return False

    hiq_magam = pag.locateCenterOnScreen('d_hiq_magam.PNG', confidence=0.9)
    if (hiq_magam):    # 접수등록 창에 마감관리버튼 보인다!
        pag.click(hiq_magam)
        t.sleep(0.5)
        pag.click(730, 62)    # 마감관리 클릭 후 수납집계로 들어감
        t.sleep(2)
        finance_on = pag.locateCenterOnScreen('d_finance_on.PNG', confidence=0.95)
        if (finance_on):   # 수납집계창 보임
            # print('수납집계 창 켜졌어요3')
            t.sleep(1)
            return True
    # else:
        # print('마감관리버튼이 안보여요')
    finance_o = pag.locateCenterOnScreen('d_finance_o.PNG', confidence=0.95)
    t.sleep(1)
    if (finance_o):     # 아이플러스차트 아이콘 + 수납집계(회색바탕) 확인
        pag.click(finance_o)
        finance_on = pag.locateCenterOnScreen('d_finance_on.PNG', confidence=0.95)
        t.sleep(2)
        if not (finance_on):      # 수납집계창의 파란바탕 수납집계 버튼 보인다! / 안보인다!
        #     print('수납집계 창 켜졌어요1')
        # else:                 # 수납집계창이 아니구먼..
            print('수납집계창이 아니구먼..')
            return False
    else:
        print('수납집계 창이 아니여...')
        return False

def crm_surgery_only():
    if keyboard.is_pressed('END'):
        return

    cond_surgery = pag.locateCenterOnScreen('c_surgery.PNG', confidence=0.9, region=(84, 222, 147, 99))
    # CRM 켜진 상태에서 '수술'클릭확인
    if (cond_surgery):  # '수술' 클릭 안되어있으면?
        pag.click(cond_surgery)  # 클릭하고 갑시다
    return

def click_start_date(yy,mm,dd):
    # print('yy,mm,dd', yy, mm, dd)
    pag.click(109, 68) # 수납시작일자의 년도 클릭
    t.sleep(0.1)
    pag.write(yy)
    t.sleep(0.2)
    pag.write(mm)
    t.sleep(0.2)
    pag.write(dd)
    t.sleep(0.2)
    return

def click_end_date(yy,mm,dd):
    pag.click(234, 70)  # 수납종료일자의 년도 클릭
    t.sleep(0.1)
    pag.write(yy)
    t.sleep(0.2)
    pag.write(mm)
    t.sleep(0.2)
    pag.write(dd)
    t.sleep(0.2)
    return

def copy_and_paste():    # 왼쪽 날짜 오른쪽에 복붙
    pag.click(133,74)   # 다시 날짜부분 클릭
    pag.hotkey('ctrl','a')
    t.sleep(0.2)
    pag.hotkey('ctrl','c')
    t.sleep(0.2)

    pag.click(234,74)   # 오른쪽 날짜부분 클릭
    pag.hotkey('ctrl','v')
    t.sleep(0.2)
    return

def normalize(s): # 콤마 없애주세요 흑흑 ※
    if s == None:
        return 0
    elif s != None:
        return s.replace(',', '').strip()






# 여기서부터 실제 코드!
past_data = False     # 하루만 돌리고 싶거든 사용합시다
hyoo_gaa = 0       # 휴가 다녀왔으면 숫자 바꾸기

if (past_data):
    today = date(2022, 4, 1)
    yy = today.year
    mm = today.month
    dd = today.day
    wkday = today.weekday()
    print('yy:', yy, 'mm:', mm, 'dd:', dd, 'wkday:', wkday)
    days_to_go = 1
else:

    today = date.today()
    # today = date(2022, 4, 18)     # 특정 날짜 돌리고 싶으면 여기를... 근데 하루이틀만 돌리는건 안되넹.. 어케하지?
    yy = today.year
    mm = today.month
    dd = today.day
    wkday = today.weekday()
    print('yy:', yy, 'mm:', mm, 'dd:', dd, 'wkday:', wkday)

    if(wkday == 0): # 만약 오늘이 월요일이면?
        print('월요일이에요!')
        dd = dd - 3 - hyoo_gaa        # 3일전으로 돌아가(금) - 휴가일자추가
        days_to_go = 2 + hyoo_gaa     # 이틀 돌릴거야 + 휴가일자추가
        # print('dd:', dd)
    elif(wkday != 0):
        print('월요일이 아니에요!')
        dd = dd - 1 - hyoo_gaa       # 어제 날짜로 돌아가!(화->월, 수->화 이런식으로) - 휴가일자 추가
        days_to_go = 1 + hyoo_gaa    # 휴가일자 추가
        # print('dd:', dd)

while True:  # 루프문 들어와써요!
    if keyboard.is_pressed('END'):
        break

    for i in range(0, days_to_go, 1):  # 원하는 날짜만큼 돌립니다
        if keyboard.is_pressed('end'):
            print('end 눌러 종료합니다.')
            break

        # 1. 수술자 목록 돌려!
        crm_on_check()   # CRM 켜져있는지 확인
        t.sleep(1)
        crm_surgery_only()  # 수술자 창인지 확인
        t.sleep(2)

        print('CRM 기준(현재)날짜:', get_date(yy, mm, dd).strftime('%Y-%m-%d'))
        what_week = get_week_of_month(yy, mm, dd)   # 몇째주
        day_day = get_date(yy, mm, dd).weekday()    # 몇번째날?
        # print(what_week - 1, '번째 주의 ', day_day, '째날이예요')
        click_date(day_day, what_week - 1)

        no_surgery_day = pag.locateCenterOnScreen('c_no_surgery_day.PNG', confidence=0.98, region=(12, 331, 178, 76))

        if (no_surgery_day):
            if keyboard.is_pressed('end'):
                print('end 눌러 종료합니다.')
                break
            print("오늘은 수술이 없어요")
        else:
            if keyboard.is_pressed('end'):
                print('end 눌러 종료합니다.')
                break
            # print("오늘은 수술이 있구만요")
            t.sleep(0.3)
            pag.click(1076, 93)
            t.sleep(0.3)
            clear_screen()  # 예 누름
            t.sleep(0.5)
            pag.write(crm_password)  # CRM 비번 입력
            t.sleep(0.3)
            clear_screen()  # 예 누름
            cur_date = get_date(yy, mm, dd).strftime('%Y-%m-%d')  # 파일이름 저장(날짜로)
            # print('cur_date:', cur_date)
            pag.write(cur_date)
            t.sleep(0.5)

            # 엑셀파일을 '작업 - 011_수술자목록 - CRM'에 저장하려고 만든거
            pag.click(pag.locateCenterOnScreen('c_Gdrive.PNG', confidence=0.9))
            t.sleep(0.3)
            pag.doubleClick(pag.locateCenterOnScreen('c_jakup.PNG', confidence=0.9))
            t.sleep(0.3)
            pag.doubleClick(pag.locateCenterOnScreen('c_surgery_folder.PNG', confidence=0.9))
            t.sleep(0.3)
            pag.doubleClick(pag.locateCenterOnScreen('c_crm_folder.PNG', confidence=0.9))
            t.sleep(0.3)
            # pag.doubleClick(pag.locateCenterOnScreen('c_2022_folder.PNG', confidence=0.9)) # 2022-01 폴더 들어가는거라 주석처리함
            clear_screen()  # 저장 누름
            t.sleep(0.1)
            clear_screen()  # 만약 같은 이름 가진 파일 있다? 덮어쓰기!

            # 수술자목록파일 저장했으니, 이제 매크로파일 불러와서 새로만든 파일에 매크로 돌려놓읍시다.
            # 엑셀 띄우지 않고 실행!
            app = xw.App(visible=False)

            file_dir_name = 'G:/작업/011_수술정리/crm/'
            wb = xw.Book('C:\\Users\\onnuri\\Documents\\crm\\2019-05-01.xlsm')
            my_macro = wb.macro('수술자2')   # 매크로 이름 여기에 넣어요
            wb1 = xw.Book(file_dir_name + cur_date+'.xls')
            my_macro()                     # 매크로 돌려요
            t.sleep(0.5)
            wb1.save(file_dir_name + cur_date + '.xls')
            t.sleep(0.5)
            wb1.close()
            t.sleep(0.5)
            wb.close()
            t.sleep(3)
            print('CRM 수술자 목록 저장 완료:'+file_dir_name+cur_date+'.xls')

            # 2. 하이큐 일일장부 만들어요
            hiq_on_check()
            t.sleep(0.5)
            print('하이큐 기준(현재)날짜:', get_date(yy, mm, dd).strftime('%Y-%m-%d'))

            y1 = str(yy)
            m1 = str(mm)
            d1 = str(dd)

            # 여기부터 돌려요
            cc = get_date(yy, mm, dd).year
            bb = get_date(yy, mm, dd).month
            aa = get_date(yy, mm, dd).day
            # print('year?:', cc, 'month?:', bb, 'day?:', aa)

            click_start_date(y1, m1, d1)
            copy_and_paste()
            search1 = pag.locateCenterOnScreen('z_search.PNG', confidence=0.9)
            pag.click(search1)  # 조회 클릭
            t.sleep(3)
            # print('조회클릭')
            pag.click(103, 333)  # 내용 클릭
            t.sleep(0.2)
            pag.hotkey('ctrl', 'a')
            t.sleep(0.2)
            # print('ctrl a')
            pag.hotkey('ctrl', 'c')
            t.sleep(0.2)
            # print('ctrl c')

            text = pyperclip.paste()  # 얘들아 복사가!! 복사가 돼!!!!

            # 방법4 된다!!!!!!
            text = normalize(text)

            # clipboard.txt 에 데이터 저장
            with open('clipboard.txt', 'w', encoding='utf8') as file:  # 클립보드 내용을 txt로 저장
                file.write(text)  # 잘 저장되네요
            t.sleep(0.5)
            df = pd.DataFrame(pd.read_csv('clipboard.txt', sep='	'))  # 1번!!! df에다가 데이타 저장

            # print('기준(현재)날짜:', get_date(yy, mm, dd))

            # 엑셀파일 열어요
            wb = openpyxl.Workbook()  # workbook 만들기
            cur_date = get_date(yy, mm, dd).strftime('%Y-%m-%d')
            # print('cur_date:', cur_date)
            filename = 'G:/작업/010_수납리스트/2022 수납리스트/' + str(cur_date) + '.xlsx'
            # print('파일저장장소+파일이름:', filename)
            wb.save(filename)  # 엑셀파일 저장하기
            df.to_excel(filename, index=False)  # 2번!!! 엑셀파일에 내용 저장
            wb.close()
            t.sleep(0.5)

            # 엑셀 띄우지 않고 실행!
            app = xw.App(visible=False)
            # 여기부터는 일일장부 매크로 돌려용
            wb3 = xw.Book('C:\\Users\\onnuri\\Documents\\crm\\2022-05-01.xlsm')  # 일일장부 매크로 저장한 파일
            my_macro = wb3.macro('일일장부서식')
            file_dir_name = 'G:/작업/010_수납리스트/2022 수납리스트/'
            wb1 = xw.Book(file_dir_name + cur_date + '.xlsx')
            my_macro()
            wb1.save(file_dir_name + cur_date + '.xlsx')
            wb1.close()
            t.sleep(0.5)
            wb3.close()
            t.sleep(3)
            print('하이큐 일일장부 작성 완료:' + file_dir_name + cur_date + '.xlsx')
            # 여기까지 엑셀 저장 + 매크로 작업 완료!


            # 달의 마지막 날짜면 ... 다음달로!
            date = get_date(yy, mm, dd)
            last_day = calendar.monthrange(date.year, date.month)[1]
            last_day1 = get_date(yy, mm, last_day)

            if (date == last_day1):
                crm_on_check()      # CRM 켜져있는지 다시 확인!
                t.sleep(1)
                pag.click(185, 94)  # CRM에서 달력의 다음달(>) 버튼 클릭해요
                if ((mm + 1) % 12 == 0):
                    mm = 12
                    dd = 1
                elif ((mm + 1) % 12 == 1):
                    yy = yy + 1
                    mm = (mm + 1) % 12
                    dd = 1
                else:
                    mm = (mm + 1) % 12
                    dd = 1
            else:
                dd = dd + 1

            t.sleep(1)
            if (i + 1 != days_to_go):
                print('다음 날짜로 진행합니다.')

    print('사이클 완료!')
    break