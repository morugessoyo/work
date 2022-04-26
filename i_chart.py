import numpy as np
import pyautogui as pag
import time as t
import keyboard
from datetime import *
import calendar
from PIL import ImageGrab

calendar.setfirstweekday(6)

pix_status_line = (210, 210, 210)

one_day_only = False

yy = 2022
mm = 4
dd = 20

end_yy = 2022
end_mm = 4
end_dd = 21

# dd_end_date = 11

password = 'dhssnfl3!'

def position_check():
    screen = ImageGrab.grab()
    pix_status = screen.getpixel((1195,62))  # 상단 회색 바 생겼는지 확인
    if (pix_status == (171, 171, 171)):
        pag.click(60, 98)
        t.sleep(0.1)
        pag.click(121, 68)
    else:
        return

def get_date(y, m, d):
    '''y: year(4 digits)
     m: month(2 digits)
     d: day(2 digits'''
    s = f'{y:04d}-{m:02d}-{d:02d}'
    return datetime.strptime(s, '%Y-%m-%d')

def get_week_of_month(year, month, day):
    x = np.array(calendar.monthcalendar(year, month))
    week_of_month = np.where(x==day)[0][0] + 1
    return(week_of_month)

# def get_week_no(y, m, d):
#     target = get_date(y, m, d)
#     firstday = target.replace(day=1)
#     if firstday.weekday() == 6:
#         origin = firstday
#     elif firstday.weekday() < 3:
#         origin = firstday - timedelta(days=firstday.weekday() + 1)
#     else:
#         origin = firstday + timedelta(days=6-firstday.weekday())
#     return (target - origin).days // 7 + 1

def Chunggoo():   # 코드에 따라 상병 추가하는 작업하기
    # t.sleep(2)
    print('청구합시다')
    dur_clear()

    # 상병
    h0411 = pag.locateCenterOnScreen('h0411.png', confidence=0.92, region=(508, 76, 782, 854))
    h0204 = pag.locateCenterOnScreen('h0204.png', confidence=0.92, region=(508, 76, 782, 854))
    h1618 = pag.locateCenterOnScreen('h1618.png', confidence=0.92, region=(508, 76, 782, 854))
    h0000 = pag.locateCenterOnScreen('h0000.png', confidence=0.92, region=(508, 76, 782, 854))
    h0001 = pag.locateCenterOnScreen('h0001.png', confidence=0.92, region=(508, 76, 782, 854))
    h001 = pag.locateCenterOnScreen('h001.png', confidence=0.92, region=(508, 76, 782, 854))
    h1111 = pag.locateCenterOnScreen('h1111.png', confidence=0.92, region=(508, 76, 782, 854))
    h101 = pag.locateCenterOnScreen('h101.png', confidence=0.92, region=(508, 76, 782, 854))
    h258 = pag.locateCenterOnScreen('h258.png', confidence=0.92, region=(508, 76, 782, 854))
    h2602 = pag.locateCenterOnScreen('h2602.png', confidence=0.92, region=(508, 76, 782, 854))
    h105 = pag.locateCenterOnScreen('h105.png', confidence=0.92, region=(508, 76, 782, 854))
    h108 = pag.locateCenterOnScreen('h108.png', confidence=0.92, region=(508, 76, 782, 854))
    h019 = pag.locateCenterOnScreen('h019.png', confidence=0.92, region=(508, 76, 782, 854))
    h182 = pag.locateCenterOnScreen('h182.png', confidence=0.92, region=(508, 76, 782, 854))
    h188 = pag.locateCenterOnScreen('h188.png', confidence=0.92, region=(508, 76, 782, 854))
    h191 = pag.locateCenterOnScreen('h191.png', confidence=0.92, region=(508, 76, 782, 854))
    h193 = pag.locateCenterOnScreen('h193.png', confidence=0.92, region=(508, 76, 782, 854))
    h33 = pag.locateCenterOnScreen('h33.png', confidence=0.92, region=(508, 76, 782, 854))
    h350 = pag.locateCenterOnScreen('h350.png', confidence=0.92, region=(508, 76, 782, 854))
    # h3338 = pag.locateCenterOnScreen('h3338.png', confidence=0.92, region=(508, 76, 782, 854))
    h400 = pag.locateCenterOnScreen('h400.png', confidence=0.92, region=(508, 76, 782, 854))
    h526 = pag.locateCenterOnScreen('h526.png', confidence=0.92, region=(508, 76, 782, 854))
    z947 = pag.locateCenterOnScreen('z947.png', confidence=0.92, region=(508, 76, 782, 854))
    z999 = pag.locateCenterOnScreen('z999.png', confidence=0.92, region=(508, 76, 782, 854))
    z998 = pag.locateCenterOnScreen('z998.png', confidence=0.92, region=(508, 76, 782, 854))
    z997 = pag.locateCenterOnScreen('z997.png', confidence=0.92, region=(508, 76, 782, 854))
    z961 = pag.locateCenterOnScreen('z961.png', confidence=0.92, region=(508, 76, 782, 854))

    # 약, 검사수가 등 차트내용
    dcombi = pag.locateCenterOnScreen('dcombi.png', confidence=0.92, region=(508, 76, 782, 854))
    dcosop = pag.locateCenterOnScreen('dcosop.png', confidence=0.92, region=(508, 76, 782, 854))
    ddicuapo = pag.locateCenterOnScreen('ddicuapo.png', confidence=0.92, region=(508, 76, 782, 854))
    dfluoro = pag.locateCenterOnScreen('dfluoro.png', confidence=0.92, region=(508, 76, 782, 854))
    dhyal_1 = pag.locateCenterOnScreen('dhyal_1.png', confidence=0.92, region=(508, 76, 782, 854))
    dmoxi = pag.locateCenterOnScreen('dmoxi.png', confidence=0.92, region=(508, 76, 782, 854))
    doflo = pag.locateCenterOnScreen('doflo.png', confidence=0.92, region=(508, 76, 782, 854))
    doflox = pag.locateCenterOnScreen('doflox.png', confidence=0.92, region=(508, 76, 782, 854))
    dolopata = pag.locateCenterOnScreen('dolopata.png', confidence=0.92, region=(508, 76, 782, 854))
    dporus = pag.locateCenterOnScreen('dporus.png', confidence=0.92, region=(508, 76, 782, 854))
    dlil = pag.locateCenterOnScreen('dlil.png', confidence=0.92, region=(508, 76, 782, 854))
    dposod = pag.locateCenterOnScreen('dposod.png', confidence=0.92, region=(508, 76, 782, 854))
    dpire = pag.locateCenterOnScreen('dpire.png', confidence=0.92, region=(508, 76, 782, 854))
    dacyclo = pag.locateCenterOnScreen('dacyclo.png', confidence=0.92, region=(508, 76, 782, 854))
    dcyclo = pag.locateCenterOnScreen('dcyclo.png', confidence=0.92, region=(508, 76, 782, 854))

    e6660 = pag.locateCenterOnScreen('e6660.png', confidence=0.92, region=(508, 76, 782, 854))
    e6670 = pag.locateCenterOnScreen('e6670.png', confidence=0.92, region=(508, 76, 782, 854))
    e6674 = pag.locateCenterOnScreen('e6674.png', confidence=0.92, region=(508, 76, 782, 854))
    e6710 = pag.locateCenterOnScreen('e6710.png', confidence=0.92, region=(508, 76, 782, 854))
    e6899 = pag.locateCenterOnScreen('e6899.png', confidence=0.92, region=(508, 76, 782, 854))
    s5160 = pag.locateCenterOnScreen('s5160.png', confidence=0.92, region=(508, 76, 782, 854))
    s5400 = pag.locateCenterOnScreen('s5400.png', confidence=0.92, region=(508, 76, 782, 854))
    s5430 = pag.locateCenterOnScreen('s5430.png', confidence=0.92, region=(508, 76, 782, 854))
    s5390 = pag.locateCenterOnScreen('s5390.png', confidence=0.92, region=(508, 76, 782, 854))
    s4960 = pag.locateCenterOnScreen('s4960.png', confidence=0.92, region=(508, 76, 782, 854))
    ez796 = pag.locateCenterOnScreen('ez796.png', confidence=0.92, region=(508, 76, 782, 854))

    hempty = pag.locateCenterOnScreen('hempty.png', confidence=0.92, region=(508, 76, 782, 425)) # 상병에 빈자리 찾아요
    hblank = pag.locateCenterOnScreen('hblank.png', confidence=0.95, region=(508, 76, 782, 425)) # 상병이 아예 없어요

    while True:
        point_y_check = hempty[1]
        point_x = hempty[0]
        point_y = hempty[1]
        # print('hempty위치:', point_y, point_y_check, hblank)

        if keyboard.is_pressed('END'):
            break

        if (dhyal_1) and not (h0411):
            pag.click(point_x, point_y)
            keyboard.write('h0411')
            pag.hotkey('enter')
            print('히알루론 상병추가: h0411')
            point_y = point_y+22
            # t.sleep(0.1)

        if (dcyclo) and not (h0411):
            pag.click(point_x, point_y)
            keyboard.write('h0411')
            pag.hotkey('enter')
            t.sleep(0.1)
            keyboard.write('h1618')
            pag.hotkey('enter')
            print('레스타시스/사이클로스 상병추가: h0411, h1618')
            point_y = point_y+44
            # t.sleep(0.1)

        if (dposod) or (dpire) and not (h258):
            pag.click(point_x, point_y)
            keyboard.write('h2582')
            pag.hotkey('enter')
            print('포소드/산텐가리 상병추가: h2582')
            point_y = point_y+22
            # t.sleep(0.1)
        if (dacyclo)  and not (h191):
            pag.click(point_x, point_y)
            keyboard.write('h191')
            pag.hotkey('enter')
            print('헤르페시드/아시클로버 상병추가: h191')
            point_y = point_y + 22

        if (dfluoro) and not (h1618) and not (h193):
            pag.click(point_x, point_y)
            keyboard.write('h1618')
            pag.hotkey('enter')
            print('플루오로메톨론 상병추가: h1618')
            point_y = point_y+22
            # t.sleep(0.1)

        if ((e6660) or (e6670) or (s5160)) and not (h350) and not (h33) and not (h258):
            if ((e6660) or (e6670)) and not (h350) :
                pag.click(point_x, point_y)
                keyboard.write('h350')
                pag.hotkey('enter')
                print('안저검사 상병추가: h350')
                point_y = point_y + 22
                # t.sleep(0.1)
            else:
                pag.click(point_x, point_y)
                keyboard.write('h3338')
                pag.hotkey('enter')
                print('안저광응고 상병추가: h3338')
                point_y = point_y + 22
                # t.sleep(0.1)

        if (e6899) and not (h182):
            pag.click(point_x, point_y)
            keyboard.write('h182')
            pag.hotkey('enter')
            print('각막부종 상병추가: h182')
            point_y = point_y + 22
            # t.sleep(0.1)

        if (e6710) and not (h526) and not (h258) and not (h2602):
            pag.click(point_x, point_y)
            keyboard.write('h526')
            pag.hotkey('enter')
            print('안경처방 상병추가: h526')
            point_y = point_y+22
            # t.sleep(0.1)

        if (s5430) and not (h0204):
            pag.click(point_x, point_y)
            keyboard.write('h0204')
            pag.hotkey('enter')
            print('s5430 상병추가: h0204')
            point_y = point_y+22
            # t.sleep(0.1)

        if (s5390) and not (z947):
            pag.click(point_x, point_y)
            keyboard.write('z947')
            pag.hotkey('enter')
            print('s5390 상병추가: z947')
            point_y = point_y+22
            # t.sleep(0.1)

        if (s4960) and not (h1111):
            pag.click(point_x, point_y)
            keyboard.write('h1111')
            pag.hotkey('enter')
            print('s4690 상병추가: h1111')
            point_y = point_y+22
            # t.sleep(0.1)

        if (s5400) and not (h0000):
            pag.click(point_x, point_y)
            keyboard.write('h0000')
            pag.hotkey('enter')
            print('s5400 상병추가: h0000')
            point_y = point_y+22
            # t.sleep(0.1)

        if (dmoxi) and not (h1618) and not (h188) and not (z961):
            pag.click(point_x, point_y)
            keyboard.write('h1618')
            pag.hotkey('enter')
            print('목시플록사신 상병추가: h1618')
            point_y = point_y+22
            # t.sleep(0.1)

        if (doflox) and not (h108) and not (h0001) and not (h105) and not (z961) and not (z997):
            pag.click(point_x, point_y)
            keyboard.write('h108')
            pag.hotkey('enter')
            print('오플록사신정 상병추가: h108')
            point_y = point_y+22
            # t.sleep(0.1)

        if ((doflo) or (dporus)) and not (h105) and not (z999) and not (h019):
            pag.click(point_x, point_y)
            keyboard.write('h105')
            pag.hotkey('enter')
            print('포러스/타리비드/오플록사신 상병추가: h105')
            point_y = point_y+22
            # t.sleep(0.1)

        if ((dcombi) or (dcosop)) and not (h400):
            pag.click(point_x, point_y)
            keyboard.write('h400')
            pag.hotkey('enter')
            print('콤비간/코솝 상병추가: h400')
            point_y = point_y+22
            # t.sleep(0.1)

        if (dolopata) and not (h101):
            pag.click(point_x, point_y)
            keyboard.write('h101')
            pag.hotkey('enter')
            print('올로파타딘 상병추가: h101')
            point_y = point_y+22
            # t.sleep(0.1)

        if (ddicuapo) and not (h101) and not (h1618) and not (z998):
            pag.click(point_x, point_y)
            keyboard.write('h1618')
            pag.hotkey('enter')
            print('디쿠아포솔 상병추가: h1618')
            point_y = point_y+22
            # t.sleep(0.1)

        if (ez796) and not (h350):
            pag.click(point_x, point_y)
            keyboard.write('h350')
            pag.hotkey('enter')
            print('안구광학단층촬영 상병추가: h350')
            point_y = point_y+22
            # t.sleep(0.1)

        if (dlil) and not (h101):
            pag.click(point_x, point_y)
            keyboard.write('h101')
            pag.hotkey('enter')
            print('릴레스타트 상병추가: h101')
            point_y = point_y + 22
            # t.sleep(0.1)

        if (e6670) and (e6674):
            print('6670 + 6674')
            if (e6660):
                print('6660도 있어')
                pag.click(e6670)
                keyboard.write('0')
                pag.hotkey('enter')
                print('광각안저촬영 with 기본안저촬영이라 지웁니다')
                # t.sleep(0.1)
            else:
                print('6660은 없어')
                pag.click(e6670)
                keyboard.write('e6660')
                pag.hotkey('enter')
                print('광각안저촬영 with 기본안저촬영이라 지우고 안저검사 추가했어요')
                point_y = point_y + 22
                # t.sleep(0.1)
                keyboard.write('2')
                pag.hotkey('enter')
                # t.sleep(0.1)

        if (point_y == point_y_check):
            print('추가상병 없음')
            if (hblank):
                print('상병이 없어서 추가: h193')
                pag.click(point_x, point_y)
                keyboard.write('h193')
                pag.hotkey('enter')
                t.sleep(0.1)
                pag.hotkey('F3')
                t.sleep(1)
                dur_clear()
                t.sleep(0.1)
                return True
            return False
        else:
            print('추가상병 있음')
            pag.hotkey('F3')
            t.sleep(1)   # 2에서 1로 바꿔봄
            dur_clear()
            t.sleep(0.1)
            return True


def Updown(updown):
    if updown == 'down':
        pag.click(1332, 905)
        t.sleep(0.5)
        pag.scroll(-2200) # 2200이면 18명정도 내리는듯
        t.sleep(1)
    if updown == 'up':
        pag.click(1332, 905)
        t.sleep(0.5)
        pag.scroll(2200)  # 올려욧 올려욧
        t.sleep(1)

def pat_clear():
    # print('  ')
    pag.click(70, 989)
    t.sleep(0.1)

def click_pat():
    if keyboard.is_pressed('END'):
        return
    pat_exist2 = pag.locateCenterOnScreen('pat_exist2.png', confidence=0.86)   # "진료실(선택돼서 어두운바탕)"
    pat_exist3 = pag.locateCenterOnScreen('pat_exist3.png', confidence=0.86)  # "진료실(흰바탕)"
    err_durinfo = pag.locateCenterOnScreen('err-durinfo.png', confidence=0.9, region=(1862,407, 66,641)) # dur 안내창 x표

    pat_x = 1390   # pat_x 환자이름클릭 x값
    pat_y = 440    # pat_y 환자이름클릭 y값
    pat_no = 1     # pat_no 순번 1부터시작
    click_position = 0   # click_position 클릭할 순번 0부터시작
    check_pat_1st = 0    # check_pat_1st top 자리에 있는 환자번호

    if (pat_exist2) or (pat_exist3):
        for o in range(0, 100, 1):    # 1~101번째 환자까지 차례대로 클릭해요
            if keyboard.is_pressed('END'):
                break
            dur_clear()      # 우선 화면 정리!
            pat_clear()

            if (err_durinfo): # dur 공지사항 팝업 있으면 삭제(화면 오른쪽아래 코너)
                pag.click(err_durinfo)
                t.sleep(0.5)
            # print('o:', o, 'click_position:', click_position)
            screen = ImageGrab.grab()
            pix_status = screen.getpixel((1340, 450 + click_position * 21))  # 1번 환자순서 오른아래코너에 회색코너 있는지 확인(환자가 있어요)
            # print('x값: 1340, y값:', 450 + click_position * 21, 'pix_status:',pix_status)

            if (pix_status == pix_status_line): # 회색 줄 있으면
                if (pat_y >= 905): # y값이 너무 크면 끝냅시다(23번 이후)
                    print('!너무 큰! pat_y', pat_y)
                    return

                if (pat_no < 23):     # 1~23번 환자까지는 그냥 돌려용
                    # print('click_position:', click_position)
                    pag.doubleClick(pat_x, pat_y)
                    print(pat_no,'번째 환자 클릭했어요.')

                if (pat_no == 23):  # 23/45번째 환자 끝나고 나면
                    Updown('up')     # 스크롤을 내려요~ 18명정도
                    Updown('up')     # 스크롤을 내려요~ 18명정도
                    Updown('down')   # 스크롤을 내려요~ 18명정도

                    check_pat_1st = find_num('top')   # 우선 목록의 첫번째 환자 번호 확인
                    check_pat_last = find_num('bottom')  # 우선 목록의 마지막 환자 번호 확인
                    # print('스크롤 후 첫번째환자번호:', check_pat_1st, '마지막환자번호:', check_pat_last)

                    # print('pat_y:', pat_y)
                    click_position = 23 - check_pat_1st

                    # print('23_click_position:', click_position)
                    pat_y = 440+ (21 * (click_position))
                    # print('pat_y', pat_y)
                    # t.sleep(0.1)

                    pag.doubleClick(pat_x, pat_y)
                    print(pat_no,'번째 환자 클릭했어요.')


                    # pat_y = pat_y + 21
                    # click_position = click_position +1


                # print('현재첫번째환자번호:', check_pat_2nd, '스크롤후?:', check_pat_1st)
                if (pat_no > 23):
                    check_pat_2nd = find_num('top')     # 현재 목록 첫번째에 있는 환자번호 확인
                    # print('check_pat_1st:', check_pat_1st, ', check_pat_2nd:', check_pat_2nd)
                    if (check_pat_1st == check_pat_2nd):
                        # print('click_position:', click_position)
                        # print('스크롤 확인. 첫번째 환자번호 같음: ', check_pat_2nd)
                        pag.doubleClick(pat_x, pat_y)
                        print(pat_no, '번째 환자 클릭했어요.')
                        # t.sleep(0.5)

                    else:
                        # print('click_position:', click_position)
                        # check_pat_2ndtime = find_num('top')
                        # print('스크롤 확인. 첫번째 환자번호 달라서 내림: ', check_pat_2nd)
                        Updown('down')
                        pag.doubleClick(pat_x, pat_y)
                        print(pat_no, '번째 환자 클릭했어요.')
                        # t.sleep(0.5)


                if (pat_no == 40):  # 23/45번째 환자 끝나고 나면
                    Updown('up')  # 스크롤을 올려요~ 18명정도
                    Updown('up')  # 스크롤을 올려요~ 18명정도
                    Updown('up')  # 스크롤을 올려요~ 18명정도
                    Updown('down')  # 스크롤을 내려요~ 18명정도
                    Updown('down')  # 스크롤을 내려요~ 18명정도

                    check_pat_1st = find_num('top')  # 우선 목록의 첫번째 환자 번호 확인
                    check_pat_last = find_num('bottom')  # 우선 목록의 마지막 환자 번호 확인
                    # print('스크롤 후 첫번째환자번호:', check_pat_1st, '마지막환자번호:', check_pat_last)

                    # print('pat_y:', pat_y)
                    click_position = 41 - check_pat_1st

                    # print('40_click_position:', click_position)
                    pat_y = 440+ (21 * (click_position))
                    # print('pat_y', pat_y)
                    t.sleep(0.1)

                    pag.doubleClick(pat_x, pat_y)
                    print(pat_no, '번째 환자 클릭했어요.')

                if (pat_no > 40):
                    check_pat_2nd = find_num('top')     # 현재 목록 첫번째에 있는 환자번호 확인
                    # print('check_pat_1st:', check_pat_1st, ', check_pat_2nd:', check_pat_2nd)

                    if (check_pat_1st == check_pat_2nd):
                        # print('click_position:', click_position)
                        # print('스크롤 확인용. 첫번째 환자번호 같음: ', check_pat_2nd)
                        pag.doubleClick(pat_x, pat_y)
                        print(pat_no, '번째 환자 클릭했어요.')
                    else:
                        # print('click_position:', click_position)
                        # check_pat_2ndtime = find_num('top')
                        # print('스크롤 확인용. 첫번째 환자번호 달라서 내림: ', check_pat_2nd)
                        Updown('down')
                        Updown('down')
                        t.sleep(0.1)
                        Updown('down')  # 두번 다운다운!
                        pag.doubleClick(pat_x, pat_y)
                        print(pat_no, '번째 환자 클릭했어요.')


                # dur에러 확인, 청구 돌리고, 다시 화면정리  - 잠시 멈춰요
                t.sleep(1.0)
                dur_clear()
                # t.sleep(0.5)
                Chunggoo()
                t.sleep(0.5)
                pat_clear()
                t.sleep(0.1)
                click_position = click_position + 1
                pat_y = pat_y + 21
                pat_no = pat_no + 1
            else:
                return

def dur_clear():     # 에러메시지 나타나면 확인확인클릭클릭
    dur_check = pag.locateCenterOnScreen('dur_check.png', confidence=0.9)
    err_check = pag.locateCenterOnScreen('err2.png', confidence=0.8, region=(745,335,650,500))
    err_1_check = pag.locateCenterOnScreen('err_1_check.png', confidence=0.9)
    save_yes = pag.locateCenterOnScreen('save_yes.png', confidence=0.9)
    if keyboard.is_pressed('END'):
        return
    if (dur_check) or (err_1_check):
        print('dur확인창!')
        if (dur_check):
            pag.click(dur_check)
        else:
            pag.click(err_1_check)
        t.sleep(0.5)

    if (err_check):
        print('err 확인창!')
        pag.click(err_check)
        t.sleep(0.5)

    if (save_yes):
        print('save!')
        pag.click(save_yes)
        t.sleep(0.5)


def click_date(day_day, what_week):       # 화면 맨오른쪽 위 펼쳐져있는 달력에서 날짜 클릭하기
    # 달력에서 날짜 클릭하기 시작
    pag.doubleClick(1631 + ((day_day + 1) % 7) * 40, 145 + (what_week * 25))  # 날짜 더블클릭!
    # print('몇번째요일?', day_day+1, '몇번째주?', what_week)
    # print('달력 클릭위치는:', 1631 + ((day_day + 1) % 7) * 40, 155 + (what_week * 25))
    t.sleep(1)
    return

def find_num(where):
    if (where == 'top'):
        x = 1282
        y = 427
    if (where == 'bottom'):
        x = 1279
        y = 887
    cond_num_1 = pag.locateCenterOnScreen('1.PNG', confidence=0.92, region=(x, y, 68, 34))
    cond_num_2 = pag.locateCenterOnScreen('2.PNG', confidence=0.92, region=(x, y, 68, 34))
    cond_num_3 = pag.locateCenterOnScreen('3.PNG', confidence=0.92, region=(x, y, 68, 34))
    cond_num_4 = pag.locateCenterOnScreen('4.PNG', confidence=0.92, region=(x, y, 68, 34))
    cond_num_5 = pag.locateCenterOnScreen('5.PNG', confidence=0.92, region=(x, y, 68, 34))
    cond_num_6 = pag.locateCenterOnScreen('6.PNG', confidence=0.92, region=(x, y, 68, 34))
    cond_num_7 = pag.locateCenterOnScreen('7.PNG', confidence=0.92, region=(x, y, 68, 34))
    cond_num_8 = pag.locateCenterOnScreen('8.PNG', confidence=0.92, region=(x, y, 68, 34))
    cond_num_9 = pag.locateCenterOnScreen('9.PNG', confidence=0.92, region=(x, y, 68, 34))
    cond_num_10 = pag.locateCenterOnScreen('10.PNG', confidence=0.92, region=(x, y, 68, 34))
    cond_num_11 = pag.locateCenterOnScreen('11.PNG', confidence=0.92, region=(x, y, 68, 34))
    cond_num_12 = pag.locateCenterOnScreen('12.PNG', confidence=0.92, region=(x, y, 68, 34))
    cond_num_13 = pag.locateCenterOnScreen('13.PNG', confidence=0.92, region=(x, y, 68, 34))
    cond_num_14 = pag.locateCenterOnScreen('14.PNG', confidence=0.92, region=(x, y, 68, 34))
    cond_num_15 = pag.locateCenterOnScreen('15.PNG', confidence=0.92, region=(x, y, 68, 34))
    cond_num_16 = pag.locateCenterOnScreen('16.PNG', confidence=0.92, region=(x, y, 68, 34))
    cond_num_17 = pag.locateCenterOnScreen('17.PNG', confidence=0.92, region=(x, y, 68, 34))
    cond_num_18 = pag.locateCenterOnScreen('18.PNG', confidence=0.92, region=(x, y, 68, 34))
    cond_num_19 = pag.locateCenterOnScreen('19.PNG', confidence=0.92, region=(x, y, 68, 34))
    cond_num_20 = pag.locateCenterOnScreen('20.PNG', confidence=0.92, region=(x, y, 68, 34))
    cond_num_21 = pag.locateCenterOnScreen('21.PNG', confidence=0.92, region=(x, y, 68, 34))
    cond_num_22 = pag.locateCenterOnScreen('22.PNG', confidence=0.92, region=(x, y, 68, 34))
    cond_num_23 = pag.locateCenterOnScreen('23.PNG', confidence=0.92, region=(x, y, 68, 34))
    cond_num_24 = pag.locateCenterOnScreen('24.PNG', confidence=0.92, region=(x, y, 68, 34))
    cond_num_25 = pag.locateCenterOnScreen('25.PNG', confidence=0.92, region=(x, y, 68, 34))
    cond_num_26 = pag.locateCenterOnScreen('26.PNG', confidence=0.92, region=(x, y, 68, 34))
    cond_num_27 = pag.locateCenterOnScreen('27.PNG', confidence=0.92, region=(x, y, 68, 34))
    cond_num_28 = pag.locateCenterOnScreen('28.PNG', confidence=0.92, region=(x, y, 68, 34))
    cond_num_29 = pag.locateCenterOnScreen('29.PNG', confidence=0.92, region=(x, y, 68, 34))
    cond_num_30 = pag.locateCenterOnScreen('30.PNG', confidence=0.92, region=(x, y, 68, 34))
    cond_num_31 = pag.locateCenterOnScreen('31.PNG', confidence=0.92, region=(x, y, 68, 34))
    cond_num_32 = pag.locateCenterOnScreen('32.PNG', confidence=0.92, region=(x, y, 68, 34))
    cond_num_33 = pag.locateCenterOnScreen('33.PNG', confidence=0.92, region=(x, y, 68, 34))
    cond_num_34 = pag.locateCenterOnScreen('34.PNG', confidence=0.92, region=(x, y, 68, 34))
    cond_num_35 = pag.locateCenterOnScreen('35.PNG', confidence=0.92, region=(x, y, 68, 34))
    cond_num_36 = pag.locateCenterOnScreen('36.PNG', confidence=0.92, region=(x, y, 68, 34))
    cond_num_37 = pag.locateCenterOnScreen('37.PNG', confidence=0.92, region=(x, y, 68, 34))
    cond_num_38 = pag.locateCenterOnScreen('38.PNG', confidence=0.92, region=(x, y, 68, 34))
    cond_num_39 = pag.locateCenterOnScreen('39.PNG', confidence=0.92, region=(x, y, 68, 34))
    cond_num_40 = pag.locateCenterOnScreen('40.PNG', confidence=0.92, region=(x, y, 68, 34))
    cond_num_41 = pag.locateCenterOnScreen('41.PNG', confidence=0.92, region=(x, y, 68, 34))
    cond_num_42 = pag.locateCenterOnScreen('42.PNG', confidence=0.92, region=(x, y, 68, 34))
    cond_num_43 = pag.locateCenterOnScreen('43.PNG', confidence=0.92, region=(x, y, 68, 34))
    cond_num_44 = pag.locateCenterOnScreen('44.PNG', confidence=0.92, region=(x, y, 68, 34))
    cond_num_45 = pag.locateCenterOnScreen('45.PNG', confidence=0.92, region=(x, y, 68, 34))
    cond_num_46 = pag.locateCenterOnScreen('46.PNG', confidence=0.92, region=(x, y, 68, 34))
    cond_num_47 = pag.locateCenterOnScreen('47.PNG', confidence=0.92, region=(x, y, 68, 34))
    cond_num_48 = pag.locateCenterOnScreen('48.PNG', confidence=0.92, region=(x, y, 68, 34))
    cond_num_49 = pag.locateCenterOnScreen('49.PNG', confidence=0.92, region=(x, y, 68, 34))
    cond_num_50 = pag.locateCenterOnScreen('50.PNG', confidence=0.92, region=(x, y, 68, 34))
    cond_num_51 = pag.locateCenterOnScreen('51.PNG', confidence=0.92, region=(x, y, 68, 34))
    cond_num_52 = pag.locateCenterOnScreen('52.PNG', confidence=0.92, region=(x, y, 68, 34))
    cond_num_53 = pag.locateCenterOnScreen('53.PNG', confidence=0.92, region=(x, y, 68, 34))
    cond_num_54 = pag.locateCenterOnScreen('54.PNG', confidence=0.92, region=(x, y, 68, 34))
    cond_num_55 = pag.locateCenterOnScreen('55.PNG', confidence=0.92, region=(x, y, 68, 34))
    cond_num_56 = pag.locateCenterOnScreen('56.PNG', confidence=0.92, region=(x, y, 68, 34))
    cond_num_57 = pag.locateCenterOnScreen('57.PNG', confidence=0.92, region=(x, y, 68, 34))
    cond_num_58 = pag.locateCenterOnScreen('58.PNG', confidence=0.92, region=(x, y, 68, 34))
    cond_num_59 = pag.locateCenterOnScreen('59.PNG', confidence=0.92, region=(x, y, 68, 34))
    cond_num_60 = pag.locateCenterOnScreen('60.PNG', confidence=0.92, region=(x, y, 68, 34))
    cond_num_61 = pag.locateCenterOnScreen('61.PNG', confidence=0.92, region=(x, y, 68, 34))
    cond_num_62 = pag.locateCenterOnScreen('62.PNG', confidence=0.92, region=(x, y, 68, 34))
    cond_num_63 = pag.locateCenterOnScreen('63.PNG', confidence=0.92, region=(x, y, 68, 34))
    cond_num_64 = pag.locateCenterOnScreen('64.PNG', confidence=0.92, region=(x, y, 68, 34))
    cond_num_65 = pag.locateCenterOnScreen('65.PNG', confidence=0.92, region=(x, y, 68, 34))
    cond_num_66 = pag.locateCenterOnScreen('66.PNG', confidence=0.92, region=(x, y, 68, 34))
    cond_num_67 = pag.locateCenterOnScreen('67.PNG', confidence=0.92, region=(x, y, 68, 34))
    cond_num_68 = pag.locateCenterOnScreen('68.PNG', confidence=0.92, region=(x, y, 68, 34))
    cond_num_69 = pag.locateCenterOnScreen('69.PNG', confidence=0.92, region=(x, y, 68, 34))
    cond_num_70 = pag.locateCenterOnScreen('70.PNG', confidence=0.92, region=(x, y, 68, 34))
    cond_num_71 = pag.locateCenterOnScreen('71.PNG', confidence=0.92, region=(x, y, 68, 34))
    cond_num_72 = pag.locateCenterOnScreen('72.PNG', confidence=0.92, region=(x, y, 68, 34))
    cond_num_73 = pag.locateCenterOnScreen('73.PNG', confidence=0.92, region=(x, y, 68, 34))
    cond_num_74 = pag.locateCenterOnScreen('74.PNG', confidence=0.92, region=(x, y, 68, 34))
    cond_num_75 = pag.locateCenterOnScreen('75.PNG', confidence=0.92, region=(x, y, 68, 34))
    cond_num_76 = pag.locateCenterOnScreen('76.PNG', confidence=0.92, region=(x, y, 68, 34))
    cond_num_77 = pag.locateCenterOnScreen('77.PNG', confidence=0.92, region=(x, y, 68, 34))

    if (cond_num_1):
        return 1
    if (cond_num_2):
        return 2
    if (cond_num_3):
        return 3
    if (cond_num_4):
        return 4
    if (cond_num_5):
        return 5
    if (cond_num_6):
        return 6
    if (cond_num_7):
        return 7
    if (cond_num_8):
        return 8
    if (cond_num_9):
        return 9
    if (cond_num_10):
        return 10
    if (cond_num_11):
        return 11
    if (cond_num_12):
        return 12
    if (cond_num_13):
        return 13
    if (cond_num_14):
        return 14
    if (cond_num_15):
        return 15
    if (cond_num_16):
        return 16
    if (cond_num_17):
        return 17
    if (cond_num_18):
        return 18
    if (cond_num_19):
        return 19
    if (cond_num_20):
        return 20
    if (cond_num_21):
        return 21
    if (cond_num_22):
        return 22
    if (cond_num_23):
        return 23
    if (cond_num_24):
        return 24
    if (cond_num_25):
        return 25
    if (cond_num_26):
        return 26
    if (cond_num_27):
        return 27
    if (cond_num_28):
        return 28
    if (cond_num_29):
        return 29
    if (cond_num_30):
        return 30
    if (cond_num_31):
        return 31
    if (cond_num_32):
        return 32
    if (cond_num_33):
        return 33
    if (cond_num_34):
        return 34
    if (cond_num_35):
        return 35
    if (cond_num_36):
        return 36
    if (cond_num_37):
        return 37
    if (cond_num_38):
        return 38
    if (cond_num_39):
        return 39
    if (cond_num_40):
        return 40
    if (cond_num_41):
        return 41
    if (cond_num_42):
        return 42
    if (cond_num_43):
        return 43
    if (cond_num_44):
        return 44
    if (cond_num_45):
        return 45
    if (cond_num_46):
        return 46
    if (cond_num_47):
        return 47
    if (cond_num_48):
        return 48
    if (cond_num_49):
        return 49
    if (cond_num_50):
        return 50
    if (cond_num_51):
        return 51
    if (cond_num_52):
        return 52
    if (cond_num_53):
        return 53
    if (cond_num_54):
        return 54
    if (cond_num_55):
        return 55
    if (cond_num_56):
        return 56
    if (cond_num_57):
        return 57
    if (cond_num_58):
        return 58
    if (cond_num_59):
        return 59
    if (cond_num_60):
        return 60
    if (cond_num_61):
        return 61
    if (cond_num_62):
        return 62
    if (cond_num_63):
        return 63
    if (cond_num_64):
        return 64
    if (cond_num_65):
        return 65
    if (cond_num_66):
        return 66
    if (cond_num_67):
        return 67
    if (cond_num_68):
        return 68
    if (cond_num_69):
        return 69
    if (cond_num_70):
        return 70
    if (cond_num_71):
        return 71
    if (cond_num_72):
        return 72
    if (cond_num_73):
        return 73
    if (cond_num_74):
        return 74
    if (cond_num_75):
        return 75
    if (cond_num_76):
        return 76
    if (cond_num_77):
        return 77
    else:
        return 0

def month_check():
    month1 = pag.locateCenterOnScreen('m1.PNG', confidence=0.95, region=(1600, 52, 326, 245))
    month2 = pag.locateCenterOnScreen('m2.PNG', confidence=0.95, region=(1600, 52, 326, 245))
    month3 = pag.locateCenterOnScreen('m3.PNG', confidence=0.95, region=(1600, 52, 326, 245))
    month4 = pag.locateCenterOnScreen('m4.PNG', confidence=0.95, region=(1600, 52, 326, 245))
    month5 = pag.locateCenterOnScreen('m5.PNG', confidence=0.95, region=(1600, 52, 326, 245))
    month6 = pag.locateCenterOnScreen('m6.PNG', confidence=0.95, region=(1600, 52, 326, 245))
    month7 = pag.locateCenterOnScreen('m7.PNG', confidence=0.95, region=(1600, 52, 326, 245))
    month8 = pag.locateCenterOnScreen('m8.PNG', confidence=0.95, region=(1600, 52, 326, 245))
    month9 = pag.locateCenterOnScreen('m9.PNG', confidence=0.95, region=(1600, 52, 326, 245))
    month10 = pag.locateCenterOnScreen('m10.PNG', confidence=0.95, region=(1600, 52, 326, 245))
    month11 = pag.locateCenterOnScreen('m11.PNG', confidence=0.95, region=(1600, 52, 326, 245))
    month12 = pag.locateCenterOnScreen('m12.PNG', confidence=0.95, region=(1600, 52, 326, 245))
    if month1:
        return 1
    elif month2:
        return 2
    elif month3:
        return 3
    elif month4:
        return 4
    elif month5:
        return 5
    elif month6:
        return 6
    elif month7:
        return 7
    elif month8:
        return 8
    elif month9:
        return 9
    elif month10:
        return 10
    elif month11:
        return 11
    elif month12:
        return 12
    else:
        return 0


while True:                # 여기서부터 돌립니다
    if keyboard.is_pressed('END'):
        break

    position_check()  # 위에 회색 선  생기지 않았는지 확인

    # 하루만 돌리고 싶다
    if (one_day_only):
        days_to_go = 1
    # 여러 날 돌리고 싶다
    else:
        start_date = get_date(yy, mm, dd)
        end_date = get_date(end_yy, end_mm, end_dd)
        print('start_date:', start_date.strftime('%Y-%m-%d'), ', end_date:', end_date.strftime('%Y-%m-%d'))
        days_to_go = (end_date - start_date).days
        days_to_go = days_to_go + 1
        # print('days_to_go:', days_to_go + 1)

    for i in range(0, days_to_go, 1):   # 원하는 날짜만큼 돌립니다, days_to_go에 1더해야하나? 그냥해도 되나?
        if keyboard.is_pressed('end'):
            print('end 눌러 종료합니다.')
            break
        # 달력의 년, 월 확인
        year2021 = pag.locateCenterOnScreen('year2021.PNG', confidence=0.92)
        year2022 = pag.locateCenterOnScreen('year2022.PNG', confidence=0.92)
        year2023 = pag.locateCenterOnScreen('year2023.PNG', confidence=0.92)

        if (yy == 2022):
            if (year2022):
                print('')
            elif (year2021):
                pag.click(1869, 93)
                pag.moveTo(1758, 63)
                t.sleep(0.5)
            elif (year2023):
                pag.click(1663, 98)
                pag.moveTo(1758, 63)
                t.sleep(0.5)
            else:
                print('error! 년도가 안맞아요')
                break

        k = month_check()
        # print('k:', k, 'mm', mm)
        if not (k == mm):
            pag.click(1838, 95) # > click
            pag.moveTo(1758, 63)
            t.sleep(0.5)
            k = month_check()
            t.sleep(0.5)
            if not (k == mm):
                print('T_T')
                break

        print('기준(현재)날짜:', get_date(yy, mm, dd).strftime('%Y-%m-%d'))
        # what_week = get_week_no(yy, mm, dd)
        what_week = get_week_of_month(yy, mm, dd)
        day_day = get_date(yy, mm, dd).weekday()
        # print(what_week-1, '번째 주의 ', day_day, '째날이예요')
        click_date(day_day, what_week-1)

        t.sleep(0.5)

        dur_clear()
        click_pat()
        dur_clear()

        # 달의 마지막 날짜면 ... 다음달로!
        # date = datetime(year=yy, month=mm, day=dd).date()
        date = get_date(yy, mm, dd)
        # print('date: ', date.strftime('%Y-%m-%d'))
        last_day = calendar.monthrange(date.year, date.month)[1]
        # print('last_day:', last_day)
        last_day1 = get_date(yy, mm, last_day)
        # print('last_day1:', last_day1.strftime('%Y-%m-%d'))

        if (date == last_day1):

            if((mm+1)%12 == 0):
                mm = 12
                dd = 1
            elif ((mm+1)%12 == 1):
                yy = yy + 1
                mm = (mm+1)%12
                dd = 1
            else:
                mm = (mm+1)%12
                dd = 1
        else:
            dd = dd + 1

        print('다음 날짜로 진행합니다.')
    print('사이클 완료!')
    break
