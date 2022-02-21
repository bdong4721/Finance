import xlrd
import pandas as pd
import pyautogui as pg
import telegram
import datetime
import time
import sys
from PIL import ImageGrab
import numpy as np

sys.setrecursionlimit(100000)
screen = ImageGrab.grab()

tick = 0.25
NUM_FUTURE = 1



def one_min():
    current = datetime.datetime.now()
    current = current.replace(second=0)
    ret = current + datetime.timedelta(minutes=1)
    ret = time.mktime(ret.timetuple())
    return ret


def Get_data():
    try:
        book = xlrd.open_workbook('C:\\Users\\DeepVisions1\\Documents\\Mini NASDAQ 100(22.03).xls', encoding_override='cp949')
    except:
        time.sleep(0.3)
        book = xlrd.open_workbook('C:\\Users\\DeepVisions1\\Documents\\Mini NASDAQ 100(22.03).xls',encoding_override='cp949')
    worksheet = book.sheet_by_index(0)
    nrows = worksheet.nrows
    row_val = [worksheet.row_values(row_num) for row_num in range(nrows)]
    data = pd.DataFrame(row_val[1:], columns=['날짜', '시간', '시가', '고가', '저가', '종가', '5MA', '20MA', '60MA', 'SPD'])
    return data




def OCO_exe(is_original, is_BUYoco, min_list, max_list, multiple_num):
    if is_original:
        if is_BUYoco:
            if long_profit:
                numtick = abs(round(max(max_list), 2) - stopprice_min) / 0.25
                tick = round(numtick * multiple_num, 0)
                half = stopprice_min - tick * 0.25
                oco_low = round(half)
            else:
                # 매수로 익절(limit)
                oco_low = round((2 * min(min_list) - max(max_list)), 2)


            # 매수로 손절(limit)
            oco_high = round(max(max_list), 2)
            print("oco_high : ",oco_high, "oco_low : ", oco_low)
            print('매도 -> 매수 OCO')
        else:
            if long_profit:
                numtick = abs(round(min(min_list), 2) - stopprice_max) / 0.25
                tick = round(numtick * multiple_num, 0)
                half = stopprice_max + tick * 0.25
                oco_high = round(half)
            else:
                # 매도로 익절(limit)
                oco_high = round((2 * max(max_list) - min(min_list)), 2)


            # 매도로 손절(limit)
            oco_low = round(min(min_list), 2)
            print("oco_high : ", oco_high, "oco_low : ", oco_low)
            print('매수 -> 매도 OCO')
    else:
        if is_BUYoco:
            time.sleep(1)
            pg.click(x=510, y=100, button='left')
            ## 매도 -> 매수 OCO
            pg.click(x=462, y=169, button='left')
            time.sleep(1)  ## 매수 OCO 상단
            pg.click(x=553, y=225, button='left')  # 스크롤 선택
            time.sleep(1)
            pg.click(x=467, y=313, button='left')  # OCO선택
            time.sleep(1)
            pg.click(x=509, y=250, button='left')
            time.sleep(1)  # 계약수 지정
            pg.write(str(NUM_FUTURE));
            time.sleep(1)  # 매수로 익절
            pg.click(x=513, y=306, button='left')
            time.sleep(1)
            print(min_list);
            print(max_list)

            # 매수로 익절(limit)
            if half_profit:
                numtick = abs(round(min(min_list), 2) - stopprice_max) / 0.25
                halftick = round(numtick / 2, 0)
                half = stopprice_max - halftick * 0.25
                pg.write(str(half));
                oco_high = round(half)
            else:
                pg.write(str(round(min(min_list), 2)))
                oco_high = round(min(min_list), 2)

            time.sleep(1)
            pg.click(x=510, y=333, button='left')
            time.sleep(1)

            # 매수로 손절(limit)
            pg.write(str(round((2 * max(max_list) - min(min_list)), 2)));
            oco_low = round((2 * max(max_list) - min(min_list)), 2)

            time.sleep(1)
            pg.click(x=451, y=366, button='left')  # 매수 클릭
            time.sleep(1)
            print('매도 -> 매수 OCO')
        else:
            time.sleep(1)
            pg.click(x=505, y=103, button='left')
            ## 매수 -> 매도 OCO
            pg.click(x=601, y=171, button='left')
            time.sleep(1)  ## 매도 OCO 상단
            pg.click(x=553, y=225, button='left')  # 스크롤 선택
            time.sleep(1)
            pg.click(x=467, y=313, button='left')  # OCO선택
            time.sleep(1)
            pg.click(x=526, y=254, button='left')
            time.sleep(1)  # 계약수 지정
            pg.write(str(NUM_FUTURE))
            time.sleep(1)
            pg.click(x=513, y=306, button='left')
            time.sleep(1)
            print(min_list)
            print(max_list)

            print("stopprice_min으로 진입후 oco")
            # 매도로 익절(limit)
            if half_profit:
                numtick = (round(max(max_list), 2) - stopprice_min) / 0.25
                halftick = round(numtick / 2, 0)
                half = stopprice_min + halftick * 0.25
                pg.write(str(half));
                oco_low = round(half)
            else:
                pg.write(str(round(max(max_list), 2)))
                oco_low = round(max(max_list), 2)

            time.sleep(1)
            pg.click(x=510, y=332, button='left');
            time.sleep(1)

            # 매도로 손절(limit)
            pg.write(str(round((2 * min(min_list) - max(max_list)), 2)))
            oco_high_local = round((2 * min(min_list) - max(max_list)), 2)
            time.sleep(1)
            pg.click(x=456, y=365, button='left')  # 매수 클릭
            time.sleep(1)
            print('=====', time.ctime(), '=====')
            print('매수 -> 매도 OCO')
    return [oco_low, oco_high]




def Updatingdat(time_iter):
    # while True:
    data_len = len(data)
    new_data = data[(data_len-23-time_iter):data_len]
    # if new_data.iloc[2, 1] == old_data.iloc[1, 1]:
    global Oprice_0, Cprice_0, Hprice_0, Hprice_1, Hprice_2, Lprice_0, Lprice_1, Lprice_2, MA5_0, MA5_1, MA20_0, MA20_1, \
        MA60_0, MA60_1, MA60_2, SPD_0, SPD_1, TIME_curr_0, DATE_curr_0

    Oprice_0 = new_data.iloc[1, 2]
    Cprice_0 = new_data.iloc[1, 5]

    Hprice_0 = new_data.iloc[1, 3]
    Hprice_1 = new_data.iloc[2, 3]
    Hprice_2 = new_data.iloc[3, 3]

    Lprice_0 = new_data.iloc[1, 4]
    Lprice_1 = new_data.iloc[2, 4]
    Lprice_2 = new_data.iloc[3, 4]

    MA5_0 = new_data.iloc[1, 6]
    MA5_1 = new_data.iloc[2, 6]
    MA20_0 = new_data.iloc[1, 7]
    MA20_1 = new_data.iloc[2, 7]
    MA60_0 = new_data.iloc[1, 8]
    MA60_1 = new_data.iloc[2, 8]
    MA60_2 = new_data.iloc[3, 8]

    SPD_0 = new_data.iloc[1, 9]
    SPD_1 = new_data.iloc[2, 9]

    DATE_curr_0 = new_data.iloc[1, 0]
    TIME_curr_0 = new_data.iloc[1, 1]

    return new_data





def test_1(min_list, max_list, time_iter):
    while True:
        old_data = Updatingdat(time_iter)
        if SPD_1 < 0 < SPD_0:
            max_list = [Hprice_2 + tick, Hprice_1 + tick, Hprice_0 + tick]
            return 'increase', old_data, min_list, max_list, time_iter

        elif SPD_1 > 0 > SPD_0:
            min_list = [Lprice_2 - tick, Lprice_1 - tick, Lprice_0 - tick]
            return 'decrease', old_data, min_list, max_list, time_iter

        else:
            time_iter += 1
            print("Current time = ", time_iter)


def test_2(x, old_data, min_list, max_list, time_iter):
    time_iter = time_iter + 1
    # print(time_iter)

    old_data = Updatingdat(time_iter)
    if x == 'increase':

        if not (SPD_0 == "" or SPD_1 == "") :
            if SPD_0 < 0:
                #print('반대로 크로스 [매도]')
                #print(min_list)
                #print(max_list)
                return test_2(x='decrease', old_data=old_data, min_list=[Lprice_2 - tick, Lprice_1 - tick, Lprice_0 - tick],
                              max_list=[], time_iter=time_iter)

            elif SPD_1 > SPD_0:
                #print('스프레드 증가 -> 감소 [매수] : 스프레드 꺾임 성공')
                #print(min_list)
                min_list = []
                min_list.append(Lprice_0 - tick)
                #print(min_list)

                global stopprice_max
                stopprice_max = max(max_list)

                #STOP_BUYorSELL(is_original, is_original, max_list, min_list)

                if stopprice_max <= max(max_list):
                    #print('스프레드 증가 -> 감소 [매수], 진입가 : ' + str(stopprice_max) + '\n')
                    return test_3(x='increase', old_data=old_data, min_list=min_list, max_list=max_list, alarm=1, time_iter=time_iter, TOTAL_REWARD = TOTAL_REWARD)
                else:
                    return test_3(x='increase', old_data=old_data, min_list=min_list, max_list=max_list, alarm=0, time_iter=time_iter, TOTAL_REWARD = TOTAL_REWARD)

            else:
                #print('스프레드 꺾이는 것 대기 [매수] \n')
                max_list.append(Hprice_0 + tick)
                #print(max_list)
                #print(min_list)
                return test_2(x='increase', old_data=old_data, min_list=min_list, max_list=max_list, time_iter=time_iter)
        else:
            max_list.append(Hprice_0 + tick)
            # print(max_list)
            # print(min_list)
            return test_2(x='increase', old_data=old_data, min_list=min_list, max_list=max_list, time_iter=time_iter)

    else:

        if not (SPD_0 == "" or SPD_1 == ""):
            if SPD_0 > 0:
                #print('반대로 크로스 [매수]')
                return test_2(x='increase', old_data=old_data, min_list=min_list,
                              max_list=[Hpㅇice_2 + tick, Hprice_1 + tick, Hprice_0 + tick], time_iter=time_iter)

            elif SPD_1 < SPD_0:
                #print('스프레드 증가 -> 감소 [매도] : 스프레드 꺾임 성공')

                max_list = []
                max_list.append(Hprice_0 + tick)
                # max_list.append(Hprice_1 + tick)
                global stopprice_min
                stopprice_min = min(min_list)

                #STOP_BUYorSELL(is_original, not is_original, max_list, min_list)

                if stopprice_min >= min(min_list):
                    return test_3(x='decrease', old_data=old_data, min_list=min_list, max_list=max_list,
                                  alarm=1, time_iter=time_iter, TOTAL_REWARD = TOTAL_REWARD)  # 아직 진입가도달 x
                else:
                    return test_3(x='decrease', old_data=old_data, min_list=min_list, max_list=max_list,
                                  alarm=0, time_iter=time_iter, TOTAL_REWARD = TOTAL_REWARD)  # alarm=0 이미 진입가 도달

            else:
                # print('스프레드 꺾이는 것 대기 [매도]')
                min_list.append(Lprice_0 - tick)
                return test_2(x='decrease', old_data=old_data, min_list=min_list, max_list=max_list, time_iter=time_iter)
        else:
            # print('스프레드 꺾이는 것 대기 [매도]')
            min_list.append(Lprice_0 - tick)
            return test_2(x='decrease', old_data=old_data, min_list=min_list, max_list=max_list, time_iter=time_iter)


def test_3(x, old_data, min_list, max_list, alarm, time_iter,TOTAL_REWARD):
    time_iter = time_iter + 1
    old_data = Updatingdat(time_iter)
    if x == 'increase':
        if not SPD_0 == "":
            if SPD_0 < 0:
                #print('반대로 크로스 [매도]')
                return 'decrease', old_data, [Lprice_2 - tick, Lprice_1 - tick, Lprice_0 - tick], [], time_iter
            elif Hprice_0 > max(max_list) and alarm == 1:
                if MA5_0 > MA20_0 and MA20_0 > MA60_0:
                    #print('진입 완료 [매수] ' + Time_curr_0);
                    print('===== 스프레드 ===== [매수] 진입가 : ' + str(round(max(max_list), 2)) + ' 익절가 : ' + str(
                            round(min(min_list), 2)) + ' 손익절 : ' + str(
                            round((max(max_list) - min(min_list)) / tick, 2)) + ' 틱 \n')
                    TOTAL_TIME.append(TIME_curr_0)
                    TOTAL_DATE.append(DATE_curr_0)
                    # reverse일때는 True로 해야지 OCO매수
                    # 여기서 두번째는 is_BUY

                    oco_array = OCO_exe(is_original, not is_original, min_list, max_list, multiple_num)
                    # original 정방향시 매도 OCO
                    oco_low = oco_array[0]
                    oco_high = oco_array[1]
                    # OCO 에러체크
                    while not ((Lprice_0 < oco_low) or (Hprice_0 > oco_high)):
                        #print("손절 또는 익절 체킹 시작")
                        time_iter = time_iter + 1
                        #print(time_iter)
                        Updatingdat(time_iter)
                        # print("oco_low",oco_low);print("oco_high", oco_high);print(Lprice_0);print(Hprice_0);print(min_list);print(max_list)
                        # print("손절 또는 익절 체킹중")

                    print("OCO탈출", time_iter)
                    #TOTAL_TRANJ += 1

                    if Lprice_0 < oco_low and Hprice_0 > oco_high:  # 위아래로 찌른경우
                        TOTAL_REWARD.append(-abs(max(max_list) - oco_low))
                    elif Hprice_0 > oco_high:  # 매수진입 -> OCO 매도 익절
                        TOTAL_REWARD.append(abs(max(max_list)-oco_high))
                    else: #Lprice_0 < oco_low :    # 매수진입 -> OCO 매도 손절
                        TOTAL_REWARD.append(-abs(max(max_list)-oco_low))

                    alarm = 0
                    #print("TOTAL_REWARD : ", TOTAL_REWARD)
                else:
                    alarm = 0

            else:
                max_list.append(Hprice_0 + tick)
                min_list.append(Lprice_0 - tick)
        else:
         max_list.append(Hprice_0 + tick)
         min_list.append(Lprice_0 - tick)

    else:
        #print('매수 크로스 또는 매도 진입 대기중2')

        if not SPD_0 == "":
            if SPD_0 > 0:
                #print('반대로 크로스 [매수]')
                return 'increase', old_data, min_list, [Hprice_2 + tick, Hprice_1 + tick, Hprice_0 + tick], time_iter

            elif Lprice_0 < min(min_list) and alarm == 1:
                if MA5_0 < MA20_0 and MA20_0 < MA60_0:
                    print('** ' + '===== 스프레드 ===== [매도 진입 완료] 진입가 : ' + str(round(min(min_list), 2)) + ' 손절가 : ' + str(
                        round(max(max_list), 2)) + ' 손익절 : ' + str(
                        round((max(max_list) - min(min_list)) / tick, 2)) + ' 틱\n')
                    print('진입 완료 [매도] ' + TIME_curr_0 )
                    TOTAL_TIME.append(TIME_curr_0)
                    TOTAL_DATE.append(DATE_curr_0)
                    # reverse일때는 False로 해야지 OCO매도
                    oco_array = OCO_exe(is_original, is_original, min_list, max_list, multiple_num)
                    oco_low = oco_array[0]
                    oco_high = oco_array[1]
                    # original 정방향시 매수 OCO

                    # OCO로 손절 또는 익절 청산이 나는지 아닌지 확인하는 단계
                    while not ((Lprice_0 < oco_low) or (Hprice_0 > oco_high)):
                        time_iter = time_iter + 1
                        #print(time_iter)
                        Updatingdat(time_iter)
                        # print("oco_low",oco_low);print("oco_high", oco_high);
                        # print("손절 또는 익절 체킹시작")

                    print("OCO탈출", time_iter)
                    #TOTAL_TRANJ += 1
                    if Lprice_0 < oco_low and Hprice_0 > oco_high : #위아래로 찌른경우
                        TOTAL_REWARD.append(-abs(min(min_list) - oco_high))
                    elif Lprice_0 < oco_low :   # 매도진입 -> OCO 매수 익절
                        TOTAL_REWARD.append(abs(min(min_list)-oco_low))
                    else: #Hprice_0 > oco_high :  # 매도진입 -> OCO 매수 손절
                        TOTAL_REWARD.append(-abs(min(min_list)-oco_high))
                    alarm = 0
                else:
                    alarm = 0
            else:
                min_list.append(Lprice_0 - tick)
                max_list.append(Hprice_0 + tick)
                #print(min_list)
                #print(max_list)
                # print("time_iter : ", time_iter)

        else:
            min_list.append(Lprice_0 - tick)
            max_list.append(Hprice_0 + tick)
            # print(min_list)
            # print(max_list)
            #print("time_iter : ", time_iter)

    return test_3(x, old_data, min_list, max_list, alarm, time_iter, TOTAL_REWARD)


long_profit = True
half_profit = False  # True는 절반 익절
is_original = True  # False는 리버스 매매
global time_iter
global TOTAL_REWARD
global TOTAL_TIME
global TOTAL_DATE

min_list = []
max_list = []
TOTAL_REWARD = []
TOTAL_TIME = []
TOTAL_DATE = []

time_iter=10
multiple_num = 2
data = Get_data()
x, old_data, min_list, max_list, time_iter = test_1(min_list, max_list, time_iter)

for i in range(len(data)):
    x, old_data, min_list, max_list, time_iter = test_2(x, old_data, min_list=min_list, max_list=max_list, time_iter=time_iter)

len(TOTAL_REWARD)
len(TOTAL_TIME)


# TOTAL_REWARD.append(0)
tmp=np.array(TOTAL_REWARD)
pos_ind = tmp > 0

nonmul_profit = tmp
nonmul_profit[pos_ind] = nonmul_profit[pos_ind]/multiple_num
nonmul_profit = abs(nonmul_profit)

winorlose = np.array(np.array(TOTAL_REWARD) > 0, dtype=int).tolist()

TOTAL = pd.DataFrame({ "date" : TOTAL_DATE, "time" :TOTAL_TIME, "reward" : TOTAL_REWARD, "손절" : list(nonmul_profit), "승패": winorlose})
TOTAL.to_excel("C:\\Users\\DeepVisions1\\Dropbox\\PC\\Documents\\0_선물분석\\output.xlsx")



sel_ind = abs(nonmul_profit) < 40
sel_ind = np.logical_and(60 < abs(nonmul_profit) , abs(nonmul_profit) < 100)




list(tmp[sel_ind])
sum(tmp[sel_ind])
len(tmp[sel_ind])

np.logical_and([True, False, False, True, True], [True, True, False, True, False])
[True, False, False, True, True]






list(nonmul_profit)
tmp[pos_ind] = tmp[pos_ind]/2
tmp[tmp > 0] = tmp[tmp > 0]
tmp = tmp[abs(tmp) < 40 ]
sum(tmp[abs(tmp) < 70])

size = 50
sum(tmp[abs(tmp) > size]) - len(tmp[abs(tmp) > size])*10

ttmp = 0
cumtmp = []
i = 0
for i in range(len( tmp[  abs(tmp) > size    ])-1):
    ttmp = ttmp + tmp[  abs(tmp) > size    ][i]
    cumtmp.append(ttmp)

    min(cumtmp)

    tmp1 = np.sort(tmp[  abs(tmp) > size    ])
