import xlrd,xlwt
from xlutils.copy import copy   #导入复制模块

#返回当前工作路径
# import os
# def path():
#     return os.getcwd()
# print(path())

#-----------------自定义函数--------------------------
#处理时间格式数据
def time_data_deal(str):
    p = str.find(":")
    l1 = str[0:p]
    l2 = str[p+1:]
    hour = int(l1)
    min = round(int(l2) / 60, 2)     #使用round(num,2)函数保留两位小数
    time = hour + min
    return time

#加班时间处理，0.5h起步
def overtimedeal(num):
    if num < 0.5:
        num = 0
    elif (num >= 0.5) and (num < 1):
        num = 0.5
    elif (num >= 1) and (num < 1.5):
        num = 1
    elif (num >= 1.5) and (num < 2):
        num = 1.5
    elif (num >= 2) and (num < 2.5):
        num = 2
    elif (num >= 2.5) and (num < 3):
        num = 2.5
    elif (num >= 3) and (num < 3.5):
        num = 3
    elif (num >= 3.5) and (num < 4):
        num = 3.5
    elif (num >= 4) and (num < 4.5):
        num = 4
    elif (num >= 4.5) and (num < 5):
        num = 4.5
    elif (num >= 5) and (num < 5.5):
        num = 5
    elif (num >= 5.5) and (num < 6):
        num = 5.5
    elif (num >= 6) and (num < 6.5):
        num = 6
    elif (num >= 6.5) and (num < 7):
        num = 7
    elif (num >= 7) and (num < 7.5):
        num = 7
    elif (num >= 7.5) and (num < 8):
        num = 7.5
    elif (num >= 8) and (num < 8.5):
        num = 8
    elif (num >= 8.5) and (num < 9):
        num = 8.5
    elif (num >= 9) and (num < 9.5):
        num = 9
    elif (num >= 9.5) and (num < 10):
        num = 9.5
    elif (num >= 10) and (num < 10.5):
        num = 10
    elif (num >= 10.5) and (num < 11):
        num = 10.5
    elif (num >= 11) and (num < 11.5):
        num = 11
    elif (num >= 11.5) and (num < 12):
        num = 11.5
    elif (num >= 12) and (num < 12.5):
        num = 12
    elif (num >= 12.5) and (num < 13):
        num = 12.5
    elif (num >= 13) and (num < 13.5):
        num = 13
    elif (num >= 13.5) and (num < 14):
        num = 13.5
    elif (num >= 14) and (num < 14.5):
        num = 14
    elif (num >= 14.5) and (num < 15):
        num = 14.5
    elif (num >= 15) and (num < 15.5):
        num = 15
    else:
        num = 15
    return(num)
#-----------------自定义函数--------------------------
wb = xlrd.open_workbook('考勤表.xls')
ws = wb.sheet_by_name('考勤数据')
# nwb = xlwt.Workbook(encoding='utf-8')     #在原表中创建新的Sheet
# nws = nwb.add_sheet('加班统计结果')
nwb = copy(wb)  #复制读取工作簿
nws = nwb.get_sheet('考勤数据')
nws2 = nwb.add_sheet('研发中心人员考勤统计结果')
namelist = ['徐海霆','李卫','肖岩','陈雅雪','陈海林','王俊','潘佩华','陆小彬','戴宇文','蔡囯章','卢志铭','赵淑强','周海舰',
            '蒋卓洪','何云山','冯俊','李彬','丘祥观','陈艳艳','黄南熙','罗嘉基','黄志勇','袁焕伦','罗曦','梁隽',
            '霍汉明','曾湘','赵乐飞','杨超群','胡红星','刘日平','黄展钊']
#-------------------在新的Sheet中加入表头------------------------
nws2.write(0, 0, ws.cell_value(0, 0))
nws2.write(0, 1, ws.cell_value(0, 1))
nws2.write(0, 2, ws.cell_value(0, 2))
nws2.write(0, 3, ws.cell_value(0, 3))
nws2.write(0, 4, ws.cell_value(0, 4))
nws2.write(0, 5, ws.cell_value(0, 5))
nws2.write(0, 6, ws.cell_value(0, 6))
nws2.write(0, 7, ws.cell_value(0, 7))
nws2.write(0, 8, ws.cell_value(0, 8))
nws2.write(0, 9, ws.cell_value(0, 9))
nws2.write(0, 10, ws.cell_value(0, 10))
nws2.write(0, 11, ws.cell_value(0, 11))
nws2.write(0, 12, '工作日加班（h）')
nws2.write(0, 13, '周末加班（h）')
nws2.write(0, 14, '备注')

r = 0
timedatafilter = []
workinghoursdatacount = []
while r < ws.nrows-1:
    r += 1
    # -------------------将人员信息和打卡数据复制到新的Sheet中------------------------
    if ws.cell_value(r, 1) in namelist:
        nws2.write(r, 0, ws.cell_value(r, 0))
        nws2.write(r, 1, ws.cell_value(r, 1))
        nws2.write(r, 2, ws.cell_value(r, 2))
        nws2.write(r, 3, ws.cell_value(r, 3))
        nws2.write(r, 4, ws.cell_value(r, 4))
        nws2.write(r, 5, ws.cell_value(r, 5))
        nws2.write(r, 6, ws.cell_value(r, 6))
        nws2.write(r, 7, ws.cell_value(r, 7))
        nws2.write(r, 8, ws.cell_value(r, 8))
        nws2.write(r, 9, ws.cell_value(r, 9))
        nws2.write(r, 10, ws.cell_value(r, 10))
        nws2.write(r, 11, ws.cell_value(r, 11))

    row_test = ws.row_values(r)[5:12]
    timedatafilter = []
    workinghoursdatacount = []
    # -------------------列表中删除空字符串，并使用函数转化成时间数值格式--------------------
    for txt in row_test:
        if txt != '':
            numT = time_data_deal(txt)
            timedatafilter.append(numT)
    for numT1 in timedatafilter:
        if (numT1 > 9 and numT1 < 12) or (numT1 > 13 and numT1 < 17.25):
            workinghoursdatacount.append(numT1)     #统计正常工作时间内的打卡，用于是否有请假的判断
            # print(workinghoursdatacount)
    if ws.cell_value(r, 1) in namelist:
        if ws.cell_value(r,3) in ['星期一','星期二','星期三','星期四','星期五']:
            if len(timedatafilter) == 0:
                nws2.write(r, 14, '无打卡数据')
            elif len(timedatafilter) == 1:
                nws2.write(r, 14, '忘打卡')
            elif (len(timedatafilter) >= 2) and (len(workinghoursdatacount) <= 1):
                if timedatafilter[0] < 8.25:
                    timedatafilter[0] = 8.25
                if timedatafilter[-1] - timedatafilter[0] >= 9:
                    overtime_weekday = timedatafilter[-1] - timedatafilter[0] - 9.5
                    overtime_weekday_after = overtimedeal(overtime_weekday)
                    nws2.write(r, 12, overtime_weekday_after)
                    if timedatafilter[0] <= 9:
                        nws2.write(r, 14, '正常')
                    else:
                        nws2.write(r, 14, '迟到，但工作时长够')
                else:
                    nws2.write(r, 14, '工作时长不足8小时')
            else:
                nws2.write(r, 14, '工作时段打卡次数超过一次，请核实请假情况')

        elif ws.cell_value(r,3) in ['星期六','星期日']:
            if len(timedatafilter) == 0:
                nws2.write(r, 14, '未加班')
            elif len(timedatafilter) == 1:
                nws2.write(r, 14, '忘打卡')
            elif len(timedatafilter) >= 2:
                if timedatafilter[0] < 12:
                    overtime_weekend = timedatafilter[-1] - timedatafilter[0] - 1
                else:
                    overtime_weekend = timedatafilter[-1] - timedatafilter[0]
                overtime_weekend_after = overtimedeal(overtime_weekend)
                nws2.write(r, 13, overtime_weekend_after)
                nws2.write(r, 14, '正常')

    # print(len(timedatafilter))
    # print(ws.cell_value(r,1),ws.cell_value(r,3),timedatafilter)
print('处理完毕')
# -------------------------------------------------------------------------------
nwb.save('研发中心人员考勤统计结果.xls')


