#!/usr/bin/env python
#_*_ coding:utf-8 _*_
#Create by qilin 2018-8-15 Version V1.0
#Filename: WorkAttendance.py
import os

import xlrd
import xlwt
import datetime
import calendar

from xlwt.Workbook import Workbook
from xlrd import open_workbook
from xlutils.copy import copy
from calendar import calendar

print '程序开始运行'
#获取3个excel表
location = 'D:\kqtj\\'
book_01 = xlrd.open_workbook(location + '01.xls')
book_oaqj = xlrd.open_workbook(location + 'oaqj.xls')
book_oanj = xlrd.open_workbook(location + 'oanj.xls')

#拷贝创建01表形成一个新表
today = datetime.date.today()
today = bytes(today)
book = open_workbook(location + '01.xls')
wbook = copy(book)
wsheet = wbook.get_sheet(0)


#获取3张表中的sheet
sheet_01 = book_01.sheets()[0]
sheet_oaqj = book_oaqj.sheets()[0]
sheet_oanj = book_oanj.sheets()[0]


#01和oaqj两张表对比，如果时间相同，并且姓名相同，就把该行的请假信息合并到01表对应的行中
for i in range(sheet_oaqj.nrows):
    #qjts=请假天数，qjlx=请假类型
    qjts = sheet_oaqj.cell_value(i,4)          
    qjlx = sheet_oaqj.cell_value(i,3)
    #表oa姓名列
    name_oaqj = sheet_oaqj.cell(i,0).value.encode('utf-8')
    #OA请假表中的开始日期
    start_date = sheet_oaqj.cell(i,1).value.encode('utf-8')


    #oa请假表中的每一条信息与01考勤表里的信息逐条核对  
    for j in range(sheet_01.nrows):
        #表01里的时间列,且为字符串类型
        time_01 = sheet_01.cell(j,3).value.encode('utf-8')
        #表01姓名列
        name_01 = sheet_01.cell(j,1).value.encode('utf-8')


        if start_date == time_01 and name_01 == name_oaqj:

            #把请假类型和请假时间写入到01后续对应的列中
            wsheet.write(j,5,qjts)
            wsheet.write(j,6,qjlx)

#01和oanj两张表对比，如果时间相同，并且姓名相同，就把该行的请假信息合并到01表对应的行中
for i in range(sheet_oanj.nrows):
    #qjts=请假天数，qjlx=请假类型
    qjts = sheet_oanj.cell_value(i,5)          
    qjlx = sheet_oanj.cell_value(i,4)
    #表oa姓名列
    name_oanj = sheet_oanj.cell(i,0).value.encode('utf-8')
    #OA请假表中的开始日期
    start_date = sheet_oanj.cell(i,2).value.encode('utf-8')


    #oa请假表中的每一条信息与01考勤表里的信息逐条核对  
    for j in range(sheet_01.nrows):
        #表01里的时间列,且为字符串类型
        time_01 = sheet_01.cell(j,3).value.encode('utf-8')
        #表01姓名列
        name_01 = sheet_01.cell(j,1).value.encode('utf-8')


        if start_date == time_01 and name_01 == name_oanj:

            #把请假类型和请假时间写入到01后续对应的列中
            wsheet.write(j,5,qjts)
            wsheet.write(j,6,qjlx)

#写入后两列表头
qjts_title = sheet_oaqj.cell_value(0,4)
qjlx_title = sheet_oaqj.cell_value(0,3)
wsheet.write(0,5,qjts_title)
wsheet.write(0,6,qjlx_title)

wbook.save(location + today + '.xls')
print '原始考勤数据合并工作完成！'

#------------------------写入统计表--------------------------
#获取数据合并表
book_combin = xlrd.open_workbook(location + today + '.xls')
sheet_combin = book_combin.sheets()[0]

#获取汇总统计表模板并拷贝形成一个新表
book_yjy = open_workbook('C:\Users\木\Desktop\考勤管理\姚家园.xls'.decode('utf8').encode('gbk'),formatting_info=True)
wbook_yjy = copy(book_yjy)
wsheet_yjy = wbook_yjy.get_sheet(0)

book_dag = open_workbook('C:\Users\木\Desktop\考勤管理\档案馆.xls'.decode('utf8').encode('gbk'),formatting_info=True)
wbook_dag = copy(book_dag)
wsheet_dag = wbook_dag.get_sheet(0)

book_jsc = open_workbook('C:\Users\木\Desktop\考勤管理\技术处.xls'.decode('utf8').encode('gbk'),formatting_info=True)
wbook_jsc = copy(book_jsc)
wsheet_jsc = wbook_jsc.get_sheet(0)

#将数据合并表中的考勤记录信息填入到汇总统计表中
#如果数据合并表中的时间列为空，则对应统计表中该日期下的值就为空
#请假类型与请假天数一并统计，不同类型与特殊字符以字典形式对应，天数累计但要考虑隔开周末的情况

'''
获取当前年，上个月，以参数形式代入，得到上个月的天数
'''

import calendar
from calendar import monthrange

now = datetime.datetime.now()
year = int(now.strftime('%Y'))
month = int(now.strftime('%m'))
month = month - 1
last_day = calendar.monthrange(year,month)[1]

'''
#遍历出合并表中请假类型对应的符号
for i in range(1,sheet_combin.nrows):
    #qjts=请假天数，qjlx=请假类型
    qjts_combin = sheet_combin.cell_value(i,5)          
    qjlx_combin = sheet_combin.cell_value(i,6)

    qjlx_symbol_dict = {
            '正常出勤' : '√',
            '旷工' : '×',
            '迟到' : '△',
            '早退' : '▽',
            '出差' : '□',
            '年假' : '◇',
            '事假' : '○',
            '病假' : '☆',
            '产假' : '▲',
            '半日' : '※',
            '加班' : '#',
            '调休' : '◆',
            '产业园' : '●',
            '': ''
        }
'''

def returnQjlx(rows):
    #定义函数returnQjlx，调用时输入参数“行数”，返回请假类型对应的符号
    qjts_combin = sheet_combin.cell_value(rows,5)          
    qjlx_combin = sheet_combin.cell_value(rows,6)

    qjlx_symbol_dict = {
            '正常出勤' : '√',
            '旷工' : '×',
            '迟到' : '△',
            '早退' : '▽',
            '出差' : '□',
            '年假' : '◇',
            '事假' : '○',
            '病假' : '☆',
            '产假' : '▲',
            '半日' : '※',
            '加班' : '#',
            '调休' : '◆',
            '产业园' : '●',
            '': ''
        }
    qjlx_combin = qjlx_combin.encode("utf8")
    return qjlx_symbol_dict[qjlx_combin]
 
#--------------------------------------统计姚家园考勤---------------------------------   
yjy_list = ['潘新辉','陈石羡','张伊菲','王寒梅','林红路','王利娜','刘佳硕']
h = 6
for i in yjy_list:
    #print 'h= ',h
    l = 2
    for j in range(sheet_combin.nrows): #行数
        
        name_combin = sheet_combin.cell_value(j,1).encode('utf8')
        qjDate = sheet_combin.cell_value(j,3)   #日期
        qjts_combin = sheet_combin.cell_value(j,5).encode('utf8')   #请假天数     
        #qjts_combin = int(qjts_combin)     
        qjlx_combin = sheet_combin.cell_value(j,6).encode('utf8')   #请假类型
        
        qjTime = sheet_combin.cell_value(j,4)   #时间
       
        #判断名字是否相同
        if i == name_combin:
            #判断请假类型是否为空
            '''
            print '请假人：',name_combin
            print 'j= ', j
            print '请假类型= ',qjlx_combin
            '''
            if qjlx_combin != '':
                #判断请假天数是否>1
                if qjts_combin > 1:
                    #根据天数写入，隔开周末
                    qjts_combin = round(float(qjts_combin))
                    qjts_combin = int(qjts_combin)
                    #print '请假天数= ',qjts_combin
                    
                    for k in range(qjts_combin):
                        #判断请假时间是否是周末
                        #print 'k= ',k
                        #print '请假日期= ', qjDate
                   
                        qjDateFormat = qjDate.replace('-', ',')
                        qjDateFormat = qjDateFormat.split(',')
                        qjDateFormat[0] = int(qjDateFormat[0])
                        qjDateFormat[1] = int(qjDateFormat[1])
                        qjDateFormat[2] = int(qjDateFormat[2]) + k
                        #print 'qjDateFormat[2]= ',qjDateFormat[2]
                        if qjDateFormat[2] > last_day:
                            print '本月%s申请了跨月休假' %name_combin
                            break
                        theDay = datetime.datetime(qjDateFormat[0],qjDateFormat[1],qjDateFormat[2]).strftime('%w')
                        #print 'theDay= ',theDay
                        if theDay != 6: #不是周末
                            wsheet_yjy.write(h,l + k,returnQjlx(j))
                            #qjDate
                        elif theDay == 6:   #是周六
                            l += 2
                            wsheet_yjy.write(h,l + k,returnQjlx(j))
                            
                
                elif qjts_combin == 1:
                    #写入请假类型符号
                    wsheet_yjy.write(h,l,returnQjlx(j))
                elif qjts_combin < 1:
                    wsheet_yjy.write(h,l,'※')
            
            elif qjlx_combin == '':
                
                if qjTime != '':         
                #判断时间是否有2个值
                
                    if len(qjTime.split()) > 1:
                        #continue
                        wsheet_yjy.write(h,l,'√')
                    else:
                        #暂定先按半日算，写入'半日'对应符号
                        wsheet_yjy.write(h,l,'※')
                        
            else:
                continue
                
        else:
            continue
        
        l += 1

    h += 1
        
wbook_yjy.save(location + today + 'yjy' + '.xls')

#--------------------------------------统计档案馆考勤---------------------------------
dag_list = ['丁万利','倪艳春','张洪伟','董丽萍','芮民','李鑫鑫','刘云菲','尹凤兰','程驰']
h = 6
for i in dag_list:
    #print 'h= ',h
    l = 2
    for j in range(sheet_combin.nrows): #行数
        
        name_combin = sheet_combin.cell_value(j,1).encode('utf8')
        qjDate = sheet_combin.cell_value(j,3)   #日期
        qjts_combin = sheet_combin.cell_value(j,5).encode('utf8')   #请假天数     
        #qjts_combin = int(qjts_combin)     
        qjlx_combin = sheet_combin.cell_value(j,6).encode('utf8')   #请假类型
        
        qjTime = sheet_combin.cell_value(j,4)   #时间
       
        #判断名字是否相同
        if i == name_combin:
            #判断请假类型是否为空
            '''
            print '请假人：',name_combin
            print 'j= ', j
            print '请假类型= ',qjlx_combin
            '''
            if qjlx_combin != '':
                #判断请假天数是否>1
                if qjts_combin > 1:
                    #根据天数写入，隔开周末
                    qjts_combin = round(float(qjts_combin))
                    qjts_combin = int(qjts_combin)
                    #print '请假天数= ',qjts_combin
                    
                    for k in range(qjts_combin):
                        #判断请假时间是否是周末
                        #print 'k= ',k
                        #print '请假日期= ', qjDate
                   
                        qjDateFormat = qjDate.replace('-', ',')
                        qjDateFormat = qjDateFormat.split(',')
                        qjDateFormat[0] = int(qjDateFormat[0])
                        qjDateFormat[1] = int(qjDateFormat[1])
                        qjDateFormat[2] = int(qjDateFormat[2]) + k
                        #print 'qjDateFormat[2]= ',qjDateFormat[2]
                        if qjDateFormat[2] > last_day:
                            print '本月%s申请了跨月休假' %name_combin
                            break
                        theDay = datetime.datetime(qjDateFormat[0],qjDateFormat[1],qjDateFormat[2]).strftime('%w')
                        #print 'theDay= ',theDay
                        if theDay != 6: #不是周末
                            wsheet_dag.write(h,l + k,returnQjlx(j))
                            #qjDate
                        elif theDay == 6:   #是周六
                            l += 2
                            wsheet_dag.write(h,l + k,returnQjlx(j))
                            
                
                elif qjts_combin == 1:
                    #写入请假类型符号
                    wsheet_dag.write(h,l,returnQjlx(j))
                elif qjts_combin < 1:
                    wsheet_dag.write(h,l,'※')
            
            elif qjlx_combin == '':
                
                if qjTime != '':         
                #判断时间是否有2个值
                
                    if len(qjTime.split()) > 1:
                        #continue
                        wsheet_dag.write(h,l,'√')
                    else:
                        #暂定先按半日算，写入'半日'对应符号
                        wsheet_dag.write(h,l,'※')
                        
            else:
                continue
                
        else:
            continue
        
        l += 1

    h += 1
        
wbook_dag.save(location + today + 'dag' + '.xls')

#--------------------------------------统计技术处考勤---------------------------------
jsc_list = ['王羽佳','秦建华','张洪义','齐林','树国威','郑迪乔']
h = 6
for i in jsc_list:
    #print 'h= ',h
    l = 2
    for j in range(sheet_combin.nrows): #行数
        
        name_combin = sheet_combin.cell_value(j,1).encode('utf8')
        qjDate = sheet_combin.cell_value(j,3)   #日期
        qjts_combin = sheet_combin.cell_value(j,5).encode('utf8')   #请假天数     
        #qjts_combin = int(qjts_combin)     
        qjlx_combin = sheet_combin.cell_value(j,6).encode('utf8')   #请假类型
        
        qjTime = sheet_combin.cell_value(j,4)   #时间
       
        #判断名字是否相同
        if i == name_combin:
            #判断请假类型是否为空
            '''
            print '请假人：',name_combin
            print 'j= ', j
            print '请假类型= ',qjlx_combin
            '''
            if qjlx_combin != '':
                #判断请假天数是否>1
                if qjts_combin > 1:
                    #根据天数写入，隔开周末
                    qjts_combin = round(float(qjts_combin))
                    qjts_combin = int(qjts_combin)
                    #print '请假天数= ',qjts_combin
                    
                    for k in range(qjts_combin):
                        #判断请假时间是否是周末
                        #print 'k= ',k
                        #print '请假日期= ', qjDate
                   
                        qjDateFormat = qjDate.replace('-', ',')
                        qjDateFormat = qjDateFormat.split(',')
                        qjDateFormat[0] = int(qjDateFormat[0])
                        qjDateFormat[1] = int(qjDateFormat[1])
                        qjDateFormat[2] = int(qjDateFormat[2]) + k
                        #print 'qjDateFormat[2]= ',qjDateFormat[2]
                        if qjDateFormat[2] > last_day:
                            print '本月%s申请了跨月休假' %name_combin
                            break
                        theDay = datetime.datetime(qjDateFormat[0],qjDateFormat[1],qjDateFormat[2]).strftime('%w')
                        #print 'theDay= ',theDay
                        if theDay != 6: #不是周末
                            wsheet_jsc.write(h,l + k,returnQjlx(j))
                            #qjDate
                        elif theDay == 6:   #是周六
                            l += 2
                            wsheet_jsc.write(h,l + k,returnQjlx(j))
                            
                
                elif qjts_combin == 1:
                    #写入请假类型符号
                    wsheet_jsc.write(h,l,returnQjlx(j))
                elif qjts_combin < 1:
                    wsheet_jsc.write(h,l,'※')
            
            elif qjlx_combin == '':
                
                if qjTime != '':         
                #判断时间是否有2个值
                
                    if len(qjTime.split()) > 1:
                        #continue
                        wsheet_jsc.write(h,l,'√')
                    else:
                        #暂定先按半日算，写入'半日'对应符号
                        wsheet_jsc.write(h,l,'※')
                        
            else:
                continue
                
        else:
            continue
        
        l += 1

    h += 1
        
wbook_jsc.save(location + today + 'jsc' + '.xls')

print '数据汇总统计工作完成！'
