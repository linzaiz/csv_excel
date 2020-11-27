# -*- coding:utf-8 -*-     （  ）
# ****************************************************************
# 程序性质：Excel xls/xlsx 处理
# 功能：从付款信息大表读入，按要求分别存到付款文件中
#
# 使用方法： 1. 在工作目录建立Templates子目录，先填好里面模板
#           2. 更新工作目录 working_path
#           3. python运行
# 输出： 全部输出到工作目录里
# 注意事项： 1. 设置 working_path， 如 working_path = r'\财务\付款申请——2020-11--III'
#           2. 设置Excel读取记录数 num_of_row
#           3. 以及Excel列标题从第几行开始等
# *****************************************************************
# auth__ = 'Larry Zhang' # 2020-11-24 18:08
# v1.00     2020-11-24      不同的乙方，这版全，但日期都留空了（因合同没签完）

# import sys
import os
import pandas as pd
# import re
# import datetime, time
# import numpy as np    # 20191102 np.NaN用这个
# import xlrd.xldate
# from openpyxl.workbook import Workbook
from openpyxl.reader.excel import load_workbook
# from openpyxl.writer.excel import ExcelWriter
from openpyxl.utils import get_column_letter, column_index_from_string  # 不是from .cells里，改为.utils了


def int_or_money(obj):
    if isinstance(obj, float):
        if obj % 1 == 0.0:
            return int(obj)
        else:
            return round(obj, 2)
    else:
        return obj


def chn_date(np_datetime):
    if not pd.isna(np_datetime):
        return pd.to_datetime(np_datetime).strftime( '%Y年%m月%d日'.encode('unicode_escape').decode('utf8')).encode('utf-8').decode('unicode_escape')
    else:
        return ''


def a_contract_prt(outtable, i):
    wb = load_workbook('templates\\' + outtable)  # encoding='gb18030'不像csv模块，没有这个encoding。
    wb.guess_types = True  # 加个这个，写入百分数后就会显示百分数，否则显示小数
    # print( 'wb.get_named_ranges()=', wb.get_named_ranges() )
    sheetNames = wb.sheetnames
    print( 'wb.sheet_names[0]= 【', sheetNames[0], f'】  第【{i}】个。' )
    sh = wb[(sheetNames[0])]   # sh = wb.get_sheet_by_name('Sheet1')
    # sh = wb.active #也可。没有()
    print( sh.title, sh.max_row, sh.max_column )
    print( get_column_letter(sh.max_column) )  # 从1开始

    if '付款申请单' in outtable:
        sh["B6"].value = dfBig['收款方名称'][i]
        sh["B7"].value = dfBig['合同名称'][i].replace('\n', '')
        # sh["B8"].value = chn_date(dfBig['合同签订时间'][i])  # pd.to_datetime(dfBig['合同签订时间'][i]).strftime( r'%Y年%m月%d日' )
        #                                  # str.split(dfBig['合同签订时间'][i].astype(str), 'T')[0]  # dfBig['合同签订时间'][i].astype(datetime.date)
        sh["G8"].value = dfBig['合同金额\n（元）'][i]
        sh["C11"].value = dfBig['应付金额\n（元）'][i]
        sh["E12"].value = dfBig['开户银行'][i]
        sh["H12"].value = dfBig['银行账号'][i]
        if not pd.isna(dfBig['成果文件说明'][i]):
            resultss = str(dfBig['成果文件说明'][i])
        else:
            resultss = ''
        sh["B13"].value = "    根据与贵单位签订的《" + dfBig['合同名称'][i].replace('\n', '') + '》，“' \
                          + dfBig['付款条款'][i].replace('\n', '') + "”\n    " + resultss \
                          + "\n    按照约定申请支付" + str(int_or_money(dfBig['应付金额\n（元）'][i])) + "元。"
    elif '付款台账' in outtable:
        sh["B2"].value = dfBig['收款方名称'][i]
        sh["B4"].value = dfBig['付款单位'][i]
        sh["C4"].value = dfBig['合同金额\n（元）'][i]
        sh["D4"].value = dfBig['项目'][i]
        sh["E4"].value = dfBig['未付金额\n（元）'][i]
        sh["F4"].value = dfBig['已付金额'][i]
        sh["G4"].value = dfBig['应付金额\n（元）'][i]
        sh["H4"].value = chn_date(dfBig['付款时间'][i])
        sh["I4"].value = dfBig['备注'][i]
    elif '付款审批单' in outtable:
        sh["B6"].value = dfBig['收款方名称'][i]
        sh["B10"].value = dfBig['项目'][i].replace('\n', '')
        sh["B14"].value = dfBig['合同名称'][i].replace('\n', '')
        # sh["B13"].value = chn_date(dfBig['合同签订时间'][i])  # str.split(dfBig['合同签订时间'][i].astype(str), 'T')[0]
        sh["G14"].value = dfBig['合同金额\n（元）'][i]
        sh["C17"].value = dfBig['应付金额\n（元）'][i]
        sh["E18"].value = dfBig['开户银行'][i]
        sh["H18"].value = dfBig['银行账号'][i]
        sh["B19"].value = "    根据我单位与天津智土科技有限公司签订的《" + dfBig['合同名称'][i].replace('\n', '') + "》的约定，申请支付" \
                          + str(int_or_money(dfBig['应付金额\n（元）'][i])) + "元。\n    妥否，请领导批示。"

    newf1 = os.path.splitext(outtable)[0] + dfBig['文件名缩写'][i] + '.xlsx'
    newf1 = str(i) + newf1[1:]
    wb.save(filename=os.path.join(working_path, newf1))
    wb.close()


if __name__ == '__main__':

    working_path = r'\财务\付款申请——2020-11--III'
    templ_path = os.path.join(working_path, "Templates")
    iptFP = os.path.join(templ_path, r'zzz、付款信息汇总大表--III.xlsx')
    num_of_row = 3

    outt1 = 'x1、付款申请单 - .xlsx'
    outt5 = 'x5、付款台账 - .xlsx'
    outt6 = 'x6、付款审批单 - .xlsx'

    outt_hz = 'z7、票据汇总单 -  1踏勘+2林地（鑫茂兴）.xlsx'

    os.chdir( working_path )

    if not os.access( iptFP, os.F_OK ):
        print( "访问不了源表(付款信息大表)【%s】 ! please check..." % iptFP )
        os.system('pause')
        os._exit(-1)    # 这个不报错退出。 exit(-1)会报错。sys.exit(-1)

    # ##############读取大表#########################
    print( 'Loading Mapping File【', iptFP, '】，Please wait ...... ' )
    mapX = pd.ExcelFile( iptFP )
    dfBig = pd.read_excel( mapX, sheet_name=u'大表', header=3 - 1, index_col=u'序号' )  # header默认=0，是第0+1行，=2是指2+1行。
    print( '大表: \n', dfBig, end='\n\n' )                  # , usecols="A: T"

    wbhz = load_workbook('templates\\' + outt_hz)  # encoding='gb18030'不像csv模块，没有这个encoding。
    wbhz.guess_types = True  # 加个这个，写入百分数后就会显示百分数，否则显示小数
    for i in range(1, num_of_row + 1):
        # print( 'wbhz.get_named_ranges()=', wbhz.get_named_ranges() )
        sheetNames = wbhz.sheetnames
        print( 'wbhz.get_sheet_names()=', sheetNames )
        sh = wbhz[(sheetNames[0])]   # sh = wbhz.get_sheet_by_name('Sheet1')
        # sh = wbhz.active #也可。没有()
        print( sh.title, sh.max_row, sh.max_column )
        print( get_column_letter(sh.max_column) )  # 从1开始
        sh[f"A{7 + i}"].value = dfBig['项目'][i].replace('\n', '')
        sh[f"F{7 + i}"].value = dfBig['应付金额\n（元）'][i]
    wbhz.save(filename=os.path.join(working_path, outt_hz))
    wbhz.close()

    for i in range(1, num_of_row + 1):   # 直接改数吧 range1-13 =12个。 len(dfBig.index)
        a_contract_prt(outt1, i)
        a_contract_prt(outt5, i)
        a_contract_prt(outt6, i)
