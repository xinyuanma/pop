#!/usr/bin/env python
# coding: utf-8

import datetime
import time
import numpy as np
import pandas as pd
import warnings
import os
import xlwings as xw
import sys
from win32com.client import Dispatch
warnings.filterwarnings("ignore")  # 取消警告







####获取data文件路径#####
get_path = os.getcwd() + r'\Data'
(pop_period, first_report_period, period_list) = checkperiod()
pop_version = checkversion(pop_period)
if pop_version is None:
    sys.exit(0)
if input("Please confirm you want to create %s %s POP File (Y/N)" %
         (pop_period, pop_version)) == "N":
    sys.exit(0)








data_path = os.getcwd() + "\Data"
PCMapping = pd.read_excel(data_path + "\Profit Center Hierarchy Flattened (MDG).xlsx")
popdata = data_path + "\POPData"
popyearlydata1 = data_path + "\Data1"
popyearlydata2 = data_path + "\Data2"





#计算POP
file_names = []
file_paths = []
for roots, dirs, files in os.walk(popdata):
    for file in files:
        if file[-5:] == '.xlsx':
            file_names.append(file[:file.rfind(' Data')])
            file_paths.append(popdata + "\\" + file)
for file_path, file_name in zip(file_paths, file_names):
    cal_pop(file_path, file_name)
endtime = datetime.datetime.now()
print('用时: %d s' % ((endtime - starttime).seconds))  # 程序用时





#计算POP Yearly
file_names = []
file_data1paths = []
file_data2paths = []
file_dataRYpaths =[]
for roots, dirs, files in os.walk(popyearlydata1):
    for file in files:
        if file[-5:] == '.xlsx':
            file_names.append(file[:file.rfind(' Data1')])
            file_data1paths.append(popyearlydata1 + "\\" + file)
for o in file_data1paths:
    file_data2paths.append(o.replace("Data1", "Data2"))
for o in file_data1paths:
    file_dataRYpaths.append(o.replace("Data1", "DataRY"))
for file_path1, file_path2, file_pathRY,file_name in zip(file_data1paths, file_data2paths,file_dataRYpaths,
                                             file_names):
    cal_yearlypop(file_path1, file_path2, file_pathRY,file_name)






lastfilepath = os.getcwd() + r"\LastFiles"
pathdata = os.getcwd() + r'\POP\%s.csv'
pathyearlydata = os.getcwd() + r'\POPYearly\%s-yearly.csv'
file_names = []
file_poppaths = []
file_poppath1s = []
file_poppath2s = []
for roots, dirs, files in os.walk(lastfilepath):
    for file in files:
        if file[:1] != '~' and file[-5:] == '.xlsx':
            file_poppaths.append(lastfilepath + "\\" + file)
for i in file_poppaths:
    file_name = i.split("\\")[-1][:3]
    file_name = ("JGP" if (i.split("\\")[-1][:7] == 'JABIL G') else "JRI" if
                 (i.split("\\")[-1][:7] == 'JABIL R') else "JPS" if
                 (i.split("\\")[-1][:7] == 'JABIL P') else file_name)
    file_names.append(file_name)
    file_poppath1s.append(pathdata % file_name)
    file_poppath2s.append(pathyearlydata % file_name)






#把结果复制到对应的POP文件

popfilesavepath = os.getcwd() + r"\POPFiles"
popperiodsort = []
for i in range(0, 4):
    if int(pop_period[1:2]) + i > 4:
        a = int(pop_period[1:2]) + i - 4
    else:
        a = int(pop_period[1:2]) + i
    popperiodsort.append("Q%s" % str(a))
current_version = "8 Quarter (%s %s)" % (pop_period, pop_version)
try:
    for path, path1, path2, file_name in zip(file_poppaths, file_poppath1s,
                                             file_poppath2s, file_names):

        app = xw.App(visible=True, add_book=False)  # 程序可见，只打开不新建工作薄
        app.display_alerts = False  # 警告关闭
        app.screen_updating = True  # 屏幕更新关闭
        wb = app.books.open(path)  # 打开该excel
        Sheets = wb.sheets
        wb1 = app.books.open(path1)  # 打开该excel
        wb2 = app.books.open(path2)

        sheet_POPdata = wb1.sheets[file_name]  # 指定sheet
        sheet_POPyearlydata = wb2.sheets['%s-yearly' % file_name]

        sheet_POP = wb.sheets['Data']  # 指定sheet
        if pop_version == 'Bid1':
            popdata = sheet_POPdata.range('A1').expand('table').value
            sheet_POP.clear()
            sheet_POP.range('A1').options(expand='table').value = popdata
        else:
            popdata = sheet_POPdata.range('A1').expand('table').value
            row = sheet_POP.range('A1').expand('table').last_cell.row + 1
            sheet_POP.range('A' +
                            str(row)).options(expand='table').value = popdata
            sheet_POP.range('A' + str(row)).api.EntireRow.Delete()
        sheet_YearlyPOP = wb.sheets["YearlyData"]  # 指定sheet
        popyearlydata = sheet_POPyearlydata.range('A1').expand('table').value
        sheet_YearlyPOP.clear()
        sheet_YearlyPOP.range('A1').options(
            expand='table').value = popyearlydata

        #sheet_POP.api.visible = False  #hide sheet
        #sheet_YearlyPOP.api.visible = False
        wb1.close()
        wb2.close()

        #time.sleep(15)
        wb1 = app.books.open(os.getcwd()+r"\refreshall.xlsm")
        refreshall = wb1.macro('refreshall')
        for Sheet in Sheets:
            wb.sheets[Sheet].activate()
            if (Sheet.name == 'PoP'):
                refreshall()
                time.sleep(15)
            #wb.api.ActiveSheet.PivotTables('PivotTable1').PivotCache().refreshall

            if (Sheet.name == 'PoP') or (Sheet.name == 'Capex'):
                for p, i in zip(popperiodsort, range(1, 5)):
                    wb.sheets[Sheet].activate()
                    wb.api.ActiveSheet.PivotTables("PivotTable1").PivotFields(
                        "Version").PivotItems("%s Locked Forecast" %
                                              p).Position = i

            if Sheet.name not in [
                    'Data', 'YearlyData', 'PoP', 'Yearly PoP', 'Capex',
                    'CHECK', 'Yearly Capex'
            ]:
                wb.sheets[Sheet].activate()
                wb.api.ActiveSheet.PivotTables('PivotTable1').PivotFields(
                    "Version").CurrentPage = current_version

            if Sheet.name not in [
                    'Data', 'YearlyData', 'Yearly PoP', 'CHECK', 'Yearly Capex'
            ]:
                wb.sheets[Sheet].activate()
                if (wb.sheets[Sheet].range('C11') == period_list[0]) | (
                        wb.sheets[Sheet].range('D11') == period_list[0]):
                    wb.api.ActiveSheet.PivotTables("PivotTable1").PivotFields(
                        "Period").PivotItems("%s" %
                                             period_list[0]).Visible = False

            if Sheet.name in ['Yearly PoP', 'Yearly Capex']:
                wb.sheets[Sheet].activate()
                wb.api.ActiveSheet.PivotTables("PivotTable1").PivotFields(
                    "Period").PivotItems("FY 2017").Visible = False
                wb.api.ActiveSheet.PivotTables("PivotTable1").PivotFields(
                    "Period").PivotItems("FY 2022").Visible = True
        file_name = changename(file_name)
        wb.sheets['PoP'].activate()
        #Del Account:  Other Profit/(Loss)
        #wb.api.ActiveSheet.PivotTables("PivotTable1").PivotFields(
        #                "Account").PivotItems(" Other Profit/(Loss)").Visible = False
        wb.save(popfilesavepath + '\%s POP File_%s.xlsx' %
                (file_name, current_version))
        wb.close()
        #app.quit()
        print("%s POP File Done!" % file_name)
except Exception as e:
    print(e)
finally:
    app.quit()






# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:






# In[ ]:





# In[ ]:





# In[ ]:




