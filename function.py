######import_data方法：读取excel数据######
def import_data(path):
    #####读取excel#####
    split_data = pd.read_excel(path, header=None)
    #####读取已有列名######
    split_data.columns = split_data.iloc[6, :].values.tolist()
    #####添加列名######
    if (path[-6:-5] == '1')|(path[-6:-5] == '2'):
        addColumns = ['CUSTOMER', 'Cost Center', 'Version', 'Period']
    else:
        addColumns = ['CUSTOMER', 'Cost Center', 'Period', 'Version']
    #####遍历列名######
    newColumns = [c for c in split_data.columns if c not in addColumns]
    for i in range(0, 4):
        newColumns[i] = addColumns[i]
    #####重设列名######
    split_data.columns = newColumns
    ######判断是否需要删除无效行##########
    if_clear_excel_head = split_data['CUSTOMER'][0]
    ######删除无效行##########
    if if_clear_excel_head == 'CUBE:':
        split_data = split_data.drop(split_data.index[:7])
    ######重新索引#####
    split_data = split_data.reset_index(drop=True)
    return split_data

######createMapping方法：读取PC Mapping文件######
def createMapping(df):
    ######删除前四列['PC','CUS_SUBGROUP3','CUS_SUBGROUP2','CUS_SUBGROUP1']#######
    df = df.drop((PCMapping.columns[0:4]), axis=1)
    ###### 删除最后一列ALL_CUSTOMERS #########
    df = df.drop("ALL_CUSTOMERS", 1)
    ###### 选择部分列 #########
    df = df[[
        'SEC_SUBGROUP5', 'SEC_SUBGROUP4', 'SEC_SUBGROUP3', 'SEC_SUBGROUP2',
        'SEC_SUBGROUP1', 'SECTOR', 'DIV_SUBGROUP1', 'DIVISION',
        'SEG_SUBGROUP1', 'SEGMENT', 'CUSTOMER'
    ]]
    ###### 改变列顺序 ######
    df.columns = [
        'SECTOR SUBGROUP5', 'SECTOR SUBGROUP4', 'SECTOR SUBGROUP3',
        'SECTOR SUBGROUP2', 'SECTOR SUBGROUP1', 'SECTOR', 'DIVISION SUBGROUP1',
        'DIVISION', 'SEGMENT SUBGROUP1', 'SEGMENT', 'CUSTOMER'
    ]
    return df


###### merge方法：合并 ######
def merge(left, right):
    result = pd.merge(left, right, on="CUSTOMER", how="right")
    result_column = list(result.columns)
    result = result.drop_duplicates(subset=result_column, keep='first')  # 去重
    return result


def checkversion(pop_period):
    pop_folder_names = []
    for roots, dirs, files in os.walk(
            r"C:\Users\1243712\OneDrive - Jabil\Desktop\Work\POP 2022"):
        for dir in dirs:
            if dir.rfind('%s' % pop_period) != -1:
                pop_folder_names.append(dir[14:18])
    if 'Lock' in pop_folder_names:
        return print(
            "%s Locked POP file folder has been created! Please check!" %
            pop_period)
    if 'Bid2' in pop_folder_names:
        return "Locked"
    if 'Bid1' in pop_folder_names:
        return "Bid2"
    if 'Bid1' not in pop_folder_names:
        if 'Bid2' not in pop_folder_names:
            if 'Lock' not in pop_folder_names:
                return "Bid1"

###### checkperiod方法：检查当前period ######
def checkperiod():
    ###### 读取POPData文件夹里面的xlsx的名字（前三个字符）######
    popfilespaths = []
    for roots, dirs, files in os.walk(os.getcwd() + r"\Data\POPData"):
        for file in files:
            if file[:1] != '~' and file[-5:] == '.xlsx':
                if ((file[:3] == 'JGP') | (file[:3] == 'ALL')):
                    name = file[:3]
    ###### 读取上面所选的文件 ######
    read_data = import_data(get_path + '\\POPData\\%s Data.xlsx' % name)
    read_periods = list(read_data['Period'].drop_duplicates(keep='first'))
    ####### period排序 ######
    df = pd.DataFrame(read_periods)
    df.columns = ["period"]
    df["sort"] = [int(o[1]) + int(o[6]) * 4 for o in df.period.values]
    df.sort_values(by=['sort'], inplace=True)
    df.reset_index(drop=True, inplace=True)
    if len(df) != 9:
        print(df["period"])
    pop_period = str(df['period'][5][:2]) + str(df['period'][5][-2:])
    first_report_period = str(df['period'][1][:2]) + str(df['period'][1][-2:])
    period_list = df["period"].tolist()
    #######period_list = period_list + ["Q0 0000"] * (9- len(period_list))#quarters不够的话自动补全 ######
    return (pop_period, first_report_period, period_list)


def changename(oldname):
    if oldname == "JGP":
        newname = 'JABIL GREEN POINT SEGMENT'
    if oldname == "ALL":
        newname = 'ALL CUSTOMERS'
    return newname



def cal_AVGROIC(path):
    #################计算AVG ROIC##########################
    #################导入25个月的ROIC数据##########################
    split_data = pd.read_excel(path, header=None)
    periodColumns = split_data.iloc[8, :].values.tolist()
    versionColumns = split_data.iloc[7, :].values.tolist()
    split_data = split_data.drop(split_data.index[:9])
    periodVersion = []
    for (p, v) in zip(periodColumns, versionColumns):
        periodVersion.append(str(v)[:9] + ' ' + str(p))
    periodVersion[0] = "CUSTOMER"
    periodVersion[1] = "Cost Center"
    split_data.columns = periodVersion
    split_data = split_data.reset_index(drop=True)
    ###############设置公式###############################


    ####Q4 Locked###
    split_data['Q4 Locked Forecast Q4 2021'] = (
        split_data['8 Quarter May-21'] + split_data['Q4 Locked Jun-21'] +
        split_data['Q4 Locked Jul-21'] + split_data['Q4 Locked Aug-21']) / 4
    split_data['Q4 Locked Forecast Q1 2022'] = (
        split_data['Q4 Locked Aug-21'] + split_data['Q4 Locked Sep-21'] +
        split_data['Q4 Locked Oct-21'] + split_data['Q4 Locked Nov-21']) / 4
    split_data['Q4 Locked Forecast Q2 2022'] = (
        split_data['Q4 Locked Nov-21'] + split_data['Q4 Locked Dec-21'] +
        split_data['Q4 Locked Jan-22'] + split_data['Q4 Locked Feb-22']) / 4
    split_data['Q4 Locked Forecast Q3 2022'] = (
        split_data['Q4 Locked Feb-22'] + split_data['Q4 Locked Mar-22'] +
        split_data['Q4 Locked Apr-22'] + split_data['Q4 Locked May-22']) / 4

    ####Q1 Locked###
    split_data['Q1 Locked Forecast Q1 2022'] = (
        split_data['8 Quarter Aug-21'] + split_data['Q1 Locked Sep-21'] +
        split_data['Q1 Locked Oct-21'] + split_data['Q1 Locked Nov-21']) / 4
    split_data['Q1 Locked Forecast Q2 2022'] = (
        split_data['Q1 Locked Nov-21'] + split_data['Q1 Locked Dec-21'] +
        split_data['Q1 Locked Jan-22'] + split_data['Q1 Locked Feb-22']) / 4
    split_data['Q1 Locked Forecast Q3 2022'] = (
        split_data['Q1 Locked Feb-22'] + split_data['Q1 Locked Mar-22'] +
        split_data['Q1 Locked Apr-22'] + split_data['Q1 Locked May-22']) / 4
    split_data['Q1 Locked Forecast Q4 2022'] = (
        split_data['Q1 Locked May-22'] + split_data['Q1 Locked Jun-22'] +
        split_data['Q1 Locked Jul-22'] + split_data['Q1 Locked Aug-22']) / 4

    ####Q2 Locked###
    split_data['Q2 Locked Forecast Q2 2022'] = (
        split_data['8 Quarter Nov-21'] + split_data['Q2 Locked Dec-21'] +
        split_data['Q2 Locked Jan-22'] + split_data['Q2 Locked Feb-22']) / 4
    split_data['Q2 Locked Forecast Q3 2022'] = (
        split_data['Q2 Locked Feb-22'] + split_data['Q2 Locked Mar-22'] +
        split_data['Q2 Locked Apr-22'] + split_data['Q2 Locked May-22']) / 4
    split_data['Q2 Locked Forecast Q4 2022'] = (
        split_data['Q2 Locked May-22'] + split_data['Q2 Locked Jun-22'] +
        split_data['Q2 Locked Jul-22'] + split_data['Q2 Locked Aug-22']) / 4
    split_data['Q2 Locked Forecast Q1 2023'] = (
        split_data['Q2 Locked Aug-22'] + split_data['Q2 Locked Sep-22'] +
        split_data['Q2 Locked Oct-22'] + split_data['Q2 Locked Nov-22']) / 4

    ####Q3 Locked###
    split_data['Q3 Locked Forecast Q3 2022'] = (
        split_data['8 Quarter Feb-22'] + split_data['Q3 Locked Mar-22'] +
        split_data['Q3 Locked Apr-22'] + split_data['Q3 Locked May-22']) / 4
    split_data['Q3 Locked Forecast Q4 2022'] = (
        split_data['Q3 Locked May-22'] + split_data['Q3 Locked Jun-22'] +
        split_data['Q3 Locked Jul-22'] + split_data['Q3 Locked Aug-22']) / 4
    split_data['Q3 Locked Forecast Q1 2023'] = (
        split_data['Q3 Locked Aug-22'] + split_data['Q3 Locked Sep-22'] +
        split_data['Q3 Locked Oct-22'] + split_data['Q3 Locked Nov-22']) / 4
    split_data['Q3 Locked Forecast Q2 2023'] = (
        split_data['Q3 Locked Nov-22'] + split_data['Q3 Locked Dec-22'] +
        split_data['Q3 Locked Jan-23'] + split_data['Q3 Locked Feb-23']) / 4

    ####8 Quarter##

    split_data['8 Quarter Q4 2021'] = (
        split_data['8 Quarter May-21'] + split_data['8 Quarter Jun-21'] +
        split_data['8 Quarter Jul-21'] + split_data['8 Quarter Aug-21']) / 4
    split_data['8 Quarter Q1 2022'] = (
        split_data['8 Quarter Aug-21'] + split_data['8 Quarter Sep-21'] +
        split_data['8 Quarter Oct-21'] + split_data['8 Quarter Nov-21']) / 4
    split_data['8 Quarter Q2 2022'] = (
        split_data['8 Quarter Nov-21'] + split_data['8 Quarter Dec-21'] +
        split_data['8 Quarter Jan-22'] + split_data['8 Quarter Feb-22']) / 4
    split_data['8 Quarter Q3 2022'] = (
        split_data['8 Quarter Feb-22'] + split_data['8 Quarter Mar-22'] +
        split_data['8 Quarter Apr-22'] + split_data['8 Quarter May-22']) / 4
    split_data['8 Quarter Q4 2022'] = (
        split_data['8 Quarter May-22'] + split_data['8 Quarter Jun-22'] +
        split_data['8 Quarter Jul-22'] + split_data['8 Quarter Aug-22']) / 4
    split_data['8 Quarter Q1 2023'] = (
        split_data['8 Quarter Aug-22'] + split_data['8 Quarter Sep-22'] +
        split_data['8 Quarter Oct-22'] + split_data['8 Quarter Nov-22']) / 4
    split_data['8 Quarter Q2 2023'] = (
        split_data['8 Quarter Nov-22'] + split_data['8 Quarter Dec-22'] +
        split_data['8 Quarter Jan-23'] + split_data['8 Quarter Feb-23']) / 4
    split_data['8 Quarter Q3 2023'] = (
        split_data['8 Quarter Feb-23'] + split_data['8 Quarter Mar-23'] +
        split_data['8 Quarter Apr-23'] + split_data['8 Quarter May-23']) / 4
    df = split_data[[
        'CUSTOMER', 'Cost Center',
        'Q4 Locked Forecast Q4 2021','Q4 Locked Forecast Q1 2022', 'Q4 Locked Forecast Q2 2022','Q4 Locked Forecast Q3 2022',
        'Q1 Locked Forecast Q1 2022','Q1 Locked Forecast Q2 2022', 'Q1 Locked Forecast Q3 2022','Q1 Locked Forecast Q4 2022',
        'Q2 Locked Forecast Q2 2022','Q2 Locked Forecast Q3 2022', 'Q2 Locked Forecast Q4 2022','Q2 Locked Forecast Q1 2023',
        'Q3 Locked Forecast Q3 2022','Q3 Locked Forecast Q4 2022', 'Q3 Locked Forecast Q1 2023','Q3 Locked Forecast Q2 2023',
        '8 Quarter Q4 2021', '8 Quarter Q1 2022', '8 Quarter Q2 2022','8 Quarter Q3 2022',
        '8 Quarter Q4 2022', '8 Quarter Q1 2023', '8 Quarter Q2 2023','8 Quarter Q3 2023'
    ]]
    ################## unpivot##################################
    df = pd.melt(df,
                 id_vars=["CUSTOMER", "Cost Center"],
                 value_vars=[
                     'CUSTOMER', 'Cost Center',
        'Q4 Locked Forecast Q4 2021','Q4 Locked Forecast Q1 2022', 'Q4 Locked Forecast Q2 2022','Q4 Locked Forecast Q3 2022',
        'Q1 Locked Forecast Q1 2022','Q1 Locked Forecast Q2 2022', 'Q1 Locked Forecast Q3 2022','Q1 Locked Forecast Q4 2022',
        'Q2 Locked Forecast Q2 2022','Q2 Locked Forecast Q3 2022', 'Q2 Locked Forecast Q4 2022','Q2 Locked Forecast Q1 2023',
        'Q3 Locked Forecast Q3 2022','Q3 Locked Forecast Q4 2022', 'Q3 Locked Forecast Q1 2023','Q3 Locked Forecast Q2 2023',
        '8 Quarter Q4 2021', '8 Quarter Q1 2022', '8 Quarter Q2 2022','8 Quarter Q3 2022',
        '8 Quarter Q4 2022', '8 Quarter Q1 2023', '8 Quarter Q2 2023','8 Quarter Q3 2023'
                 ],
                 var_name='versionPeriod')
    df['Period'] = [o[-7:] for o in df.versionPeriod.values]
    df['Version'] = [(o.replace(" "+p, ""))
                     for (o, p) in zip(df.versionPeriod.values, df.Period.values)]
    df = df.rename(columns={'value':'Avg ROIC Total Net Assets (Less Customer Gear)'})
    df1 = df[['CUSTOMER', 'Cost Center', 'Period', 'Version', 'Avg ROIC Total Net Assets (Less Customer Gear)']]
    return df1


def add_AVGROIC(df, df1):
    #df = df.drop('Avg ROIC Total Net Assets (Less Customer Gear)', 1)
    result = df.append(df1, ignore_index=True, sort=False)
    #add_columns = ['Avg ROIC Total Net Assets (Less Customer Gear)']
    #new_columns = [c for c in df.columns if c not in add_columns] + add_columns
    #result = result[new_columns]
    result.to_csv('addc.csv')
    return result



def cal_YEARLYAVGROIC(path):
    split_data = pd.read_excel(path ,header = None)
    periodColumns = split_data.iloc[8, :].values.tolist()
    split_data = split_data.drop(split_data.index[:9])
    period_Months = []
    for (p) in periodColumns:
        period_Months.append(p)
    period_Months[0] = "CUSTOMER"
    period_Months[1] = "Cost Center"
    split_data.columns = period_Months
    split_data = split_data.reset_index(drop=True)
    period_FY18 = ['CUSTOMER','Cost Center']+periodColumns[2:15]
    period_FY19 = ['CUSTOMER','Cost Center']+periodColumns[14:27]
    period_FY20 = ['CUSTOMER','Cost Center']+periodColumns[26:39]
    period_FY21 = ['CUSTOMER','Cost Center']+periodColumns[38:51]
    period_FY22 = ['CUSTOMER','Cost Center']+periodColumns[50:63]
    split_data_AVG = pd.DataFrame()
    #FY18
    split_data_AVG_FY18 = split_data[period_FY18]
    split_data_AVG_FY18['FY 2018'] = (split_data_AVG_FY18[period_FY18[2:]].T.sum())/13
    split_data_AVG = split_data_AVG.append(split_data_AVG_FY18[['CUSTOMER','Cost Center','FY 2018']], ignore_index=True, sort=False)
    #FY19
    split_data_AVG_FY19 = split_data[period_FY19]
    split_data_AVG_FY19['FY 2019'] = (split_data_AVG_FY19[period_FY19[2:]].T.sum())/13
    split_data_AVG = split_data_AVG.append(split_data_AVG_FY19[['CUSTOMER','Cost Center','FY 2019']], ignore_index=True, sort=False)
    #FY20
    split_data_AVG_FY20 = split_data[period_FY20]
    split_data_AVG_FY20['FY 2020'] = (split_data_AVG_FY20[period_FY20[2:]].T.sum())/13
    split_data_AVG = split_data_AVG.append(split_data_AVG_FY20[['CUSTOMER','Cost Center','FY 2020']], ignore_index=True, sort=False)
    #FY21
    split_data_AVG_FY21 = split_data[period_FY21]
    split_data_AVG_FY21['FY 2021'] = (split_data_AVG_FY21[period_FY21[2:]].T.sum())/13
    split_data_AVG = split_data_AVG.append(split_data_AVG_FY21[['CUSTOMER','Cost Center','FY 2021']], ignore_index=True, sort=False)
    #FY22
    split_data_AVG_FY22 = split_data[period_FY22]
    split_data_AVG_FY22['FY 2022'] = (split_data_AVG_FY22[period_FY22[2:]].T.sum())/13
    split_data_AVG = split_data_AVG.append(split_data_AVG_FY22[['CUSTOMER','Cost Center','FY 2022']], ignore_index=True, sort=False)
    #result
    result = pd.melt(split_data_AVG,id_vars = ["CUSTOMER", "Cost Center"],value_vars=['CUSTOMER','Cost Center','FY 2018','FY 2019','FY 2020','FY 2021','FY 2022'],var_name='Period')
    result['Version'] = ['8 Quarter' for o in result.Period.values]

    result = result.rename(columns={'value':'Avg ROIC Total Net Assets (Less Customer Gear)'})
    result = result[['CUSTOMER', 'Cost Center', 'Version', 'Period', 'Avg ROIC Total Net Assets (Less Customer Gear)']]
    return result

def add_YEARLYAVGROIC(df, df1):
    result = df.append(df1, ignore_index=True, sort=False)
    result.to_csv('yearlyroicadd.csv')
    return result



class POPData:
    def CCIA(df):
        df0 = df[[
            'CUSTOMER', 'Cost Center', 'Version', 'Period',
            'Net Working Capital', 'Revenue', 'Capex AP Adjustment'
        ]]
        df0 = POPData.BuQuan(df0, period_list)
        df0['Year'] = [o[3:] for o in df0.Period.values]
        df0['Quarter'] = [o[:3] for o in df0.Period.values]
        df0['V'] = [o[:2] if o[0] == 'Q' else '8Q' for o in df0.Version.values]
        df0['MinY'] = [
            str((int(o[2:]) - 1)) if
            ((o[:2] == 'Q1') & ((o[5:] == '21') | (o[5:] == '22') |
                                (o[5:] == '23'))) else o[3:]
            for o in df0.Period.values
        ]
        df0['MinQ'] = ['Q%s' % (int(o[1:]) - 1) for o in df0.Quarter.values]
        df0.replace({'MinQ': 'Q0'}, 'Q4', inplace=True)
        df0['X'] = df0['Quarter'] + df0['V']
        df0['MinV'] = ['8Q' if o[1] == o[4] else o[3:] for o in df0.X.values]
        df0.drop('X', 1, inplace=True)
        df0['label1'] = df0['V'] + df0['Year'] + [
            o[:2] for o in df0.Quarter.values
        ] + df0['Cost Center'] + df0['CUSTOMER']
        df0['label2'] = df0['MinV'] + df0['MinY'] + df0['MinQ'] + df0[
            'Cost Center'] + df0['CUSTOMER']
        df1 = df0[['Net Working Capital', 'Revenue', 'label1']]
        df3 = pd.merge(df0,
                       df1,
                       left_on="label2",
                       right_on='label1',
                       how='left')
        df3.fillna(0.00, inplace=True)
        df3['Net Working Capital_x'] = pd.to_numeric(
            df3['Net Working Capital_x'], errors='coerce')
        df3['Net Working Capital_y'] = pd.to_numeric(
            df3['Net Working Capital_y'], errors='coerce')
        df3['Revenue_x'] = pd.to_numeric(df3['Revenue_x'], errors='coerce')
        df3['Revenue_y'] = pd.to_numeric(df3['Revenue_y'], errors='coerce')
        df3['Capex AP Adjustment'] = pd.to_numeric(df3['Capex AP Adjustment'],
                                                   errors='coerce')
        df3.fillna(0.00, inplace=True)
        df3['Change in Working Capital'] = df3['Net Working Capital_x'] - df3[
            'Net Working Capital_y'] - df3['Capex AP Adjustment']
        df3['Change in Revenue'] = df3['Revenue_x'] - df3['Revenue_y']
        for n in [
                'Net Working Capital_x', 'Revenue_x', 'Capex AP Adjustment',
                'Year', 'Quarter', 'V', 'MinY', 'MinQ', 'MinV', 'label1_x',
                'label2', 'Net Working Capital_y', 'Revenue_y', 'label1_y'
        ]:
            result = df3[(df3['Change in Working Capital'] != 0) |
                         (df3['Change in Revenue'] != 0)]
        result.reset_index(drop=True)
        #result.to_csv('result.csv')
        return result

    def add_accounts(df, df1):
        result = df.append(df1, ignore_index=True, sort=False)
        add_columns = ['Change in Working Capital', 'Change in Revenue']
        new_columns = [c for c in df.columns if c not in add_columns
                       ] + add_columns
        result = result[new_columns]
        return result

    def ShaiXuan(df, period_list, pop_period):
        c = []
        a = int(pop_period[1])
        for i in range(0, 4):
            b = 4 if (a + i) % 4 == 0 else (a + i) % 4
            c.append(b)

        df = df[((df['Version'] == '8 Quarter') &
                 ((df['Period'] == period_list[1])
                  | (df['Period'] == period_list[2])
                  | (df['Period'] == period_list[3])
                  | (df['Period'] == period_list[4])
                  | (df['Period'] == period_list[5])
                  | (df['Period'] == period_list[6])
                  | (df['Period'] == period_list[7])
                  | (df['Period'] == period_list[8]))) |
                ((df['Version'] == 'Q%d Locked Forecast' % (c[0])) &
                 ((df['Period'] == period_list[1])
                  | (df['Period'] == period_list[2])
                  | (df['Period'] == period_list[3])
                  | (df['Period'] == period_list[4]))) |
                ((df['Version'] == 'Q%d Locked Forecast' % (c[1])) &
                 ((df['Period'] == period_list[2])
                  | (df['Period'] == period_list[3])
                  | (df['Period'] == period_list[4])
                  | (df['Period'] == period_list[5]))) |
                ((df['Version'] == 'Q%d Locked Forecast' % (c[2])) &
                 ((df['Period'] == period_list[3])
                  | (df['Period'] == period_list[4])
                  | (df['Period'] == period_list[5])
                  | (df['Period'] == period_list[6]))) |
                ((df['Version'] == 'Q%d Locked Forecast' % (c[3])) &
                 ((df['Period'] == period_list[4])
                  | (df['Period'] == period_list[5])
                  | (df['Period'] == period_list[6])
                  | (df['Period'] == period_list[7])))]
        df = df.reset_index(drop=True)  # 重新索引
        return df

    def BuQuan(df, period_list):
        df['CCPC'] = df['Cost Center'] + df['CUSTOMER']
        df1 = pd.DataFrame()
        df1['CCPC'] = df['CCPC']
        df1['Cost Center'] = df['Cost Center']
        df1['CUSTOMER'] = df['CUSTOMER']
        versions = ['8Q', 'Q1', 'Q2', 'Q3', 'Q4']
        periods = []
        for p in period_list:
            periods.append(p[3:] + p[:2])
        df1 = df1.drop_duplicates(subset=list(df1.columns), keep='first')
        df1 = df1.reset_index(drop=True)
        list1 = []
        list2 = []
        list3 = []
        for p in periods:
            for v in versions:
                list1.extend([v + p + o for o in df1.CCPC.values])
                list2.extend([
                    '8 Quarter' if v == '8Q' else
                    'Q1 Locked Forecast' if v == 'Q1' else 'Q2 Locked Forecast'
                    if v == 'Q2' else 'Q3 Locked Forecast' if v ==
                    'Q3' else 'Q4 Locked Forecast' if v == 'Q4' else '0'
                    for o in df1.CCPC.values
                ])
                list3.extend([p[4:] + " " + p[:4] for o in df1.CCPC.values])

        df2 = pd.DataFrame(data=list1, columns=['list1'])
        df2['list2'] = pd.DataFrame(list2)
        df2['list3'] = pd.DataFrame(list3)
        df2['KE'] = [o[8:] for o in df2.list1.values]
        df4 = pd.merge(df2, df1, left_on="KE", right_on='CCPC', how='left')
        df['V'] = [o[:2] if o[0] == 'Q' else '8Q' for o in df.Version.values]
        df['Year'] = [o[3:] for o in df.Period.values]
        df['Quarter'] = [o[:2] for o in df.Period.values]
        df['K'] = df['V'] + df['Year'] + df['Quarter'] + df[
            'Cost Center'] + df['CUSTOMER']
        df = df.drop("CCPC", 1)
        df = df.drop("V", 1)
        df = df.drop("Year", 1)
        df = df.drop("Quarter", 1)
        df3 = pd.merge(df4, df, left_on="list1", right_on='K', how='left')

        result = df3[[
            'CUSTOMER_x', 'Cost Center_x', 'list2', 'list3',
            'Net Working Capital', 'Revenue', 'Capex AP Adjustment'
        ]]
        result.fillna(0.00, inplace=True)
        result.columns = [
            'CUSTOMER', 'Cost Center', 'Version', 'Period',
            'Net Working Capital', 'Revenue', 'Capex AP Adjustment'
        ]
        return result

    def clear_lockedversion(df):
        df = df[(df['Version'] == '8 Quarter')]
        return df


class POPYearly:
    def CData1(df):
        df = df.reset_index(drop=True)  # 重新索引

        return df

    def CData2(df):
        df = df.reset_index(drop=True)  # 重新索引
        df['Version'] = [
            'Actual' if o[3:] != '2021' else '8 Quarter'
            for o in df.Period.values
        ]
        df['Period'] = ['FY' + o[2:] for o in df.Period.values]
        df.rename(columns={
            'Revenue': 'Last Q Revenue',
            'Bill of Materials': 'Last Q Bill of Materials',
            'Scrap Freight-In and Duty': 'Last Q Scrap Freight-In and Duty',
            'Cost of Materials': 'Last Q Cost of Materials',
            'Manufacturing Cost': 'Last Q Manufacturing Cost',
            'Intercompany Revenue': 'Last Q Intercompany Revenue'
        },
                  inplace=True)

        return df

    def BuQuan(df):
        df['CCPC'] = df['Cost Center'] + df['CUSTOMER']
        df1 = pd.DataFrame()
        df1['CCPC'] = df['CCPC']
        df1['Cost Center'] = df['Cost Center']
        df1['CUSTOMER'] = df['CUSTOMER']
        periods = []
        #Add FY22 remember to change to range(0,6) after FY21 final
        for i in range(-1, 6):
            pop_year = datetime.date.today().year - i
            periods.append("FY " + str(pop_year))

        df1 = df1.drop_duplicates(subset=list(df1.columns), keep='first')
        df1 = df1.reset_index(drop=True)
        list1 = []

        for p in periods:
            list1.extend(
                [str('8 Quarter') + str(p) + str(o) for o in df1.CCPC.values])

        df2 = pd.DataFrame(data=list1, columns=['list1'])
        df2['KE'] = [o[16:] for o in df2.list1.values]
        df4 = pd.merge(df2, df1, left_on="KE", right_on='CCPC', how='left')
        df['K'] = df['Version'] + df['Period'] + df['Cost Center'] + df[
            'CUSTOMER']
        df = df.drop("CCPC", 1)
        df3 = pd.merge(df4, df, left_on="list1", right_on='K', how='left')
        df3['Version'] = [o[:9] for o in list1]
        df3['Period'] = [o[9:16] for o in list1]
        result = df3[[
            'CUSTOMER_x', 'Cost Center_x', 'Version', 'Period',
            'Net Working Capital', 'Revenue', 'Capex AP Adjustment'
        ]]
        result.fillna('0.00', inplace=True)
        result.columns = [
            'CUSTOMER', 'Cost Center', 'Version', 'Period',
            'Net Working Capital', 'Revenue', 'Capex AP Adjustment'
        ]
        return result

    def CCIA(df):
        df0 = df[[
            'CUSTOMER', 'Cost Center', 'Version', 'Period',
            'Net Working Capital', 'Revenue', 'Capex AP Adjustment'
        ]]
        df0 = POPYearly.BuQuan(df0)
        df0['MinY'] = [
            'FY ' + str((int(o[3:]) - 1)) for o in df0.Period.values
        ]
        df0['label1'] = df0['Period'] + df0['Cost Center'] + df0['CUSTOMER']
        df0['label2'] = df0['MinY'] + df0['Cost Center'] + df0['CUSTOMER']
        df1 = df0[['Net Working Capital', 'Revenue', 'label1']]
        df3 = pd.merge(df0,
                       df1,
                       left_on="label2",
                       right_on='label1',
                       how='left')

        df3.fillna('0.00', inplace=True)

        df3['Net Working Capital_x'] = pd.to_numeric(
            df3['Net Working Capital_x'], errors='coerce')
        df3['Net Working Capital_y'] = pd.to_numeric(
            df3['Net Working Capital_y'], errors='coerce')
        df3['Revenue_x'] = pd.to_numeric(df3['Revenue_x'], errors='coerce')
        df3['Revenue_y'] = pd.to_numeric(df3['Revenue_y'], errors='coerce')
        df3['Capex AP Adjustment'] = pd.to_numeric(df3['Capex AP Adjustment'],
                                                   errors='coerce')
        df3.fillna(0.00, inplace=True)
        df3['Change in Working Capital'] = df3['Net Working Capital_x'] - df3[
            'Net Working Capital_y'] - df3['Capex AP Adjustment']
        df3['Change in Revenue'] = df3['Revenue_x'] - df3['Revenue_y']

        for n in [
                'Net Working Capital_x', 'Revenue_x', 'Capex AP Adjustment',
                'MinY', 'label1_x', 'label2', 'Net Working Capital_y',
                'Revenue_y', 'label1_y'
        ]:
            df3 = df3.drop(n, 1)
        result = df3[(df3['Change in Working Capital'] != 0) |
                     (df3['Change in Revenue'] != 0)]
        result.reset_index(drop=True)
        #result.to_csv("result.csv")
        return result

    def add_accounts1(df, df1):
        result = df.append(df1, ignore_index=True, sort=False)
        add_columns = ['Change in Working Capital', 'Change in Revenue']
        new_columns = [c for c in df.columns if c not in add_columns
                       ] + add_columns
        result = result[new_columns]
        return result

    def add_accounts(df, df1):
        add_columns = df1.columns.values[4:]
        result = df.append(df1, ignore_index=True, sort=False)
        new_columns = df.columns.values.tolist() + add_columns.tolist()
        result = result[new_columns]
        return result

    def add_CANCOI_COI(df):
        df['Corp Adj Net Core Op Income'] = pd.to_numeric(
            df['Corp Adj Net Core Op Income'], errors='coerce')
        df['Material Price Variance'] = pd.to_numeric(
            df['Material Price Variance'], errors='coerce')
        df['Tax Rate'] = [
            0.265 if o == 'FY 2015' else 0.24 if o == 'FY 2016' else 0.27
            for o in df.Period.values
        ]
        df['Corp Adj Net Core Op Income*(1-Tax Rate)'] = (
            1 - df['Tax Rate']) * df['Corp Adj Net Core Op Income']
        df['Core Operating Income*(1-Tax Rate)'] = (
            1 - df['Tax Rate']) * df['Material Price Variance']
        df = df.drop('Tax Rate', 1)
        return df



#计算POP 和POP Yearly
def cal_pop(path, file_name):
    Data = import_data(path)
    CIA = POPData.CCIA(Data)
    Data = POPData.add_accounts(Data, CIA)
    Data = POPData.ShaiXuan(Data, period_list, pop_period)
    Data['Avg ROIC Total Net Assets (Less Customer Gear)'] = Data['Avg ROIC Total Net Assets (Less Customer Gear)']*0
    ###############2022.Feb.23 update##################
    Data = add_AVGROIC(Data,cal_AVGROIC(path.replace('POPData','ROICData')))
    Data.to_csv('data.csv')
    ###################################################
    Mapping = createMapping(PCMapping)
    result = merge(Mapping, Data)
    result = result.rename(columns={'Cost Center': 'PLANT'})
    a = '8 Quarter (%s)' % (pop_period + " " + pop_version)
    if (a[-2] == '2') | (a[-2] == 'd'):
        result = POPData.clear_lockedversion(result)
    result.replace({'Version': '8 Quarter'}, a, inplace=True)
    result_path = os.getcwd() + r"\POP"
    if not os.path.isdir(result_path):
        os.mkdir(result_path)
    result_savepath = result_path + r"\%s.csv" % file_name
    result.to_csv(result_savepath, index=0)
    print("%s Done!" % file_name)


def cal_yearlypop(path1, path2,pathRY,file_name):
    YearlyData1 = POPYearly.CData1(import_data(path1))
    YearlyData2 = POPYearly.CData2(import_data(path2))
    CIA = POPYearly.CCIA(YearlyData1)
    YearlyData = POPYearly.add_accounts(YearlyData1, YearlyData2)
    YearlyData1plusCIA = POPYearly.add_accounts1(YearlyData, CIA)
    YearlyDataTotal = POPYearly.add_CANCOI_COI(YearlyData1plusCIA)
    YearlyDataTotal['Avg ROIC Total Net Assets (Less Customer Gear)'] = YearlyDataTotal['Avg ROIC Total Net Assets (Less Customer Gear)']*0
    YearlyDataTotal.to_csv('CHECK1.csv')
    ###############2022.Feb.28 update##################
    YearlyDataTotal = add_YEARLYAVGROIC(YearlyDataTotal,cal_YEARLYAVGROIC(pathRY.replace('POPData','ROICYearly')))
    YearlyDataTotal.to_csv('CHECK2.csv')
    ###################################################
    Mapping = createMapping(PCMapping)
    result = merge(Mapping, YearlyDataTotal)
    result = result.rename(columns={'Cost Center': 'PLANT'})
    result = result.reset_index(drop=True)  # 重新索引
    a = '8 Quarter (%s)' % (pop_period + " " + pop_version)
    result['Version'] = [
        a if o == 'FY 2022' else 'Actual' for o in result.Period.values
    ]
    result_path = os.getcwd() + r"\POPYearly"
    if not os.path.isdir(result_path):
        os.mkdir(result_path)
    result_savepath = result_path + r"\%s-yearly.csv" % file_name
    result.to_csv(result_savepath, index=0)
    print("%s Yearly Done!" % file_name)




def processtime(f):
    ####记录开始时间####
    starttime = datetime.datetime.now()
    f()
    endtime = datetime.datetime.now()
    print('用时: %d s' % ((endtime - starttime).seconds))  # 程序用时