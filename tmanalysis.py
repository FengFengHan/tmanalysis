#coding=utf-8

import pandas as pd
import datetime

file_name= './backup_20160301.csv'
start_date = datetime.datetime(2016,2, 21,0,0,0)
end_date = datetime.datetime(2016, 2, 29,23,59,59)

def timemeter_analysis(file_name,start_date, end_date):
    f = open(file_name, 'r')
    lines = f.readlines()
    f.close()
    f = open('backup.csv', 'w')
    for line in lines:
        if line[1] != '#':
            f.write(line)
    f.close()

    items = pd.read_csv("backup.csv")
    items.columns = ['Time', 'Description', 'Label', 'Start', 'End', 'Comment', 'ID']
    dateparse = lambda x: pd.datetime.strptime(x, '%Y-%m-%d %H:%M:%S')
    items[['Start', 'End']] = items[['Start', 'End']].applymap(dateparse)
    items = items[items['Start'] > start_date]
    items = items[items['Start'] < end_date]
    items = items.sort_values(by = 'Start')

    groupped = items.groupby(items['Label'])
    times = groupped['Time'].agg('sum')
    def sumString(se):
        result = ""
        uniqueWord = set()
        for words in se:
            if not  words in uniqueWord:
                uniqueWord.add(words)
                result += words + "。 "
        return result
    contens = groupped['Description'].apply(sumString)

    result = pd.DataFrame(times)
    result['Contens'] = contens
    total_time = items['Time'].sum()
    get_rate = lambda x: str(int((x * 100)/total_time)) + '%'
    result['Rate'] = result['Time'].map(get_rate)
    def get_time(x):
        minute = int(x*60)
        hour = minute // 60
        minute %= 60
        if minute < 10:
            minute_s = str(0) + str(minute)
        else:
            minute_s = str(minute)
        return str(hour) + ':' + minute_s

    result['Time'] = result['Time'].map(get_time)
    result['Label'] = result.index
    #result.to_csv('result.csv', columns = ['Rate', 'Time', 'Contens'])
    writer = pd.ExcelWriter(path='result.xlsx',engine= 'openpyxl')
    result.to_excel(writer,sheet_name='sheet1',columns=['Label', 'Rate', 'Time', 'Contens'],
                    startrow = 2, index = False)
    wb= writer.book
    ws = wb.active
    ws.merge_cells('A1:D1')
    ws['A1'] = start_date.strftime('%Y.%m.%d') + "~" + end_date.strftime('%Y.%m.%d')
    ws.merge_cells('A2:B2')
    ws['A' + '2'] = 'Total'
    ws.merge_cells('C2:D2')
    ws['C' + '2'] = get_time(total_time)
    writer.save()

timemeter_analysis(file_name,start_date,end_date)