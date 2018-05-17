#!/usr/bin/env python3
# coding: utf8

import openpyxl
import datetime
import MySQLdb
import os
from resources import Functions as f
import warnings

warnings.filterwarnings("ignore")


def main():
    connect = MySQLdb.connect(host='localhost', user='operator', password='operator', db='inet', charset='utf8')
    cursor = connect.cursor()
    
    day = datetime.datetime.now().date() - datetime.timedelta(days=1)
    wb = openpyxl.load_workbook('resources{}template.xlsx'.format(os.sep))
    
    f.sessions_report(wb, cursor)
    f.speed_report(wb, cursor)
    f.modems_report(wb, cursor)
    
    wb.save('out{}Отчет по ADSL за {}.xlsx'.format(os.sep, day.strftime('%d-%m-%Y')))
    connect.close()
    
    
if __name__ == '__main__':
    main()