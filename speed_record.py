import xlsxwriter
import openpyxl

from datetime import datetime
import threading
import smtplib
import os
import time
import csv
import schedule

csv_file_name = "ouput.csv"
excel_file_name = "output.xlsx"
speedtest_out = "speedtest_out"

#def parse_speedtest_file(file_path):
     

def speed_test_func():
    time = datetime.now()
    print("speed-test beginning ... {}".format(time))
    os.system("/home/pi/auto_speed_test/speedtest-cli > " + speedtest_out)
    time = datetime.now()
    print("speed-test done {:d}".format(time))
   
def csv_write_meta():
    with open(csv_file_name, 'a', newline='') as csv_file:
        fwrite = csv.writer(csv_file)
        fwrite.writerow(['a', 'b', 'c'])
        fwrite.writerow(['a1', 'b1', 'c1'])
        fwrite.writerow(['2a', 'b', 'c'])

def excel_write_meta():

    now = datetime.now()
    date_str = now.strftime("%Y%m%d")

    if (os.path.exists(excel_file_name) == True):
        wb = openpyxl.load_workbook(excel_file_name)
        if date_str in wb.sheetnames:
            write_sheet = wb[date_str]
            new_row = write_sheet.max_row + 1
            write_sheet['A' + str(new_row)] = now.strftime("%Y_%m_%d_%H_%M_%S")
        else:
            wb = openpyxl.Workbook()
            wb.create_sheet(date_str)
            write_sheet = wb[date_str]
            write_sheet['A1'] = now.strftime("%Y_%m_%d_%H_%M_%S")
    else:
        wb = openpyxl.Workbook()
        wb.create_sheet(date_str)
        write_sheet = wb[date_str]
        write_sheet['A1'] = now.strftime("%Y_%m_%d_%H_%M_%S")

    wb.save(excel_file_name)

def main():
    print("__ start ++")
    csv_write_meta()
    excel_write_meta()

    now = datetime.now()
    date_str = now.strftime("%Y%m%d")
    print("%s" % date_str)
    #timer_func()

    schedule.every(1).minute.do(speed_test_func)

    '''while True:
        schedule.run_pending()
        time.sleep(1)'''

    
    speed_test_func()

if __name__ == "__main__":
    main()
    
