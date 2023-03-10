from threading import Thread
from threading import Timer
import pymysql
import psutil
from datetime import date
import datetime
import openpyxl
import cx_Oracle
import time
import subprocess
from tkinter import *


def bdoracle(quvery):
    connection = cx_Oracle.connect(user="SYSDBA", password="ssxrtx3198", dsn="MSDAORA.1/someBase", encoding="UTF-8")
    for i in quvery:
        cur = connection.cursor()
        cur.execute(query= i)
    connection.close()


def bd(quvery):
    con = pymysql.connect(host='localhost', user='root', password='1234', db='marlo')  # конект к бд
    for i in quvery:
        cur = con.cursor()
        cur.arraysize = 56000
        print(i)
        cur.execute(query= i)
    con.close()


def ReadTxtAndBackMassive():
    massive = []
    with open("DecrypedData.txt","r") as file:
        for line in file:
            lineClear = line.replace("\n","")
            massive.append(lineClear)
    return massive


def readxlsx():
    wb = openpyxl.load_workbook('sql.xlsx')
    sheet = wb.active
    xlsx = []
    for i in range(sheet.max_row):
        xlsx.append(str(sheet["A" + str(i+1)].value))
    return xlsx




def log(num,nummax,sended_request):
    current_date = date.today()
    dt_now = datetime.datetime.now()
    f = open(str(current_date)+'.txt','a')
    f.write(str(dt_now) + " active_thread: " + str(num)  + " max_thread: " + str(nummax) + " sended_request: " + str(sended_request) + "\n")
    f.close()


if __name__ == '__main__':
    #test()

    subprocess.Popen('Project1.exe')
    time.sleep(5)
    ReadTxtAndBackMassive()
    xlsx = readxlsx()
    i = 0
    T = 500 #стартовое число потоков
    countermax = T
    sended_request = 0
    while True:
         if psutil.cpu_percent(interval= 0.1) > 90: #верхний порог нагрузки
             T -= 1
         if psutil.cpu_percent(interval= 0.1) < 90: #верхний порог нагрузки
             T += 1
         if T > countermax: #
             counter = T
         if T <= 0:
             T = 1
         time = Timer(5,log,args=(T,countermax,sended_request,))
         time.start()

         threads = []
         for n in range(int(T)):
             t = Thread(target=bdoracle, args=(xlsx,),daemon=False)
             t.start()
             threads.append(t)
         for t in threads:
             t.join()
             sended_request = sended_request + 1




