import random
from tkinter import *
from tkinter import ttk
from functools import partial
import psutil
import pymysql
import openpyxl
import subprocess
from threading import Thread
from threading import Timer
import datetime
from datetime import date
from multiprocessing import Process
import os
import cx_Oracle
import time
from tkinter.messagebox import showinfo,showerror,showwarning





def bdoracle(quvery):
    user = decryptedAccses[1]
    password = decryptedAccses[0]
    DB = "MSDAORA.1/" + decryptedAccses[2]
    connection = cx_Oracle.connect(user=user, password=password, dsn=DB, encoding="UTF-8")
    # ещё как вариант con = cx_Oracle.connect('username/password@localhost')
    for i in quvery:
        cur = connection.cursor()
        cur.execute(query=i)
    connection.close()


def ReadTxtAndBackMassive():
    massive = []
    #with open("../DecrypedDataByLine.txt","r") as file: #попровь Project1
    d =  open("DecrypedData.txt", "r")
    for line in d:
        lineClear = line.replace("\n","")
        massive.append(lineClear)
    d.close()
    f = open('DecrypedData.txt', 'w')
    f.close()
    #os.remove("DecrypedData.txt")
    #os.remove("../DecrypedDataByLine.txt")
    return massive

#лимит
maxLimit = 50

def GetCpuPersents():
    output = str(subprocess.check_output('wmic cpu get loadpercentage'))
    outEnd = int(output[24] + output[25])
    return  outEnd

def threads():
    zgluchka =0
    import time
    subprocess.Popen('Project1.exe')
    time.sleep(5)
    xlsx = readxlsx()
    i = 0
    T = 1  # стартовое число потоков
    countermax = T
    sended_request = 0
    while zgluchka == 0:
        #для того чтобы узнать GPU -> psutil.virtual_memory()[2]
        if GetCpuPersents() < maxLimit:
            T += 1
        if GetCpuPersents() > maxLimit:  # верхний порог нагрузки
            T -= 1
        if T > countermax:  #
            countermax = T
        if T <= 0:
            T = 1
        time = Timer(5, log, args=(T, countermax, sended_request,))
        time.start()

        threads = []
        for n in range(int(T)):
            t = Thread(target=bdoracle, args=(xlsx,), daemon=False)
            t.start()
            threads.append(t)
        for t in threads:
            t.join()
            sended_request = sended_request + 1

#MainThread = Thread(target=threads, args=(), daemon=False)
#p1 = Process(target=threads, daemon=False)


def proc_start():
    p_to_start = Process(target=threads,daemon=False)
    p_to_start.start()
    return p_to_start


def proc_stop(p_to_stop):
    p_to_stop.kill()




#окно нагрузки ЦПУ
class WindowCPU(Tk):

    #метод создания окна
    def __init__(self):
        super().__init__()
        #главная настройка окна
        self.title("Тестирование")  # заголовок окна
        self.geometry(
            '400x130')
        self.resizable(height=False, width=False)
        self.label1 = Label(self, text="В поле ниже введите максимальную\nнагрузку на ПК в процентах",
                            font=("Arial Bold", 12), justify='left')  # просто информационное поле
        self.label1.grid(column=0, row=0, padx=10)  # это пишется чтобы элемент был хотя-бы виден и настроен по сетке
        self.labelStatus = Label(self, text="Остановлен", font=("Arial Bold", 12), justify='right')
        self.labelStatus.grid(column=1, row=0,
                         padx=10)  # это поле будем изменять. Оно нужно для отображения статуса программыlabel2.grid(column=1,row=0,sticky="e")
        self.procents = Entry(self, width=20, font=("Arial Bold", 12))  # поле ввода
        self.procents.grid(column=0, row=1, pady=10, sticky="w", padx=10)
        btnStart = ttk.Button(self, text="Старт")
        btnStart.grid(column=0, row=2, pady=10, sticky="w", padx=10)
        btnStart["command"] = self.btnStart

        btnPause = ttk.Button(self, text="Пауза")
        btnPause.grid(column=0, row=2, pady=10)
        btnPause["command"] = self.btnPause

        btnStop = ttk.Button(self, text="Остановить")
        btnStop.grid(column=1, row=2, pady=10, sticky="w")
        btnStop["command"] = self.btnStop


       # self.protocol("WM_DELETE_WINDOW", self.on_closing)

   # 3def on_closing(self):
    #    os.remove("DecrypedData.txt")
    #старт теста
    def btnStart(self):
        self.labelStatus.config(text="Работает")
        maxLimit = int(self.procents.get())
        global p
        p = proc_start()



    #пауза
    def btnPause(self):
        self.labelStatus.config(text="На пауза")
        proc_stop(p)


    #стоп
    def btnStop(self):
        self.labelStatus.config(text="Остановлен")
        proc_stop(p)





def readxlsx():
    wb = openpyxl.load_workbook('sql.xlsx')
    sheet = wb.active
    xlsx = []
    for i in range(sheet.max_row):
        xlsx.append(str(sheet["A" + str(i + 1)].value))
    return xlsx

def bd(quvery):
    con = pymysql.connect(host='localhost', user='root', password='1234', db='marlo')  # конект к бд
    for i in quvery:
        cur = con.cursor()
        cur.arraysize = 56000
        print(i)
        cur.execute(query=i)
    con.close()

def log(num,nummax,sended_request):
    current_date = date.today()
    dt_now = datetime.datetime.now()
    f = open(str(current_date)+'.txt','a')
    f.write(str(dt_now) + " active_thread: " + str(num)  + " max_thread: " + str(nummax) + " sended_request: " + str(sended_request) + "\n")
    f.close()









#это код для запуска приложения, так сказать главное окно для начала переходов (костыли)
def click():
    sus = subprocess.Popen('Project1.exe')
    time.sleep(5)
    sus.terminate()
    decryptedAccses = ReadTxtAndBackMassive()
    xlsx = readxlsx()
    if (xlsx != [] and decryptedAccses != []):
        root.destroy()
        window = WindowCPU()
    else:
        showerror(title="Ошибка",
                  message="К сожалению первичная настройка приложения прошла не успешно\nПроверьте, всё-ли верно вы настроили")

decryptedAccses = []

if __name__ == '__main__':
        root = Tk()
        root.title("Окно запуска")
        root.geometry("250x200")
        open_button = ttk.Button(text="Запуск", command=click)
        open_button.pack(anchor="center", expand=1)
        root.mainloop()