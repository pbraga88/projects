#-*- coding: latin1 -*-
import speedTest
import xlsxwriter
from tkinter import *
import time
from threading import Thread
import os.path
import os
from tkinter import filedialog
import subprocess
from tkinter.messagebox import *
from subprocess import check_output

class Application:

    def __init__(self, master=None):
        self.controlBit = 0

#         root.geometry("1080x720+0+0") # width x height + x_offset + y_offset:
        root.geometry("300x250+50+50")
        self.fonte = ("Verdana", "8")
        root.resizable(0,0) # Disable 'Maximize' button

        self.conteiner1 = Frame(master)
        self.conteiner1["pady"] = 10
        self.conteiner1.pack()
        self.titulo = Label(self.conteiner1, text="Throughput/SSID analyzer")
        self.titulo["font"] = ("Verdana", "10", "bold")
        self.titulo.pack (side=LEFT)


        ''' Conteiner1 '''
        self.profile11 = Label(master, text="SSID:", font=self.fonte, bg="ghost white")
        self.profile11.place(x=5, y=30, width=90, height=20)

        self.profileEntry11 = Entry(master)
        self.profileEntry11.place(x=100, y=30, width=155, height=20)


        self.profile = Label(master, text="Interval (sec):", font=self.fonte, bg = "ghost white")
        self.profile.place(x = 5, y = 55, width=85, height=20)

        self.profileEntry = Entry(master)
        self.profileEntry.place(x = 100, y = 55, width=155, height=20)


        ''' Conteiner 2 '''
        self.filePath = Label(master, text="File Path:", font=self.fonte, bg = "ghost white")
        self.filePath.place(x = 5, y = 80, width=85, height=20)

        self.filePathEntry = Entry(master)
        self.filePathEntry.place(x = 100, y = 80, width=100, height=20)

        #Conteiner 2.1
        self.browseButton = Button(master, text = "Browse", font = ["Verdana", "7"])
        self.browseButton["command"] = self.browseFileLoc
        self.browseButton.place(x = 205, y = 80, width=50, height=20)


        ''' Conteiner 3 '''
        self.startButton = Button(master, text = "Start", font = ["Verdana", "9"])
        self.startButton["command"] = self.runner
        self.startButton.place(x = 100, y = 105, width=100, height=20)


        self.led17 =Label(master, bg = "orange red")
        self.led17.place(x = 220, y = 107, width=16, height=16)


        ''' Conteiner 4 '''
        self.profile = Label(master, text="Output:", font=["Verdana", "8"])
        self.profile.place(x = 30, y = 150, width=45, height=20)

        self.log1 = Entry(master)
        self.log1.place(x = 30, y = 170, width=250, height=20)
        self.log2 = Entry(master)
        self.log2.place(x = 30, y = 190, width=250, height=20)
        self.log3 = Entry(master)
        self.log3.place(x = 30, y = 210, width=250, height=20)


    def browseFileLoc(self):
        self.filePathEntry.delete(0, END)
        currdir = os.getcwd()
        fLoc = filedialog.askdirectory(parent=root, initialdir=currdir, title='Please select a directory')
        self.filePathEntry.insert(INSERT, fLoc)

    def Start(self):
        x=1
        z=0
        lista = list()
        fileLoc = self.filePathEntry.get()
        print(fileLoc)
        wbLocation = "%s\log.xlsx"%(fileLoc)
        if os.path.isdir(fileLoc) == False:
            print("path not found")
            self.led17.config(bg = 'orange red')
            self.startButton.config(text = 'Start')
            # self.__init__
            return
        if os.path.exists(wbLocation) == True:
            if askyesno('Warning', 'There is already a log.xlsx file in the folder. Want to overwrite it?!'):
                pass
            else:
                self.led17.config(bg = 'orange red')
                self.startButton.config(text = 'Start')
                return
        print(wbLocation)
        self.workbook = xlsxwriter.Workbook(wbLocation)
        worksheet = self.workbook.add_worksheet()
        worksheet.set_column('A:A', 25)
        worksheet.set_column('B:D', 25)
        worksheet.set_column('E:E', 25)
        worksheet.add_table('A1:C1000', {'header_row':True,'banded_rows': True, 'columns':[{'header':'Time'},
                                                                                         {'header': 'Download (Mbit/s)'},
                                                                                         {'header': 'Upload (Mbit/s)'},
                                                                                         ]})
        while True:
            self.controlBit = 0 # Bit de Controle (Initialize)
            # self.networkControlBit = 0 # Bit de controle da conexão
            try:
                timeToWait = int(self.profileEntry.get())
                if timeToWait < 30:
                    print('Please, enter a interval number >=60sec!')
                    self.led17.config(bg='orange red')
                    self.startButton.config(text='Start')
                    self.workbook.close()
                    self.controlBit = 0
                    return
                # Checking the connection. Timeout on failure = 30 secs
                self.checkConnection = str(check_output("netsh wlan connect name=%s" %(self.profileEntry11.get())))

                timeEnd = time.time()+30
                while(time.time()<timeEnd):
                    self.checkInterface = str(check_output("netsh wlan show interfaces"))
                    if self.profileEntry11.get() in self.checkInterface and self.profileEntry11.get() != '':
                        self.controlBit = 1
                        print('Connected')
                        # subprocess.call("netsh wlan connect name=Pace-Corp", shell=False)
                        time.sleep(10)
                        break
                    else:
                        print('Trying to Connect')
                        time.sleep(1)
                self.checkPing()
                if self.controlBit == 1:
                    self.networkControlBit = 1
                    link = speedTest.speedtest()

                    # output
                    self.log1.delete(0, END)
                    self.log2.delete(0, END)
                    self.log3.delete(0, END)
                    self.log1.insert(INSERT, time.ctime())
                    self.log2.insert(INSERT, speedTest.speedtest.Down)
                    self.log3.insert(INSERT, speedTest.speedtest.Up)

                    for itens in link:
                        lista.append(itens[0:4])

                    worksheet.write(x,z, time.ctime())
                    worksheet.write(x,z+1, lista[0])
                    worksheet.write(x,z+2, lista[1])
                    x+=1
                    z=0
                    lista = list()
                else:
                    self.log1.delete(0, END)
                    self.log2.delete(0, END)
                    self.log3.delete(0, END)
                    self.log1.insert(INSERT, time.ctime())
                    self.log2.insert(INSERT, 'Fail')
                    self.log3.insert(INSERT, 'Fail')
                    worksheet.write(x,z, time.ctime())
                    worksheet.write(x,z+1, 'Fail')
                    worksheet.write(x,z+2, 'Fail')
                    x+=1
                    z=0

                if "Conexão de Rede sem Fio" in self.checkInterface:
                    subprocess.call('netsh wlan disconnect interface="Conexão de Rede sem Fio"', shell=False)
                elif "Wireless Network Connection" in self.checkInterface:
                    subprocess.call('netsh wlan disconnect interface="Wireless Network Connection"', shell=False)
                else:
                    subprocess.call("netsh wlan connect name=Pace-Corp", shell=False) # Disconnect from current network

                timeEnd = time.time() + timeToWait
                self.controlBit = 0 # Bit de controle (Finalize)
                print('start counter')
                while (time.time() < timeEnd):
                    print(timeToWait)

                    time.sleep(1)
                    timeToWait-=1
                    if self.led17['bg'] == 'orange red':
                        self.workbook.close()
                        #self.__init__
                        return

            except subprocess.CalledProcessError:
                if self.networkControlBit == 0:
                    if askyesno('Warning', 'Network not found. Do you want continue?'):
                        self.networkControlBit = 1
                        pass
                    else:
                        self.led17.config(bg='orange red')
                        self.startButton.config(text='Start')
                        return
                print('Network not found')

                self.log1.delete(0, END)
                self.log2.delete(0, END)
                self.log3.delete(0, END)
                self.log1.insert(INSERT, time.ctime())
                self.log2.insert(INSERT, 'Network not found')
                self.log3.insert(INSERT, 'Network not found')
                worksheet.write(x, z, time.ctime())
                worksheet.write(x, z + 1, 'Network not found')
                worksheet.write(x, z + 2, 'Network not found')
                x += 1
                z = 0

                timeEnd = time.time() + timeToWait
                self.controlBit = 0  # Bit de controle (Finalize)
                print('start counter')
                while (time.time() < timeEnd):
                    print(int(timeEnd - time.time()))
                    time.sleep(1)
                    timeToWait -= 1
                    if self.led17['bg'] == 'orange red':
                        self.workbook.close()
                        return

            except TypeError:
                print('Please, enter with a valid Interval number (Integer value)')
                self.led17.config(bg = 'orange red')
                self.startButton.config(text = 'Start')
                self.workbook.close()
                self.controlBit = 0
                # self.__init__
                return


            ''' Work on this exception later'''
#             except PermissionError:
#                 print('log.xlsx is opened, please close it before start the test.')

    def runner(self):
        if self.controlBit == 1:
            showwarning('Warning!', 'Retrieving data from server. Please Wait!')
            return

        if self.led17['bg'] == 'orange red':
            self.networkControlBit = 0
            self.led17.config(bg = 'green')
            self.startButton.config(text = 'Stop')
            Thread(target=self.Start).start()

        elif self.led17['bg'] == 'green':
            self.led17.config(bg = 'orange red')
            self.startButton.config(text = 'Start')
            self.workbook.close()
            # self.__init__
            return
    def checkPing(self):
        # try:
        timeOut = time.time() + 30
        while (time.time() < timeOut):
            try:
                checkInterface = str(check_output("ping www.google.com -4 -n 1"))
                if 'Recebidos = 1' in checkInterface:
                    print('Ping OK')
                    # self.controlBit = '1'
                    return
                else:
                    print('Trying to ping host: 200.221.2.45')
                    time.sleep(1)
            except:
                print("Exception: Ping failure")
                time.sleep(1)
        print('Ping failure')
        self.controlBit = 0
        return
        # except:
        #     print("Exception: Ping failure")
        #     self.controlBit = '0'
        #     return
root = Tk()
Application(root)
root.mainloop()