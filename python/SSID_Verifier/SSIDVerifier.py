import subprocess
from subprocess import check_output, check_call
import time
import xlsxwriter
import xlrd
import tkinter
from tkinter import *
import time
from subprocess import check_output
from threading import Thread
import threading
import os.path
import pathlib
from builtins import dir
import os
from pathlib import PurePath
from tkinter import filedialog
import subprocess
from subprocess import check_output, check_call
from tkinter.messagebox import *
    
class Application:
    
    def __init__(self, master=None):
        self.controlBit = 0
#         root.geometry("1080x720+0+0") # width x height + x_offset + y_offset:
        root.geometry("1050x400+50+50")
        self.fonte = ("Verdana", "8") 
        root.resizable(0,0) # Disable 'Maximize' button
           
        self.conteiner1 = Frame(master)
        self.conteiner1["pady"] = 10
        self.conteiner1.pack()
        self.titulo = Label(self.conteiner1, text="SSID Verifier") 
        self.titulo["font"] = ("Verdana", "12", "bold") 
        self.titulo.pack (side=LEFT) 
  
  
        ''' Conteiner1 '''
        self.profile11 = Label(master, text="SSID:", font=self.fonte, bg = "ghost white") 
        self.profile11.place(x = 5, y = 55, width=90, height=20)      
            
        self.profileEntry11 = Entry(master) 
        self.profileEntry11.place(x = 100, y = 55, width=155, height=20)
        
        self.profile12 = Label(master, text="Output:", font=["Verdana", "8"]) 
        self.profile12.place(x = 5, y = 80, width=45, height=20)

        self.log11 = Entry(master) 
        self.log11.place(x = 5, y = 105, width=250, height=20)        
        self.log12 = Entry(master) 
        self.log12.place(x = 5, y = 125, width=250, height=20)
        self.log13 = Entry(master) 
        self.log13.place(x = 5, y = 145, width=250, height=20) 
        
        ''' Conteiner1.2 '''
        self.profile21 = Label(master, text="SSID:", font=self.fonte, bg = "ghost white") 
        self.profile21.place(x = 265, y = 55, width=90, height=20)      
            
        self.profileEntry21 = Entry(master) 
        self.profileEntry21.place(x = 360, y = 55, width=155, height=20)
        
        self.profile22 = Label(master, text="Output:", font=["Verdana", "8"]) 
        self.profile22.place(x = 265, y = 80, width=45, height=20)

        self.log21 = Entry(master) 
        self.log21.place(x = 265, y = 105, width=250, height=20)        
        self.log22 = Entry(master) 
        self.log22.place(x = 265, y = 125, width=250, height=20)
        self.log23 = Entry(master) 
        self.log23.place(x = 265, y = 145, width=250, height=20) 

        ''' Conteiner1.3 '''
        self.profile32 = Label(master, text="SSID:", font=self.fonte, bg = "ghost white") 
        self.profile32.place(x = 525, y = 55, width=90, height=20)      
            
        self.profileEntry31 = Entry(master) 
        self.profileEntry31.place(x = 620, y = 55, width=155, height=20)
        
        self.profile32 = Label(master, text="Output:", font=["Verdana", "8"]) 
        self.profile32.place(x = 525, y = 80, width=45, height=20)

        self.log31 = Entry(master) 
        self.log31.place(x = 525, y = 105, width=250, height=20)        
        self.log32 = Entry(master) 
        self.log32.place(x = 525, y = 125, width=250, height=20)
        self.log33 = Entry(master) 
        self.log33.place(x = 525, y = 145, width=250, height=20) 
        
        ''' Conteiner1.4 '''
        self.profile41 = Label(master, text="SSID:", font=self.fonte, bg = "ghost white") 
        self.profile41.place(x = 785, y = 55, width=90, height=20)      

        self.profileEntry41 = Entry(master) 
        self.profileEntry41.place(x = 880, y = 55, width=155, height=20)
        
        self.profile42 = Label(master, text="Output:", font=["Verdana", "8"]) 
        self.profile42.place(x = 785, y = 80, width=45, height=20)

        self.log41 = Entry(master) 
        self.log41.place(x = 785, y = 105, width=250, height=20)        
        self.log42 = Entry(master) 
        self.log42.place(x = 785, y = 125, width=250, height=20)
        self.log43 = Entry(master) 
        self.log43.place(x = 785, y = 145, width=250, height=20) 

        ''' Conteiner1.5 '''
        self.profile51 = Label(master, text="SSID:", font=self.fonte, bg = "ghost white") 
        self.profile51.place(x = 5, y = 185, width=90, height=20)      
            
        self.profileEntry51 = Entry(master) 
        self.profileEntry51.place(x = 100, y = 185, width=155, height=20)
        
        self.profile52 = Label(master, text="Output:", font=["Verdana", "8"]) 
        self.profile52.place(x = 5, y = 210, width=45, height=20)

        self.log51 = Entry(master) 
        self.log51.place(x = 5, y = 235, width=250, height=20)        
        self.log52 = Entry(master) 
        self.log52.place(x = 5, y = 255, width=250, height=20)
        self.log53 = Entry(master) 
        self.log53.place(x = 5, y = 275, width=250, height=20) 
        
        ''' Conteiner1.6 '''
        self.profile61 = Label(master, text="SSID:", font=self.fonte, bg = "ghost white") 
        self.profile61.place(x = 265, y = 185, width=90, height=20)      
            
        self.profileEntry61 = Entry(master) 
        self.profileEntry61.place(x = 360, y = 185, width=155, height=20)
        
        self.profile62 = Label(master, text="Output:", font=["Verdana", "8"]) 
        self.profile62.place(x = 265, y = 210, width=45, height=20)

        self.log61 = Entry(master) 
        self.log61.place(x = 265, y = 235, width=250, height=20)        
        self.log62 = Entry(master) 
        self.log62.place(x = 265, y = 255, width=250, height=20)
        self.log63 = Entry(master) 
        self.log63.place(x = 265, y = 275, width=250, height=20) 

        ''' Conteiner1.7 '''
        self.profile71 = Label(master, text="SSID:", font=self.fonte, bg = "ghost white") 
        self.profile71.place(x = 525, y = 185, width=90, height=20)      
            
        self.profileEntry71 = Entry(master) 
        self.profileEntry71.place(x = 620, y = 185, width=155, height=20)
        
        self.profile72 = Label(master, text="Output:", font=["Verdana", "8"]) 
        self.profile72.place(x = 525, y = 210, width=45, height=20)

        self.log71 = Entry(master) 
        self.log71.place(x = 525, y = 235, width=250, height=20)        
        self.log72 = Entry(master) 
        self.log72.place(x = 525, y = 255, width=250, height=20)
        self.log73 = Entry(master) 
        self.log73.place(x = 525, y = 275, width=250, height=20) 
        
        ''' Conteiner1.8 '''
        self.profile81 = Label(master, text="SSID:", font=self.fonte, bg = "ghost white") 
        self.profile81.place(x = 785, y = 185, width=90, height=20)      

        self.profileEntry81 = Entry(master) 
        self.profileEntry81.place(x = 880, y = 185, width=155, height=20)
        
        self.profile82 = Label(master, text="Output:", font=["Verdana", "8"]) 
        self.profile82.place(x = 785, y = 210, width=45, height=20)

        self.log81 = Entry(master) 
        self.log81.place(x = 785, y = 235, width=250, height=20)        
        self.log82 = Entry(master) 
        self.log82.place(x = 785, y = 255, width=250, height=20)
        self.log83 = Entry(master) 
        self.log83.place(x = 785, y = 275, width=250, height=20) 


  
        ''' Conteiner 2 (File Path) '''
        self.filePath = Label(master, text="File Path:", font=self.fonte, bg = "ghost white") 
        self.filePath.place(x = 5, y = 320, width=90, height=20)      
             
        self.filePathEntry = Entry(master) 
        self.filePathEntry.place(x = 100, y = 320, width=100, height=20)

        #Conteiner 2.1
        self.browseButton = Button(master, text = "Browse", font = ["Verdana", "7"])
        self.browseButton["command"] = self.browseFileLoc
        self.browseButton.place(x = 205, y = 320, width=50, height=20)
                      
                      
        ''' Conteiner 3 (Interval) '''
        self.Interval = Label(master, text="Interval (sec):", font=self.fonte, bg = "ghost white") 
        self.Interval.place(x = 5, y = 345, width=90, height=20)     
             
        self.IntervalEntry = Entry(master) 
        self.IntervalEntry.place(x = 100, y = 345, width=155, height=20)

        ''' Conteiner 4 '''
        self.startButton = Button(master, text = "Start", font = ["Verdana", "9"])
        self.startButton["command"] = self.runner
        self.startButton.place(x = 100, y = 370, width=100, height=20)
             
             
        self.led17 =Label(master, bg = "orange red")
        self.led17.place(x = 220, y = 370, width=16, height=16)

        
    def browseFileLoc(self):
        self.filePathEntry.delete(0, END)
        currdir = os.getcwd()
        fLoc = filedialog.askdirectory(parent=root, initialdir=currdir, title='Please select a directory')       
        self.filePathEntry.insert(INSERT, fLoc)
        
    def runner(self):
#         if self.controlBit == 1:
#             showwarning('Warning!', 'Retrieving data from server. Please Wait!')
#             return
#                    
        if self.led17['bg'] == 'orange red':
            self.led17.config(bg = 'green')
            self.startButton.config(text = 'Stop')
            Thread(target=self.Start).start()         

        elif self.led17['bg'] == 'green':
            self.led17.config(bg = 'orange red')
            self.startButton.config(text = 'Start')
            self.workbook.close()
            self.__init__
            return
    
    def Start(self):
        x=1
        z=0
        fileLoc = self.filePathEntry.get()
#         print(fileLoc)
        wbLocation = "%s\log.xlsx"%(fileLoc)        
        if os.path.isdir(fileLoc) == False:
            print("path not found")
            self.led17.config(bg = 'orange red')
            self.startButton.config(text = 'Start')
            self.__init__
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
        worksheet.add_table('A1:C500000', {'header_row':True,'banded_rows': True, 'columns':[{'header':'Time'},
                                                                                         {'header': 'SSID'},
                                                                                         {'header': 'Status'},
                                                                                         ]})
                
        while True:
            try:
                checkInterface = str(check_output("netsh wlan show network"))
                if self.profileEntry11.get() in checkInterface and self.profileEntry11.get() !='':        
                    self.log11.delete(0, END)
                    self.log12.delete(0, END)
                    self.log13.delete(0, END)
                    self.log11.insert(INSERT, time.ctime())
                    self.log12.insert(INSERT, self.profileEntry11.get())
                    self.log13.insert(INSERT, 'OK')    
                    worksheet.write(x,z, time.ctime())
                    worksheet.write(x,z+1, self.profileEntry11.get())
                    worksheet.write(x,z+2, 'OK')
                    x+=1
                    z=0
                    print('%s available'%(self.profileEntry11.get()))
                               
                elif self.profileEntry11.get() =='':
#                     print('Please enter a valid Network SSID')
                    self.log11.delete(0, END)
                    self.log12.delete(0, END)
                    self.log13.delete(0, END)                    
                else:        
                    self.log11.delete(0, END)    
                    self.log12.delete(0, END)
                    self.log13.delete(0, END)
                    self.log11.insert(INSERT, time.ctime())
                    self.log12.insert(INSERT, self.profileEntry11.get())
                    self.log13.insert(INSERT, 'Fail')
                    
                    worksheet.write(x,z, time.ctime())
                    worksheet.write(x,z+1, self.profileEntry11.get())
                    worksheet.write(x,z+2, 'Fail')
                    x+=1
                    z=0
                    print('%s NOT available'%(self.profileEntry11.get()))
                    
                if self.profileEntry21.get() in checkInterface and self.profileEntry21.get() !='':        
                    self.log21.delete(0, END)
                    self.log22.delete(0, END)
                    self.log23.delete(0, END)
                    self.log21.insert(INSERT, time.ctime())
                    self.log22.insert(INSERT, self.profileEntry21.get())
                    self.log23.insert(INSERT, 'OK')
                    worksheet.write(x,z, time.ctime())
                    worksheet.write(x,z+1, self.profileEntry21.get())
                    worksheet.write(x,z+2, 'OK')
                    x+=1
                    z=0
                    print('%s available'%(self.profileEntry21.get())) 
                elif self.profileEntry21.get() =='':
#                     print('Please enter a valid Network SSID')
                    self.log21.delete(0, END)
                    self.log22.delete(0, END)
                    self.log23.delete(0, END)                    
                else:        
                    self.log21.delete(0, END)    
                    self.log22.delete(0, END)
                    self.log23.delete(0, END)
                    self.log21.insert(INSERT, time.ctime())
                    self.log22.insert(INSERT, self.profileEntry21.get())
                    self.log23.insert(INSERT, 'Fail')
                    
                    worksheet.write(x,z, time.ctime())
                    worksheet.write(x,z+1, self.profileEntry21.get())
                    worksheet.write(x,z+2, 'Fail')
                    x+=1
                    z=0
                    print('%s NOT available'%(self.profileEntry21.get()))

                if self.profileEntry31.get() in checkInterface and self.profileEntry31.get() !='':        
                    self.log31.delete(0, END)
                    self.log32.delete(0, END)
                    self.log33.delete(0, END)
                    self.log31.insert(INSERT, time.ctime())
                    self.log32.insert(INSERT, self.profileEntry31.get())
                    self.log33.insert(INSERT, 'OK')
                    worksheet.write(x,z, time.ctime())
                    worksheet.write(x,z+1, self.profileEntry31.get())
                    worksheet.write(x,z+2, 'OK')
                    x+=1
                    z=0
                    print('%s available'%(self.profileEntry31.get())) 
                elif self.profileEntry31.get() =='':
#                     print('Please enter a valid Network SSID')
                    self.log31.delete(0, END)
                    self.log32.delete(0, END)
                    self.log33.delete(0, END)                    
                else:        
                    self.log31.delete(0, END)    
                    self.log32.delete(0, END)
                    self.log33.delete(0, END)
                    self.log31.insert(INSERT, time.ctime())
                    self.log32.insert(INSERT, self.profileEntry31.get())
                    self.log33.insert(INSERT, 'Fail')
                    
                    worksheet.write(x,z, time.ctime())
                    worksheet.write(x,z+1, self.profileEntry31.get())
                    worksheet.write(x,z+2, 'Fail')
                    x+=1
                    z=0
                    print('%s NOT available'%(self.profileEntry31.get()))
                    
                if self.profileEntry41.get() in checkInterface and self.profileEntry41.get() !='':        
                    self.log41.delete(0, END)
                    self.log42.delete(0, END)
                    self.log43.delete(0, END)
                    self.log41.insert(INSERT, time.ctime())
                    self.log42.insert(INSERT, self.profileEntry41.get())
                    self.log43.insert(INSERT, 'OK')
                    worksheet.write(x,z, time.ctime())
                    worksheet.write(x,z+1, self.profileEntry41.get())
                    worksheet.write(x,z+2, 'OK')
                    x+=1
                    z=0
                    print('%s available'%(self.profileEntry41.get())) 
                elif self.profileEntry41.get() =='':
#                     print('Please enter a valid Network SSID')
                    self.log41.delete(0, END)
                    self.log42.delete(0, END)
                    self.log43.delete(0, END)                    
                else:        
                    self.log41.delete(0, END)    
                    self.log42.delete(0, END)
                    self.log43.delete(0, END)
                    self.log41.insert(INSERT, time.ctime())
                    self.log42.insert(INSERT, self.profileEntry41.get())
                    self.log43.insert(INSERT, 'Fail')
                    
                    worksheet.write(x,z, time.ctime())
                    worksheet.write(x,z+1, self.profileEntry41.get())
                    worksheet.write(x,z+2, 'Fail')
                    x+=1
                    z=0
                    print('%s NOT available'%(self.profileEntry41.get()))

                if self.profileEntry51.get() in checkInterface and self.profileEntry51.get() !='':        
                    self.log51.delete(0, END)
                    self.log52.delete(0, END)
                    self.log53.delete(0, END)
                    self.log51.insert(INSERT, time.ctime())
                    self.log52.insert(INSERT, self.profileEntry51.get())
                    self.log53.insert(INSERT, 'OK')
                    worksheet.write(x,z, time.ctime())
                    worksheet.write(x,z+1, self.profileEntry51.get())
                    worksheet.write(x,z+2, 'OK')
                    x+=1
                    z=0
                    print('%s available'%(self.profileEntry51.get())) 
                elif self.profileEntry51.get() =='':
#                     print('Please enter a valid Network SSID')
                    self.log51.delete(0, END)
                    self.log52.delete(0, END)
                    self.log53.delete(0, END)                    
                else:        
                    self.log51.delete(0, END)    
                    self.log52.delete(0, END)
                    self.log53.delete(0, END)
                    self.log51.insert(INSERT, time.ctime())
                    self.log52.insert(INSERT, self.profileEntry51.get())
                    self.log53.insert(INSERT, 'Fail')
                    
                    worksheet.write(x,z, time.ctime())
                    worksheet.write(x,z+1, self.profileEntry51.get())
                    worksheet.write(x,z+2, 'Fail')
                    x+=1
                    z=0
                    print('%s NOT available'%(self.profileEntry51.get()))

                if self.profileEntry61.get() in checkInterface and self.profileEntry61.get() !='':        
                    self.log61.delete(0, END)
                    self.log62.delete(0, END)
                    self.log63.delete(0, END)
                    self.log61.insert(INSERT, time.ctime())
                    self.log62.insert(INSERT, self.profileEntry61.get())
                    self.log63.insert(INSERT, 'OK')
                    worksheet.write(x,z, time.ctime())
                    worksheet.write(x,z+1, self.profileEntry61.get())
                    worksheet.write(x,z+2, 'OK')
                    x+=1
                    z=0
                    print('%s available'%(self.profileEntry61.get())) 
                elif self.profileEntry61.get() =='':
#                     print('Please enter a valid Network SSID')
                    self.log61.delete(0, END)
                    self.log62.delete(0, END)
                    self.log63.delete(0, END)                    
                else:        
                    self.log61.delete(0, END)    
                    self.log62.delete(0, END)
                    self.log63.delete(0, END)
                    self.log61.insert(INSERT, time.ctime())
                    self.log62.insert(INSERT, self.profileEntry61.get())
                    self.log63.insert(INSERT, 'Fail')
                    
                    worksheet.write(x,z, time.ctime())
                    worksheet.write(x,z+1, self.profileEntry61.get())
                    worksheet.write(x,z+2, 'Fail')
                    x+=1
                    z=0
                    print('%s NOT available'%(self.profileEntry61.get()))          

                if self.profileEntry71.get() in checkInterface and self.profileEntry71.get() !='':        
                    self.log71.delete(0, END)
                    self.log72.delete(0, END)
                    self.log73.delete(0, END)
                    self.log71.insert(INSERT, time.ctime())
                    self.log72.insert(INSERT, self.profileEntry71.get())
                    self.log73.insert(INSERT, 'OK')
                    worksheet.write(x,z, time.ctime())
                    worksheet.write(x,z+1, self.profileEntry71.get())
                    worksheet.write(x,z+2, 'OK')
                    x+=1
                    z=0
                    print('%s available'%(self.profileEntry71.get())) 
                elif self.profileEntry71.get() =='':
#                     print('Please enter a valid Network SSID')
                    self.log71.delete(0, END)
                    self.log72.delete(0, END)
                    self.log73.delete(0, END)                    
                else:        
                    self.log71.delete(0, END)    
                    self.log72.delete(0, END)
                    self.log73.delete(0, END)
                    self.log71.insert(INSERT, time.ctime())
                    self.log72.insert(INSERT, self.profileEntry71.get())
                    self.log73.insert(INSERT, 'Fail')
                    
                    worksheet.write(x,z, time.ctime())
                    worksheet.write(x,z+1, self.profileEntry71.get())
                    worksheet.write(x,z+2, 'Fail')
                    x+=1
                    z=0
                    print('%s NOT available'%(self.profileEntry71.get()))         
          
                if self.profileEntry81.get() in checkInterface and self.profileEntry81.get() !='':        
                    self.log81.delete(0, END)
                    self.log82.delete(0, END)
                    self.log83.delete(0, END)
                    self.log81.insert(INSERT, time.ctime())
                    self.log82.insert(INSERT, self.profileEntry81.get())
                    self.log83.insert(INSERT, 'OK')
                    worksheet.write(x,z, time.ctime())
                    worksheet.write(x,z+1, self.profileEntry81.get())
                    worksheet.write(x,z+2, 'OK')
                    x+=1
                    z=0
                    print('%s available'%(self.profileEntry81.get())) 
                elif self.profileEntry81.get() =='':
#                     print('Please enter a valid Network SSID')
                    self.log81.delete(0, END)
                    self.log82.delete(0, END)
                    self.log83.delete(0, END)                    
                else:        
                    self.log81.delete(0, END)    
                    self.log82.delete(0, END)
                    self.log83.delete(0, END)
                    self.log81.insert(INSERT, time.ctime())
                    self.log82.insert(INSERT, self.profileEntry81.get())
                    self.log83.insert(INSERT, 'Fail')
                    
                    worksheet.write(x,z, time.ctime())
                    worksheet.write(x,z+1, self.profileEntry81.get())
                    worksheet.write(x,z+2, 'Fail')
                    x+=1
                    z=0
                    print('%s NOT available'%(self.profileEntry81.get()))  

                
                timeToWait = int(self.IntervalEntry.get())
                timeEnd = time.time() + timeToWait
                print('start counter')
                while (time.time() < timeEnd):
                    print(timeToWait)
                    time.sleep(1)
                    timeToWait-=1
                    if self.led17['bg'] == 'orange red':
                        self.workbook.close()
                        self.__init__
                        return 
                
            except subprocess.CalledProcessError:
                print('network not found')
                self.led17.config(bg = 'orange red')
                self.startButton.config(text = 'Start')
                self.workbook.close()
                self.controlBit = 0
                self.__init__
                return 
            except TypeError:
                print('Path not found')
                self.led17.config(bg = 'orange red')
                self.startButton.config(text = 'Start')
                self.workbook.close()
                self.controlBit = 0
                self.__init__
                return

root = Tk()
Application(root)
root.mainloop()