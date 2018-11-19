# -*- coding: latin1 -*-

import xlsxwriter
from builtins import dir
import os
import xlrd
import xml.etree.ElementTree as ET
from tkinter import *
from csv import excel
import tkinter
import tkinter.filedialog
from tkinter import filedialog
from turtledemo.chaos import jumpto
from builtins import dir
import os.path
import pathlib
from pathlib import PurePath
import time

class Application:
    def __init__(self, master=None):
        self.fonte = ("Verdana", "8") 
      
        self.conteiner1 = Frame(master)
        #self.conteiner1["pady"] = 5
        self.conteiner1.pack()
        self.conteiner2 = Frame(master)
        #self.conteiner2["pady"] = 10
        self.conteiner2.pack()
        self.conteiner3 = Frame(master)
        #self.conteiner2["padx"] = 10
        self.conteiner3.pack()        
        
        self.conteiner4 = Frame(master)
        #self.conteiner4["pady"] = 5
        self.conteiner4.pack()
        self.conteiner5 = Frame(master)
        #self.conteiner5["padx"] = 50
        self.conteiner5.pack()
        self.conteiner6 = Frame(master)
        #self.conteiner6["pady"] = 5
        self.conteiner6.pack()
        self.conteiner7 = Frame(master)
        #self.conteiner7["pady"] = 5
        self.conteiner7.pack()
        self.conteiner8 = Frame(master)
        #self.conteiner8["pady"] = 5
        self.conteiner8.pack()
        self.conteiner9 = Frame(master)
        #self.conteiner9["pady"] = 5
        self.conteiner9.pack()
        
        #Conteiner01
        self.titulo = Label(self.conteiner1, text="Set paramaters") 
        self.titulo["font"] = ("Verdana", "9", "bold") 
        self.titulo.pack (side=LEFT) 
        
        #Conteiner02
        self.lblnome = Label(self.conteiner2, text="Shared directory [task]:", font=self.fonte, width=28) 
        self.lblnome.pack(side=LEFT) 

        self.dirPath = Entry(self.conteiner2) 
        self.dirPath["width"] = 50 
        self.dirPath["font"] = self.fonte 
        self.dirPath.pack(side=LEFT) 

        self.bye = Button(self.conteiner2) 
        self.bye["text"] = "Browse" 
        self.bye["font"] = ("Verdana", "9") 
        self.bye["width"] = 7 
        self.bye["command"] = self.browseDir
        self.bye.pack ()
        
        #Conteiner03
        self.lblnome = Label(self.conteiner3, text="Shared directory [library Tasks]:", font=self.fonte, width=28) 
        self.lblnome.pack(side=LEFT) 

        self.taskDirPath = Entry(self.conteiner3) 
        self.taskDirPath["width"] = 50 
        self.taskDirPath["font"] = self.fonte 
        self.taskDirPath.pack(side=LEFT) 

        self.bye = Button(self.conteiner3) 
        self.bye["text"] = "Browse" 
        self.bye["font"] = ("Verdana", "9") 
        self.bye["width"] = 7 
        self.bye["command"] = self.browseTaskDir
        self.bye.pack ()
        
        #Conteiner04
        self.lblnome = Label(self.conteiner4, text="File directory:", font=self.fonte, width=28) 
        self.lblnome.pack(side=LEFT)        

        self.fileLoc = Entry(self.conteiner4) 
        self.fileLoc["width"] = 50 
        self.fileLoc["font"] = self.fonte 
        self.fileLoc.pack(side=LEFT)
        
        self.bye = Button(self.conteiner4) 
        self.bye["text"] = "Browse" 
        self.bye["font"] = ("Verdana", "9") 
        self.bye["width"] = 7 
        self.bye["command"] = self.browseFileLoc
        self.bye.pack ()
        
        #Conteiner05
        self.lblnome = Label(self.conteiner5, text="File name:", font=self.fonte, width=28) 
        self.lblnome.pack(side=LEFT)        

        self.fileName = Entry(self.conteiner5) 
        self.fileName["width"] = 50 
        self.fileName["font"] = self.fonte 
        self.fileName.pack(side=LEFT)
        
        self.bye = Button(self.conteiner5) 
        self.bye["text"] = "Create" 
        self.bye["font"] = ("Verdana", "9", "bold") 
        self.bye["width"] = 7 
        self.bye["command"] = self.createExcel
        self.bye.pack ()        
        
        #Conteiner 06
        self.titulo = Label(self.conteiner6) 
        #self.titulo["font"] = ("Verdana", "9", "bold") 
        self.titulo.pack (side=LEFT)
        
        #Conteiner 07
        self.titulo = Label(self.conteiner7, text="File Update") 
        self.titulo["font"] = ("Verdana", "9", "bold") 
        self.titulo.pack (side=LEFT) 
        
        #Conteiner08
        self.lblnome = Label(self.conteiner8, text="File path:", font=self.fonte, width=28) 
        self.lblnome.pack(side=LEFT) 
        
        self.filePath = Entry(self.conteiner8) 
        self.filePath["width"] = 50 
        self.filePath["font"] = self.fonte 
        self.filePath.pack(side=LEFT)

        self.bye = Button(self.conteiner8) 
        self.bye["text"] = "Browse" 
        self.bye["font"] = ("Verdana", "9") 
        self.bye["width"] = 7 
        self.bye["command"] = self.browseFile
        self.bye.pack ()
        
        #Conteiner09
        self.bye = Button(self.conteiner9) 
        self.bye["text"] = "Update" 
        self.bye["font"] = ("Verdana", "9", "bold") 
        self.bye["width"] = 7 
        self.bye["command"] = self.updateFile
        self.bye.pack () 
        
        #Messages
        self.lblmsg = Label(self.conteiner9, text="", font=self.fonte)
        self.lblmsg.pack(side=BOTTOM)
        
        self.lblmsg2 = Label(self.conteiner6, text="", font=self.fonte)
        self.lblmsg2.pack(side=BOTTOM)
        
    def browseDir(self):
        self.dirPath.delete(0, END)
        currdir = os.getcwd()
        tempdir = filedialog.askdirectory(parent=root, initialdir=currdir, title='Please select a directory')        
        self.dirPath.insert(INSERT, tempdir)
    
    def browseTaskDir(self):
        self.taskDirPath.delete(0, END)
        currdir = os.getcwd()
        tempdir = filedialog.askdirectory(parent=root, initialdir=currdir, title='Please select a directory')        
        self.taskDirPath.insert(INSERT, tempdir)
    
    def browseFileLoc(self):
        self.fileLoc.delete(0, END)
        currdir = os.getcwd()
        fLoc = filedialog.askdirectory(parent=root, initialdir=currdir, title='Please select a directory')       
        self.fileLoc.insert(INSERT, fLoc)
    
    def browseFile(self):
        self.filePath.delete(0, END)
        fpath = filedialog.askopenfilename(filetypes = (("Excel files", "*.xlsx"), ("Template files", "*.type"), ("All files", "*")))
        self.filePath.insert(INSERT, fpath)
      
    def createExcel(self):    
        #Erase label2
        self.lblmsg2.config(text = '')
        #Erase label1
        self.lblmsg.config(text = '')
        
        self.lblmsg2.config(text = "Creating .xlsx file, please wait.", fg = 'Black')
        self.lblmsg2.update_idletasks()        
        try:            
            x=1
            wbLocation = self.fileLoc.get() 
            fName = self.fileName.get()
            
            if os.path.exists(wbLocation) == False:
                self.lblmsg2.config(text = "Please, check the path", fg = 'Red')
                return
       
            wbLocation = "%s\%s.xlsx" %(wbLocation, fName)
       
            sharedDir = self.dirPath.get()
            if os.path.isdir(sharedDir) == False:
                self.lblmsg2.config(text = "Please, check the path.", fg = 'Red')
                return
                           
            workbook = xlsxwriter.Workbook(wbLocation)
            worksheet = workbook.add_worksheet()
            worksheet.set_column('A:A', 70)
            worksheet.set_column('B:D', 20)
            worksheet.set_column('E:E', 25)
            worksheet.set_column('F:F', 40)
            worksheet.set_column('G:G', 50)
            worksheet.set_column('H:K', 20)
       
            for folder in os.listdir(sharedDir):
                print(folder)
                worksheet.write(x,0,folder)
                lenght = len(folder)
                y = lenght - 4
                folderxml = '%s.xml' % folder[0:y]                  
                filepath =r"%s\%s\%s" % (sharedDir, folder, folderxml)
               
                tree = ET.parse(filepath)
                root = tree.getroot()
                for measure_id in root.iter('measure_id'):
                    measID = str(measure_id.text)
                    print ("Measure ID: %s" %(measID))
                    worksheet.write(x,1,measID)
                for interval in root.iter('interval'):
                    interv = str(interval.text)
                    print ("Interval: %s" %(interv))
                    worksheet.write(x,2,interv)
                for maxruntime in root.iter('maxruntime'):
                    maxRunt = str(maxruntime.text)
                    print ("Max run time: %s" %(maxRunt))
                    worksheet.write(x,3,maxRunt)
                for resource in root.iter('resource'):
                    resour = str(resource.text)
                    print ("Resource: %s" %(resour))
                    worksheet.write(x,4,resour)   
                for automatic in root.iter('automatic'):
                    autom = str(automatic.text)
                    print ("Automatic: %s" %(autom))
                    worksheet.write(x,5,autom) 
                for block in root.iter('block'):
                    blockID = str(block.text)
                    print ("BlockID: %s" %(blockID))
                    worksheet.write(x,6,blockID)
                       
                directory = self.taskDirPath.get()
                directory = (r'%s' %directory)
                if os.path.isdir(directory) == False:
                    self.lblmsg2.config(text = "Please, check the path.", fg = 'Red')
                    return             
                y = 1
                z = 7
                   
                for path, subdirs, files in os.walk(directory):
                    for name in files:
                        if name.endswith('.xml'):
                            pathxml = PurePath(path,name)
                            tree = ET.parse(r'%s' %pathxml)
                            root = tree.getroot()
                            for block in root.findall('block'):
                                blockIdxml = block.get('id')
                                if blockIdxml == blockID:
                                    lenght = len(PurePath(path,name).parts)
                                    while (y<lenght):
                                        pathName = PurePath(path,name).parts[y]
                                        if pathName.endswith('.xml'):
                                            vn = len(pathName)-4 
                                            newName = pathName[0:vn]
                                            worksheet.write(x,z,newName)                                              
                                        else:
                                            worksheet.write(x,z,pathName)
                                        y+=1
                                        z+=1                                                 
                x=x+1  
            worksheet.add_table('A1:K%i' %x, {'header_row':True,'banded_rows': True, 'columns':[{'header':'TestCase'},
                                                                                                                     {'header': 'MeasureID'},
                                                                                                                     {'header': 'Interval'},
                                                                                                                     {'header': 'MaxRunTime'},
                                                                                                                     {'header': 'Resource[HDMI Input]'},
                                                                                                                     {'header': 'IdStatus[1=Automatic; 0=Not Automatic]'},
                                                                                                                     {'header': 'BlockID'},
                                                                                                                     {'header': 'Directory'},
                                                                                                                     {'header': 'sub'},
                                                                                                                     {'header': 'sub2'},
                                                                                                                     {'header': 'sub3'},
                                                                                                                     ]})     
            worksheet.data_validation('F1:F%i' %x, {'validate':'list',
                                                    'source':['0','1'],})
            worksheet.data_validation('E1:E%i' %x, {'validate':'list',
                                                    'source':['desktop',
                                                              'witbe-4hdmiv1-0/vid/hdmi0',
                                                              'witbe-4hdmiv1-0/vid/hdmi1',
                                                              'witbe-4hdmiv1-0/vid/hdmi2',
                                                              'witbe-4hdmiv1-0/vid/hdmi3',
                                                              'witbe-4hdmiv1-1/vid/hdmi0',
                                                              'witbe-4hdmiv1-1/vid/hdmi1',
                                                              'witbe-4hdmiv1-1/vid/hdmi2',
                                                              'witbe-4hdmiv1-1/vid/hdmi3']})
            workbook.close()
            self.lblmsg2.config(text = "File successfully created!", fg = 'Black')
        except PermissionError:
            self.lblmsg2.config(text = "Please, close the excel file: %s.xlsx" %fName , fg = 'Red')
            return
        except :
            self.lblmsg2.config(text = "There's no .xml file in folder: %s" %folder, fg = 'Red')
            return
         
    def updateFile(self):
        #Erase label1
        self.lblmsg.config(text = '')
        #Erase label2
        self.lblmsg2.config(text = '')
        
        self.lblmsg.config(text = "Updating, please wait.", fg = 'Black')
        self.lblmsg.update_idletasks()     
        try:
            excel = self.filePath.get()
            folder = self.dirPath.get()
            workbook = xlrd.open_workbook(excel)
            sheet = workbook.sheet_by_index(0)
            rownb = sheet.nrows
            x=1
            while(x < rownb):
                def verVariables():
                    testCase = str(sheet.cell_value(x,0))
                    lenght = len(testCase)
                    y = lenght - 4
                    testCasexml = '%s.xml' % testCase[0:y]
                    # Get values in excel
                    measure_id_MeasureID = int(sheet.cell_value(x,1))
                    measure_id_MeasureID = str(measure_id_MeasureID)
                    interval_Interval = int(sheet.cell_value(x,2))
                    interval_Interval = str(interval_Interval)
                    maxruntime_MaxRunTime = int(sheet.cell_value(x,3))
                    maxruntime_MaxRunTime = str(maxruntime_MaxRunTime)
                    # This variable is a string
                    resource_Resource = sheet.cell_value(x,4)

                    automatic_IDstatus = int(sheet.cell_value(x,5))
                    automatic_IDstatus = str(automatic_IDstatus)

                    block_BlockID = sheet.cell_value(x,6)

                    def defValues(node):
                        # Check if block ID matches
                        for block in node.iter('block'):
                            blID = block.text
                        if block_BlockID == blID:
                            # Compare values from '.xml' with '.xls'
                            for measure_id in node.iter('measure_id'):
                                acc = measure_id.text
                            if acc !=  measure_id_MeasureID:
                                for measure_id in node.iter('measure_id'):
                                    measure_id.text = measure_id_MeasureID
                                    acc = measure_id_MeasureID
                                tree.write(fileName)
                                for measure_id in node.iter('measure_id'):
                                    print ("Measure ID: %s" %(measure_id.text))
                            for interval in node.iter('interval'):
                                acc = interval.text
                            if acc !=  interval_Interval:
                                for interval in node.iter('interval'):
                                    interval.text = interval_Interval
                                    acc = interval_Interval
                                tree.write(fileName)
                                for interval in node.iter('interval'):
                                    print ("Interval: %s" %(interval.text))
                            for maxruntime in node.iter('maxruntime'):
                                acc = maxruntime.text
                            if acc !=  maxruntime_MaxRunTime:
                                for maxruntime in node.iter('maxruntime'):
                                    maxruntime.text = maxruntime_MaxRunTime
                                    acc = maxruntime_MaxRunTime
                                tree.write(fileName)
                                for maxruntime in node.iter('maxruntime'):
                                    print ("MaxRunTime: %s" %(maxruntime.text))
                            for resource in node.iter('resource'):
                                acc = resource.text
                            if acc !=  resource_Resource:
                                for resource in node.iter('resource'):
                                    resource.text = resource_Resource
                                    acc = resource_Resource
                                tree.write(fileName)
                                for resource in node.iter('resource'):
                                    print ("Resource: %s" %(resource.text))
                            for automatic in node.iter('automatic'):
                                acc = automatic.text
                            if acc !=  automatic_IDstatus:
                                for automatic in node.iter('automatic'):
                                    automatic.text = automatic_IDstatus
                                    acc = automatic_IDstatus
                                tree.write(fileName)
                                for automatic in node.iter('automatic'):
                                    print ("IDstatus: %s" %(automatic.text))
                    sharedTaskFolderPath = folder
                    fileName ="%s\%s\%s" % (sharedTaskFolderPath, testCase, testCasexml)
                    tree = ET.parse(fileName)
                    root = tree.getroot()
                    defValues(root)
                verVariables()
                x=x+1
                if x == rownb:
                    self.lblmsg.config(text = "Successfully updated!", fg = 'Black')
        except:
            self.lblmsg.config(text = "Please, check the path", fg = 'Red')
root = Tk()
Application(root)
root.mainloop()