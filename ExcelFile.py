# -*- coding: utf-8 -*-
"""
Created to ease the process of processing Excel files.

@author: Carlos
"""

import os
import xlrd


class ExcelFolder:
    # Class to represent a bunch of excel files in a folder. Those files
    # will be processed as desire using also the Excel File class


    # Method: Constructor
    # Specify the path to a folder where the Excel files are. 
    # By default is the path where the script is
    
    def __init__(self, folderPath=os.path.abspath(os.path.dirname(__file__))):     
        
        if type(folderPath) == str and os.path.exists(folderPath): # Folder must be str and exists
            self.folderPath = folderPath
            self.xlsFiles, self.badFiles = self.GetXlsFilesFromFolder()
            
        else: 
            raise Exception("Path value must be valid")
    
    
    # Method: GetFilesFromFolder
    # Get all the XLS files in the path
    def GetXlsFilesFromFolder(self): 
        
        folder = os.listdir(self.folderPath)
        pathXlsFiles=[]
        pathXlsBadFiles=[]

        for file in folder:        

            if file.split('.')[1]== 'xls': # If the file has the XLS extension mark it as found
                try: 
                    pathXlsFiles.append(XlsFile(f"{self.folderPath}/{file}"))
                
                except Exception: 
                    print(f"[i] - Skipping file: {file}")
                    pathXlsBadFiles.append(f"{self.folderPath}/{file}")
                    
        return list(pathXlsFiles), list(pathXlsBadFiles)
      
    
    # Method: PrintFiles
    # Print all the files which we are working with
    def PrintGoodFiles(self): 
        for file in self.xlsFiles:
            print(file)
    
    # Method: PrintBadFiles
    # Print all the files with some problem so they can't be processed
    def PrintBadFiles(self): 
        for file in self.badFiles:
            print(f"[X] Bad file: {file}")


class XlsFile:
    # Class to represent a single XLS file with its data


    # Method: Constructor
    # Specify the XLS file path to work with and the sheet number (starting in 0) 
    # By default the excel file will work with the first sheet.
    def __init__(self, filePath, sheetNumber=0):
        
        if os.path.isfile(filePath): 
            self.fileName = filePath
            
            try: 
                self.excelWorkbook = xlrd.open_workbook(self.fileName)
                self.workingSheet = self.excelWorkbook.sheet_by_index(sheetNumber)
                
            except xlrd.biffh.XLRDError: 
                print(f"[X] - ERROR file not supported: {self.fileName}")
                raise xlrd.biffh.XLRDError
                
        else: 
            raise Exception("File must exists and has to be provided to the constructor")
    
    # Method: Printer
    # Specify the way that an object of this class is printed whith print clausule
    def __str__(self):
        return f"Filename - {self.fileName}"


    # Method: Instancer?
    # Change how to display info when using the type() clausule
    def __repr__(self):
        return f"XlsFile - {self.fileName}"
    
    
    
    def GetNumberOfRows(self): 
        pass
        
    def GetNumberOfColumnss(self): 
        pass
    
    # Method: changeWorkSheet
    # Change the current worksheet to work with
    def changeWorkingSheet(self, sheetNumber):
        
        if type(sheetNumber) == int: 
            self.workingSheet = self.excelWorkbook.sheet_by_index(sheetNumber)
            
        else: 
            raise Exception("To change the working sheet you must provide a valid one (starting in 0)")
    

##################### TESTING PART #####################

test = ExcelFolder()
print('-' * 20)
test.PrintGoodFiles()
print('-' * 20)
test.PrintBadFiles()













