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
            print(f"\n[X] Bad file: {file}")


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
                self.rowsNumber= self.GetNumberOfRows()
                self.columnsNumber = self.GetNumberOfColumns()
                
            except xlrd.biffh.XLRDError: 
                print(f"[X] - ERROR file not supported: {self.fileName}")
                raise xlrd.biffh.XLRDError
                
        else: 
            raise Exception("File must exists and has to be provided to the constructor")
    
    # Method: Printer
    # Specify the way that an object of this class is printed whith print clausule
    def __str__(self):
        return f"\nFilename - {self.fileName}\n\tColumns: {self.columnsNumber}\n\tRows: {self.rowsNumber}"


    # Method: Instancer?
    # Change how to display info when using the type() clausule
    def __repr__(self):
        return f"XlsFile - {self.fileName}"
    
    
    # Method: GetNumberOfRows
    # Get number of rows from the Excel sheet
    def GetNumberOfRows(self): 
        return self.workingSheet.nrows
    

    # Method: GetNumberOfColumns
    # Get number of columns from the Excel sheet
    def GetNumberOfColumns(self): 
        return self.workingSheet.ncols
    

    # Method: GetColumnData
    # Get all the data from a colum starting in the row startRow until the row "endRow"
    # (by defaultall the values of the column are returned)
    def GetColumnData(self, columNumber, startRow = 0, endRow = None):
        return self.workingSheet.col_values(columNumber, start_rowx=startRow, end_rowx=endRow)
        

    # Method: GetRowData
    # Get all the data from a row starting in the row startRow until the row "endRow"
    # (by defaultall the values of the column are returned)
    def GetRowData(self, rowNumber, startCol = 0, endCol = None):
        return self.workingSheet.row_values(rowNumber, start_colx=startCol, end_colx=endCol)
    

    # Method: GetAllDataByColumns
    # Get all the data from all the columns. This method return an array of arrays.
    # In each position of the array it's a whole column
    def GetAllDataByColumns(self, startColumn=0, stopColumn=None):
        
        columnsData = []
        
        if stopColumn == None: 
            stopColumn = self.GetNumberOfColumns()
            
        for column in range(startColumn, stopColumn):
            columnsData.append(self.GetColumnData(column))
            
        return columnsData
            
            
    # Method: GetAllDataByRows
    # Get all the data from all the rows. This method return an array of arrays.
    # In each position of the array it's a whole row
    def GetAllDataByRows(self, startRow=0, stopRow=None):
        
        rowsData = []
        
        if stopRow == None: 
            stopRow = self.GetNumberOfRows()
        
        for row in range(startRow, stopRow):
            rowsData.append(self.GetRowData(row))
            
        return rowsData
    
    
    # Method: changeWorkSheet
    # Change the current worksheet to work with
    def changeWorkingSheet(self, sheetNumber):
        
        if type(sheetNumber) == int: 
            self.workingSheet = self.excelWorkbook.sheet_by_index(sheetNumber)
            
        else: 
            raise Exception("To change the working sheet you must provide a valid one (starting in 0)")
    


##################### TESTING PART #####################

AllFiles = ExcelFolder("../Stuff")
# test.PrintGoodFiles()
# test.PrintBadFiles()

print(AllFiles.xlsFiles[1])
print(AllFiles.xlsFiles[1].GetRowData(0)) # Headers of the file
print(AllFiles.xlsFiles[1].GetColumnData(0))
print(AllFiles.xlsFiles[1].GetAllDataByRows(0,2))
print(AllFiles.xlsFiles[1].GetAllDataByColumns(0,2))












