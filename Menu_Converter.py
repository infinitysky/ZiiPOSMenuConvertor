import tkinter as tk
import tkinter.font as tkFont
from tkinter import messagebox
from tkinter.filedialog import askopenfilename
import sys
import os
import pyodbc 
import pandas as pd
import numpy as np
import xlsxwriter
import xlrd
import wget
import ssl
from tqdm import tqdm


ssl._create_default_https_context = ssl._create_unverified_context


def closesystem():
    sys.exit()

def convertToDDAExcel(SourceData,DDATemplate):
    DDAExcelDataFrame = pd.DataFrame()

    TemplateData=DDATemplate.iloc[0]
    TemplateData["WholesalePrice1(Inc GST)"] =0
    TemplateData["WholesalePrice2(Inc GST)"] =0

    z=0
    for z in tqdm(range(len(SourceData))):
    #for z in range(len(SourceData)):
        
        #print(z ," / ",len(SourceData) ," (2)" )
        if pd.notna(SourceData.iloc[z]["StockID"]):
            TemplateData=DDATemplate.iloc[0]
            TemplateData["ProductCode(15)"] = SourceData.iloc[z]["StockID"]
            TemplateData["Description1(100)"] = SourceData.iloc[z]["Description1"]
            TemplateData["Description2(100)"] = SourceData.iloc[z]["Description2"]
            
            TemplateData["Category(25)"] = SourceData.iloc[z]["DepartmentName"]

            TemplateData["SalesPrice1(Inc GST)"] = SourceData.iloc[z]["Price"]
            #TemplateData["WholesalePrice1(Inc GST)"] = SourceData.iloc[z]["F_WPrice"]
            
            TemplateData["LastOrderPrice(Ex GST)"] = SourceData.iloc[z]["ItemCost"]

            # BC1 = str(SourceData.iloc[z]["barcode1"])
            # BC2 = str(SourceData.iloc[z]["barcode2"])
            # BC3 = str(SourceData.iloc[z]["barcode3"])
            # BC4 = str(SourceData.iloc[z]["barcode4"])
            # BC5 = str(SourceData.iloc[z]["barcode5"])
            # BC6 = str(SourceData.iloc[z]["barcode6"])
            
            TemplateData["Barcode1(30)"] = str(SourceData.iloc[z]["barcode1"]).replace(" ", "")
            TemplateData["Barcode2(30)"] = str(SourceData.iloc[z]["barcode2"]).replace(" ", "")
            TemplateData["Barcode3(30)"] = str(SourceData.iloc[z]["barcode3"]).replace(" ", "")
            TemplateData["Barcode4(30)"] = str(SourceData.iloc[z]["barcode4"]).replace(" ", "")
            TemplateData["Barcode5(30)"] = str(SourceData.iloc[z]["barcode5"]).replace(" ", "")
            TemplateData["Barcode6(30)"] = str(SourceData.iloc[z]["barcode6"]).replace(" ", "")
            

            TemplateData["GSTRate"] = SourceData.iloc[z]["GSTRate"]
            TemplateData["Measurement (Pack)"] = SourceData.iloc[z]["PackSize"]
            
        
            if SourceData.iloc[z]["Scale"] == "Y":
                TemplateData["Scaleable"] = 1
            else:
                TemplateData["Scaleable"] = 0    

        
      

            TemplateData = TemplateData.to_frame()
            TemplateData = TemplateData.transpose()
                

            DDAExcelDataFrame = pd.concat([DDAExcelDataFrame, TemplateData],ignore_index=True)

  
    
    return DDAExcelDataFrame




         


def processProductWithBarCode(connect_string):
    PassSQLServerConnection = pyodbc.connect(connect_string)

  

    
    
    productListQuery = "select ItemDetails.*,barcode1,barcode2,barcode3,barcode4,barcode5,barcode6,y.Description as DepartmentName, NormalPrice.Price, CostCompare.ItemCost from ItemDetails left join (Select StockID,  MAX(CASE when a.rowNum=1 THEN Barcode else '' end) barcode1,  MAX(CASE when a.rowNum=2 THEN Barcode else '' end) barcode2,  MAX(CASE when a.rowNum=3 THEN Barcode else '' end) barcode3,  MAX(CASE when a.rowNum=4 THEN Barcode else '' end) barcode4,  MAX(CASE when a.rowNum=5 THEN Barcode else '' end) barcode5,  MAX(CASE when a.rowNum=6 THEN Barcode else '' end) barcode6   from( select *,ROW_NUMBER( ) OVER ( PARTITION BY StockId ORDER BY Barcode  ) AS rowNum from Barcode )a  GROUP BY StockID) x on ItemDetails.StockID =x.StockID  left join (Select * from   Department) y on ItemDetails.Department =y.DepartmentID left join (select * from NormalPrice) NormalPrice on NormalPrice.StockID=ItemDetails.StockID left join (select * from CostCompare)CostCompare on CostCompare.StockID = ItemDetails.StockID"


   

    productList = pd.read_sql_query(productListQuery, PassSQLServerConnection)
  

    result=productList
    

    print("total rows: ")
    print(len(result))
    
    directory='C:\\Ziitech'
    if not os.path.exists(directory):
        os.makedirs(directory)
        
    Export_file="C:\\Ziitech\\export_data.xlsx"
    print("Export to new excel")
    result.to_excel(Export_file, index = True, header=True,engine='xlsxwriter')
    print("Stage 1 Process completed")

    print("Stage 2 Start, converting Data to DDA Formate.....")
    DDAExcel = pd.DataFrame()
    
    DDADownloadTemplate_file="C:\\Ziitech\\ItemImportFormat.xls"
    if not os.path.exists(DDADownloadTemplate_file):
           
        try:
            downloadURL="https://download.ziicloud.com/programs/ziiposclassic/ItemImportFormat.xls"
                
            wget.download(downloadURL, DDADownloadTemplate_file)
        except wget.Error as ex:
            print("Download Files error")
        

    DDADataTemplate = pd.read_excel(DDADownloadTemplate_file, index_col=None,dtype = str)
     #DDAExcel=DDADataTemplete.astype(str)
    DDAExcel =DDADataTemplate
    
    DDAExcelFinal = convertToDDAExcel(result,DDAExcel)
    
    FinalDDA_file="C:\\Ziitech\\OutPut.xls"
    DDAExcelFinal.to_excel(FinalDDA_file, index = False, header=True,engine='xlsxwriter')
    messagebox.showinfo(title="Process Completed",message="Data Process Completed, Please check C:\\Ziitech Folder")
    
   


def processMenu(ExcelSouce):
    directory='C:\\Ziitech'
    if not os.path.exists(directory):
        os.makedirs(directory)
    fullSizeMenuTemplate_file="C:\\Ziitech\\ZiiPOS_MenuTemplate.xls"
    if not os.path.exists(fullSizeMenuTemplate_file):
           
        try:
            downloadURL="https://download.ziicloud.com/other/ZiiPOS_MenuTemplate.xlsx"
                
            wget.download(downloadURL, fullSizeMenuTemplate_file)
        except wget.Error as ex:
            print("Download Files error")
            messagebox.showerror(title="Error", message="Cannot download template file, please check your network !!")
            
            
    currentMenu = pd.read_excel(ExcelSouce, index_col=None,dtype = str)
    print(currentMenu)
            
    DDADataTemplate = pd.read_excel(fullSizeMenuTemplate_file, index_col=None,dtype = str)
    print(DDADataTemplate)
    #DDAExcel=DDADataTemplete.astype(str)



def inforProcess(ExcelSouce):
     
    if ExcelSouce=="":
        messagebox.showerror(title="Error", message="Please Select your Menu !!")
        
    else:
        processMenu(ExcelSouce)

    






class App:
    def __init__(self, root):
        #setting title
        root.title("Menu Converter V1.1")
        #setting window size
        width=600
        height=500
        screenwidth = root.winfo_screenwidth()
        screenheight = root.winfo_screenheight()
        alignstr = '%dx%d+%d+%d' % (width, height, (screenwidth - width) / 2, (screenheight - height) / 2)
        root.geometry(alignstr)
        root.resizable(width=False, height=False)

        GLabel_DB_Source=tk.Label(root)
        ft = tkFont.Font(family='Times',size=10)
        GLabel_DB_Source["font"] = ft
        GLabel_DB_Source["fg"] = "#333333"
        GLabel_DB_Source["justify"] = "left"
        GLabel_DB_Source["text"] = "Select Menu File"
        GLabel_DB_Source.place(x=50,y=90,width=100,height=30)

        MenuSource_Box=tk.Entry(root)
        MenuSource_Box["borderwidth"] = "1px"
        ft = tkFont.Font(family='Times',size=10)
        MenuSource_Box["font"] = ft
        MenuSource_Box["fg"] = "#333333"
        MenuSource_Box["justify"] = "left"

        
        MenuSource_Box.place(x=190,y=90,width=200,height=30)



        








         #-----------------Functions---------------------------------
        def getMenuSource():
            result=MenuSource_Box.get()
            return result
           
        
    
      
        
        def StartConversionProcess():
      
            MenuSource=getMenuSource()
            # username=getDBUsername()
            # password=getDBPassword()
            
            inforProcess(MenuSource)
          
            

        def getExcelSource():
            file=askopenfilename()
            
            MenuSource_Box.insert(0,file)
            print(file)
          
            




            
            


            
            
        
        
        

            















        



            
#--------------Button Actions-------------------------
        Star_Button=tk.Button(root)
        Star_Button["bg"] = "#f0f0f0"
        ft = tkFont.Font(family='Times',size=10)
        Star_Button["font"] = ft
        Star_Button["fg"] = "#000000"
        Star_Button["justify"] = "center"
        Star_Button["text"] = "Start"
        Star_Button.place(x=70,y=390,width=90,height=45)
        Star_Button["command"] = StartConversionProcess

        MenuSelectButton=tk.Button(root)
        MenuSelectButton["bg"] = "#f0f0f0"
        ft = tkFont.Font(family='Times',size=10)
        MenuSelectButton["font"] = ft
        MenuSelectButton["fg"] = "#000000"
        MenuSelectButton["justify"] = "center"
        MenuSelectButton["text"] = "Select File"
        #MenuSelectButton.place(x=250,y=390,width=90,height=45)
        MenuSelectButton.place(x=400,y=90,width=80,height=30)
        MenuSelectButton["command"] = getExcelSource

        Close_Button=tk.Button(root)
        Close_Button["bg"] = "#f0f0f0"
        ft = tkFont.Font(family='Times',size=10)
        Close_Button["font"] = ft
        Close_Button["fg"] = "#000000"
        Close_Button["justify"] = "center"
        Close_Button["text"] = "Close"
        Close_Button.place(x=420,y=390,width=90,height=45)
        Close_Button["command"] = closesystem
       
        




#----------------Not in use--------------------------------
    def Star_Button_command(self):
        print("Star_Button_command")
    def MenuSelectButton_command(self):
        print("command")
    def Close_Button_command(self):
        print("Exit")
        exit()

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
