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
directory='C:\\Ziitech'
FinalMenulFile=directory+'\\export_FullMenu.xlsx'
fullSizeMenuTemplate_file=directory+"\\ZiiPOS_MenuTemplate.xlsx"
Export_file=directory+"\\export_data.xlsx"
downloadURL="https://download.ziicloud.com/other/ZiiPOS_MenuTemplate.xlsx"

def closesystem():
    sys.exit()

    
   
def processMenuGroup(ExcelSouce,template):
    MenuGroupDataFrame = pd.DataFrame()
    
    TemplateData=template.iloc[0]
    menugroups=ExcelSouce.loc[:,"MenuGroup"]
    MenuGroupList = menugroups.drop_duplicates()
    print(MenuGroupList)
    #print(MenuGroupList[0]["MenuGroup"])
    
    z=0
    for z in tqdm(range(len(MenuGroupList))):
        for z in range(len(MenuGroupList)):
            
            TemplateData["Description"]=MenuGroupList[z]
            TemplateData["CultureDescription"]=MenuGroupList[z]
            TemplateData["Code"]=z
            TemplateData["OrderIndex"]=z
        
            TemplateData = TemplateData.to_frame()
            TemplateData = TemplateData.transpose()
                

            MenuGroupDataFrame = pd.concat([MenuGroupDataFrame, TemplateData],ignore_index=True)
    return MenuGroupDataFrame
  
    
    
def processCategory(ExcelSouce,template):
    CategoryDataFrame = pd.DataFrame()
    
    TemplateData=template.iloc[0]
    CategoryList=ExcelSouce.loc[:,"Category"]
    CategoryList = CategoryList.drop_duplicates()
    print(CategoryList)
    #print(MenuGroupList[0]["MenuGroup"])
    
    z=0
    for z in tqdm(range(len(CategoryList))):
        for z in range(len(CategoryList)):
            
            
            TemplateData["Category"]=CategoryList[z]
            TemplateData["CultureCategory"]=CategoryList[z]
            TemplateData["MenuGroupCode"]="0"
            zz=z+1
            code = ("%02d" % zz)
            TemplateData["Code"]=code
            TemplateData["OrderIndex"]=z
            TemplateData["Enable"]="TRUE"
        
            TemplateData = TemplateData.to_frame()
            TemplateData = TemplateData.transpose()
            CategoryDataFrame = pd.concat([CategoryDataFrame, TemplateData],ignore_index=True)
    return CategoryDataFrame
    
 
def processItem(ExcelSouce,template):
    
    print("============= Process Item =====================")
    ItemDataFrame = pd.DataFrame()
    
    
    TemplateData["ItemCode"]=""
    TemplateData["XeroAccountCode"]=""
    TemplateData["XeroAccountId"]=""
    TemplateData["BarCode"]=""
    TemplateData["BarCode1"]=""
    TemplateData["BarCode2"]=""
    TemplateData["BarCode3"]=""
    TemplateData["Instruction"]=""
    TemplateData["Multiple"]=""
    TemplateData["Price1"]=""
    TemplateData["Price2"]=""
    TemplateData["Price3"]=""
    TemplateData["PackagePrice"]=""
    TemplateData["PackagePrice1"]=""
    TemplateData["PackagePrice2"]=""
    TemplateData["PackagePrice3"]=""
    TemplateData["SubDescription"]=""
    TemplateData["SubDescription1"]=""
    TemplateData["SubDescription2"]=""
    TemplateData["SubDescription3"]=""
    TemplateData["Description1"]=""
    TemplateData["Description2"]=""
    TemplateData["Price"]=""
    TemplateData["TaxRate"]=""
    TemplateData["Category"]=""
    TemplateData["Active"]=""
    TemplateData["PrinterPort"]=""
    TemplateData["AllowDiscount"]=""
    TemplateData["JobListColor"]=""
    TemplateData["OpenPrice"]=""
    TemplateData["PrinterPort1"]=""
    TemplateData["PrinterPort2"]=""
    TemplateData["HappyHourPrice1"]=""
    TemplateData["HappyHourPrice2"]=""
    TemplateData["HappyHourPrice3"]=""
    TemplateData["HappyHourPrice4"]=""
    TemplateData["DefaultQty"]=""
    TemplateData["SubDescriptionSwap"]=""
    TemplateData["MainPosition"]=""
    TemplateData["POSPosition"]=""
    TemplateData["KitchenScreenFontColor"]=""
    TemplateData["PrinterPort3"]=""
    TemplateData["ItemGroup"]=""
    TemplateData["NoteGroupCode"]=""
    TemplateData["OnlyShowOnSubMenu"]=""
    TemplateData["SubCategory"]=""
    TemplateData["ButtonColor1"]=""
    TemplateData["PhoneOrderPosition"]=""
    TemplateData["AutoPopSpellInstructionKeyboard"]=""
    TemplateData["KitchenScreen1"]=""
    TemplateData["KitchenScreen2"]=""
    TemplateData["KitchenScreen3"]=""
    TemplateData["KitchenScreen4"]=""
    TemplateData["Scalable"]=""
    TemplateData["WeekendPrice"]=""
    TemplateData["WeekendPrice1"]=""
    TemplateData["WeekendPrice2"]=""
    TemplateData["WeekendPrice3"]=""
    TemplateData["Recommended"]=""
    TemplateData["PricePicture1"]=""
    TemplateData["PricePictureCloudAddr1"]=""
    TemplateData["PricePicture2"]=""
    TemplateData["PricePictureCloudAddr2"]=""
    TemplateData["PricePicture3"]=""
    TemplateData["PricePictureCloudAddr3"]=""
    TemplateData["PricePicture4"]=""
    TemplateData["PricePictureCloudAddr4"]=""
    TemplateData["PicturePath"]=""
    TemplateData["PictureCloudAddr"]=""
    TemplateData["OnlinePicturePath"]=""
    TemplateData["OnlinePictureCloudAddr"]=""
    TemplateData["OnlineDisplayName1"]=""
    TemplateData["OnlineDisplayName2"]=""
    TemplateData["PricePictureUrl1"]=""
    TemplateData["PricePictureUrl2"]=""
    TemplateData["PricePictureUrl3"]=""
    TemplateData["PricePictureUrl4"]=""
    TemplateData["PictureUrl"]=""
    TemplateData["OnlinePictureUrl"]=""
    TemplateData["Description3"]=""
    TemplateData["Description4"]=""
    TemplateData["ItemDescription1"]=""
    TemplateData["ItemDescription2"]=""
    TemplateData["ItemDescription3"]=""
    TemplateData["ItemDescription4"]=""
    TemplateData["TimeChargeItem"]=""
    TemplateData["SoldOut"]=""
    TemplateData["SoldOutUpdateTime"]=""
    TemplateData["PromotionItem"]=""
    TemplateData["CanBeRedeemItem"]=""
    TemplateData["TareWeight"]=""
    TemplateData["Cost"]=""
    TemplateData["Cost1"]=""
    TemplateData["Cost2"]=""
    TemplateData["Cost3"]=""
    TemplateData["QuantityFollowByPeopleCount"]=""
    TemplateData["RedeemPoints"]=""
    TemplateData["OnlineOrderItem"]=""
    TemplateData["OtherChargeItem"]=""
    TemplateData["WeightDivideMeasureAsQty"]=""
    TemplateData["MeasureWeight"]=""
    TemplateData["CategoryList"]=""
    TemplateData["ForeColor"]=""
    TemplateData["BorderColor"]=""
    TemplateData["Ingredients"]=""
    TemplateData["CultureDescription"]=""
    TemplateData["CultureCategory"]=""
    TemplateData["CultureItemDescription"]=""
    TemplateData["MaximumQty"]=""
    TemplateData["TimeConsumingItem"]=""
    TemplateData["OnlineStatus"]=""
    TemplateData["QRCodeStatus"]=""
    TemplateData["OnlinePrice1"]=""
    TemplateData["OnlinePrice2"]=""
    TemplateData["OnlinePrice3"]=""
    TemplateData["OnlinePrice4"]=""
    TemplateData["OrderIndex"]=""
    TemplateData["AllowGift"]=""
    TemplateData["MenuItemCategorySort"]=""
    TemplateData["DoNotAutoEnterSubmenuPage"]=""
    TemplateData["SoldOutSyncFlag"]=""
    TemplateData["BgColor"]=""
    TemplateData["ItemFontColor"]=""
    TemplateData=template.iloc[0]

  
   
  
    #print(ExcelSouce.iloc[0]["ItemCode"])
    #print(MenuGroupList[0]["MenuGroup"])
    
    z=0
    for z in tqdm(range(len(ExcelSouce))):
        for z in range(len(ExcelSouce)):
            
            TemplateData["ItemCode"]=ExcelSouce.iloc[z]["ItemCode"]
            TemplateData["Description1"]=ExcelSouce.iloc[z]["Description1"]
            TemplateData["Description2"]=ExcelSouce.iloc[z]["Description2"]
            TemplateData["Description3"]=ExcelSouce.iloc[z]["Description3"]
            TemplateData["Description4"]=ExcelSouce.iloc[z]["Description4"]
            TemplateData["Price"]=ExcelSouce.iloc[z]["Price"]
            TemplateData["Price1"]=ExcelSouce.iloc[z]["Price1"]
            TemplateData["Price2"]=ExcelSouce.iloc[z]["Price2"]
            TemplateData["Price3"]=ExcelSouce.iloc[z]["Price3"]
            
            TemplateData["CultureDescription"]=ExcelSouce.iloc[z]["Description1"]
            TemplateData["Category"]=ExcelSouce.iloc[z]["Category"]
            if ExcelSouce.iloc[z]["ItemGroup"]=="":
                ExcelSouce.iloc[z]["ItemGroup"]="OTHERS"
            else:
                TemplateData["ItemGroup"]=ExcelSouce.iloc[z]["ItemGroup"]
                
            
            TemplateData["HappyHourPrice1"]=ExcelSouce.iloc[z]["HappyHourPrice1"]
            TemplateData["HappyHourPrice2"]=ExcelSouce.iloc[z]["HappyHourPrice2"]
            TemplateData["HappyHourPrice3"]=ExcelSouce.iloc[z]["HappyHourPrice3"]
            TemplateData["HappyHourPrice4"]=ExcelSouce.iloc[z]["HappyHourPrice4"]
            
            TemplateData["Instruction"]=ExcelSouce.iloc[z]["Instruction"]
            TemplateData["Multiple"]=ExcelSouce.iloc[z]["Multiple"]
            TemplateData["Scalable"]=ExcelSouce.iloc[z]["Scalable"]
            TemplateData["OpenPrice"]=ExcelSouce.iloc[z]["OpenPrice"]
            TemplateData["OnlineStatus"]=ExcelSouce.iloc[z]["OnlineStatus"]
            TemplateData["QRCodeStatus"]=ExcelSouce.iloc[z]["QRCodeStatus"]
            TemplateData["PrinterPort1"]=ExcelSouce.iloc[z]["PrinterPort1"]
            TemplateData["PrinterPort2"]=ExcelSouce.iloc[z]["PrinterPort2"]
            TemplateData["PrinterPort3"]=ExcelSouce.iloc[z]["PrinterPort3"]
            TemplateData["PrinterPort4"]=ExcelSouce.iloc[z]["PrinterPort4"]

            
           
         
        #   ItemCode	
        #   Description1	
        #   Description2	
        #   Description3	
        #   Description4	
        #   MenuGroup	
        #   Category	
        #   ItemGroup	
        #   Price	
        #   Price1	
        #   Price2	
        #   Price3	
        #   HappyHourPrice1	
        #   HappyHourPrice2	
        #   HappyHourPrice3	
        #   HappyHourPrice4	
        #   WeekendPrice	
        #   WeekendPrice1	
        #   WeekendPrice2	
        #   WeekendPrice3	
        #   TaxRate	
        #   SubDescription	
        #   SubDescription1	
        #   SubDescription2	
        #   SubDescription3	
        #   Instruction	
        #   Scalable	
        #   OpenPrice	
        #   Multiple	
        #   OnlineStatus	
        #   QRCodeStatus	
        #   PrinterPort1	
        #   PrinterPort2	
        #   PrinterPort3	
        #   PrinterPort4

           # TemplateData["Enable"]="TRUE"
            print("+================++++++")
            print(TemplateData)
            TemplateData = TemplateData.to_frame()
            TemplateData = TemplateData.transpose()
            ItemDataFrame = pd.concat([ItemDataFrame, TemplateData],ignore_index=True)
    
    print(ItemDataFrame)
    return ItemDataFrame
    
    


def processMenu(ExcelSouce):
  
    
    if not os.path.exists(directory):
        os.makedirs(directory)
    if not os.path.exists(fullSizeMenuTemplate_file):
        try:
            wget.download(downloadURL, fullSizeMenuTemplate_file)
        except wget.Error as ex:
            print("Download Files error")
            messagebox.showerror(title="Error", message="Cannot download template file, please check your network !!")
            
            
    sampleMenu = pd.read_excel(ExcelSouce, index_col=None)
   
    
    fullSizeExcel=pd.ExcelFile(fullSizeMenuTemplate_file)
   
    
    MenuGroupTable = pd.read_excel(fullSizeExcel, 'MenuGroupTable')
    ItemGroupTable = pd.read_excel(fullSizeExcel, 'ItemGroupTable')
    MenuitemTable=pd.read_excel(fullSizeExcel, 'MenuItem')
    
    
    
    Coursetable = pd.read_excel(fullSizeExcel, 'Course')
    CategoryTable = pd.read_excel(fullSizeExcel, 'Category')
    PresetNoteGroupTable = pd.read_excel(fullSizeExcel, 'PresetNoteGroup')
    MenuItemRelationTable = pd.read_excel(fullSizeExcel, 'MenuItemRelation')
    SubMenuLinkHeadTable = pd.read_excel(fullSizeExcel, 'SubMenuLinkHead')
    SubMenuLinkDetailTable = pd.read_excel(fullSizeExcel, 'SubMenuLinkDetail')
    SubItemGroupTable = pd.read_excel(fullSizeExcel, 'SubItemGroup')
    InstructionLinkGroupTable = pd.read_excel(fullSizeExcel, 'InstructionLinkGroup')
    InstructionLinkTable = pd.read_excel(fullSizeExcel, 'InstructionLink')

    
    
    
    #MenuGroupSheet=processMenuGroup(sampleMenu,MenuGroupTable)
    CategoryTable=processCategory(sampleMenu,CategoryTable)
    #MenuitemTable = processItem(sampleMenu,MenuitemTable)



    #-------------------------output excel-----------------------------------------------------------
    writer = pd.ExcelWriter(FinalMenulFile, engine = 'xlsxwriter')
    MenuGroupTable.to_excel(writer, sheet_name = 'MenuGroupTable',index = False, header=True)


    CategoryTable.to_excel(writer, sheet_name = 'Category',index = False, header=True)
    MenuitemTable.to_excel(writer, sheet_name = 'Menuitem',index = False, header=True)
    
    Coursetable = Coursetable[0:0]
    Coursetable.to_excel(writer, sheet_name = 'Course',index = False, header=True)
    
    ItemGroupTable.to_excel(writer, sheet_name = 'ItemGroupTable',index = False, header=True)
    
    PresetNoteGroupTable.to_excel(writer, sheet_name = 'PresetNoteGroup',index = False, header=True)
    
    MenuItemRelationTable.to_excel(writer, sheet_name = 'MenuGroupTable',index = False, header=True)
    
    SubMenuLinkHeadTable.to_excel(writer, sheet_name = 'SubMenuLinkHead',index = False, header=True)
    
    SubMenuLinkDetailTable.to_excel(writer, sheet_name = 'SubMenuLinkDetail',index = False, header=True)
    
    SubItemGroupTable.to_excel(writer, sheet_name = 'SubItemGroup',index = False, header=True)
    
    InstructionLinkGroupTable.to_excel(writer, sheet_name = 'InstructionLinkGroup',index = False, header=True)
    
    InstructionLinkTable.to_excel(writer, sheet_name = 'InstructionLink',index = False, header=True)
    

    

    # Course
    # Category
    # PresetNoteGroup
    # MenuItemRelation
    # SubMenuLinkHead
    # SubMenuLinkDetail
    # SubItemGroup
    # InstructionLinkGroup
    # InstructionLink
    
    

    writer.close()
    

   

    
    
    
   
    
    #DDADataTemplate = pd.read_excel(fullSizeMenuTemplate_file, index_col=None,dtype = str)
    
    #DDAExcel=DDADataTemplete.astype(str)
    messagebox.showinfo(title="Process Completed",message="Data Process Completed, Please check C:\\Ziitech Folder")



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
