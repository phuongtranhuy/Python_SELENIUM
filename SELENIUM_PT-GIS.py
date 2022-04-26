# -*- coding: utf-8 -*-
"""
Created on Wed Jun  9 14:31:33 2021

@author: TNP2HC
"""

# -*- coding: utf-8 -*-

import BASE as base
#import os
#import subprocess
#import winreg
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium import webdriver
from selenium.webdriver.support.ui import Select
#from selenium.webdriver.ie.options import Options
import time
#import win32com.client #to read and write excel 
#from win32com.client import Dispatch
import sys

global description_xpath , otherinfor_xpath, drw_gis_xpath, first_cell_xpath, first_partnumber_xpath
global driver, r, ws, partnumber_pep, error_found, list_QTY
                                  
description_xpath = '/html/body/form/table[2]/tbody/tr[2]/td[2]'
bcode_xpath  = '/html/body/form/table[2]/tbody/tr[2]/td[3]'
BOM_plant_xpath  = '/html/body/form/table[2]/tbody/tr[2]/td[4]'
otherinfor_xpath = '/html/body/form/table[3]/tbody/tr[2]/td[4]/font'
otherinfor_xpath = '/html/body/form/table[3]/tbody/tr[2]/td[4]'
drw_gis_xpath = 'html/body/form/table[2]/tbody/tr[2]/TD[5]'
first_cell_xpath = '/html/body/center/table/tbody/tr/td/a'
first_partnumber_xpath = "html/body/table/tbody/tr[1]/TD[2]/a"


wb = base.Open_EXCEL_File(r'U:\PEP FILE')
wb.Worksheets("Sheet1").Visible = True
ws = wb.Sheets('Sheet1')


def is_visible_ID(driver_,locator,timeout = 7): 
    try: 
        WebDriverWait(driver_, timeout).until(EC.visibility_of_element_located((By.ID, locator))) 
        return True
    except TimeoutException: 
        return False
    
   


class WebTable:
    def __init__(self, webTable):
        self.table = webTable        
    #Get number of Rows : 
    def get_row_count(self):
        num_of_row = len(self.table.find_elements_by_tag_name("tr")) - 1 #get rid of header
        #num_of_row = len(driver.find_element_by_xpath("/html/body/table/tbody/tr"))
        return num_of_row    #Xpath must be changed when customizing your requirement
    
    def get_BOM_count(self):
        num_of_plant = len(self.table.find_elements_by_tag_name("a"))  
        #num_of_row = len(driver.find_element_by_xpath("/html/body/table/tbody/tr"))
        return num_of_plant    #Xpath must be changed when customizing your requirement
    
    def column_data(self, column_number):
       	col = self.table.find_elements_by_xpath("//tr/td["+str(column_number)+"]")
       	rData = []
        #print(len(col)-1)   
       	for ind, webElement in enumerate(col):
            rData.append(webElement.text)
            if ind == len(col)-2: break       
       	return rData
    
    def column_data_other(self, column_number):
       	col = self.table.find_elements_by_xpath("tbody/tr/td["+str(column_number)+"]")
       	rData = []
        #print(len(col)-1)   
       	for ind, webElement in enumerate(col):
            rData.append(webElement.text)
            if ind == len(col)-2: break       
       	return rData   
 

def is_DRW_valid(drws): #MULTIPLEDOK
    drw_valid = False
    if(drws != 'MULTIPLEDOK' and drws != '9.999.999.999' and drws != 'notdefined!' and drws != 'not defined!' and drws != '' and drws != None):
        drw_valid = True
    return drw_valid

def login_GIS(driver):

    #driver.maximize_window()
    
    
    #Access to PT-GIS web-based
    driver.get("URL....")
    
    # get to sign in textbox
    if not is_visible_ID(driver, "loginname"): raise RuntimeError("Wait too long to load default page :(") 
    username = driver.find_element_by_id("loginname")
    password = driver.find_element_by_id("password")
    
    # enter username, password and submit
    username.send_keys("USERNAME")
    password.send_keys("PASSWORD")
    username.submit()
    if not is_visible_ID(driver, "NavMenu"): raise RuntimeError("Wait too long to load search page :(") 
    


def input_partnumber(driver, partnumber_pep, r):
    try:
        #before approch any element in html, make sure which window and content (defaut page ex.) first, then switch to specific frame
         #driver.switch_to.window(driver.window_handles[0])
        driver.implicitly_wait(2)
        driver.switch_to.default_content()
        driver.switch_to.frame("gbom_stammdaten")
        select = Select(driver.find_element_by_id('param_WERK'))
        select.select_by_visible_text('* all plants')
     
    except:
        #error handling when selenium detect wrong window (openning new window to show DRW (in this case)
         #driver.switch_to.window(driver.window_handles[1])
        driver.implicitly_wait(2)
        driver.switch_to.default_content()
        driver.switch_to.frame("gbom_stammdaten")
        select = Select(driver.find_element_by_id('param_WERK'))
        select.select_by_visible_text('* all plants')    
            
    
    
    print("Part number's checking: ", partnumber_pep,"at Row:",r)
    inputpart =  driver.find_element_by_id("param_snr")
    inputpart.clear()
    
    inputpart.send_keys(partnumber_pep)
    search = driver.find_element_by_id("ButtonSearch")
    driver.execute_script("arguments[0].click();", search)
    
def get_other_infor_GIS(driver,otherinfor_xpath):

    try:
        driver.implicitly_wait(2)
        #description = driver.find_element_by_xpath(description_xpath).text.upper()
        #bcode = driver.find_element_by_xpath(bcode_xpath).text.upper()
        w1 = WebTable(driver.find_element_by_xpath('/html/body/form/table[2]'))
        other_infor = w1.column_data(4) 
        #driver.implicitly_wait(2)
        #otherinfor = driver.find_element_by_xpath(otherinfor_xpath).text.upper()
    except:
        driver.implicitly_wait(2)
        #description = driver.find_element_by_xpath(description_xpath).text.upper()
        #bcode = driver.find_element_by_xpath(bcode_xpath).text.upper()
        w1 = WebTable(driver.find_element_by_xpath('/html/body/form/table[2]'))
        other_infor = w1.column_data(4) 
        #driver.implicitly_wait(2)
        #otherinfor = driver.find_element_by_xpath(otherinfor_xpath).text.upper()
    else:
        driver.implicitly_wait(2)
        #description = driver.find_element_by_xpath(description_xpath).text.upper()
        #bcode = ""
        print("Fail to get other information")
        other_infor = ""
    
    compare_str = "".join([other_infor])
    return compare_str


def Get_DRW_Description(driver,ws,partnumber_pep,r):
    
    driver.get("URL.....")
    driver.implicitly_wait(2)
    driver.switch_to.frame("gbom_stammdaten")
    if not is_visible_ID(driver, "form_gbom"): raise RuntimeError("Wait too long to load search page :(") 
    if driver.title != 'GBOM - Global Bill of Material': driver.window_handles[1]
    is_has_technical = False
    list_drw_gis = []
    list_other_infor = []
    input_partnumber(driver, partnumber_pep,r)
    
    # do something with frame (gbom_stammdaten) first, if fail, switch to frame(gbom_positionen), click to part number and turn it back to frame (gbom_stammdaten)
    try:
        driver.switch_to.default_content()
        driver.switch_to.frame("gbom_positionen")
        driver.implicitly_wait(2)
        partnum_xpath  = driver.find_element_by_xpath(first_cell_xpath)
        driver.implicitly_wait(1)
        driver.execute_script("arguments[0].click();", partnum_xpath)
        if not is_visible_ID(driver, "form_gbom"): raise RuntimeError("Wait too long to go to inside TABLE for checking :(")    
        #print('gbom_positionen')
    except:
        pass
    
    driver.implicitly_wait(2)
    driver.switch_to.default_content()
    driver.switch_to.frame("gbom_stammdaten")  
    
    try: # check if "No data found for these criteria" 
        status = driver.find_element_by_xpath('html/body/b').text
        ws.Cells(r,'H').Value = status
        #is_has_technical = False
    except:
        
        is_has_technical = True  
        driver.switch_to.default_content()
        driver.switch_to.frame("gbom_stammdaten") 
        driver.implicitly_wait(2)   
        #driver.switch_to.frame("gbom_stammdaten")
        w = WebTable(driver.find_element_by_xpath(BOM_plant_xpath))
        num_of_plant = w.get_BOM_count()
        print("There are: ",num_of_plant , "BOM plant existing")
        # do some error handling right here
        #error = False 
        more_than_2_plant = False
        is_in_SET = False
        has_got_DRW = False 
        
        if int(num_of_plant) > 1:
            more_than_2_plant = True
            
        if more_than_2_plant == False:
            num_of_plant = 1
    #else:
        #driver.implicitly_wait(1) 
        
        #error = False
        for i in range(1,num_of_plant+1): # Loop over the plants ******************************
            is_have_right = True
            print("Plant:",i)
            try:
                status = driver.find_element_by_xpath('html/body/form/p/b').text
                is_have_right = False
                #error = True
            except:
                
                is_have_right = True
                
            each_plant_xpath = BOM_plant_xpath + '/a[' + str(i) + ']'             
            driver.implicitly_wait(2)
            
            try:               
                driver.switch_to.default_content()
                driver.switch_to.frame("gbom_stammdaten")
                driver.implicitly_wait(3)
                driver.execute_script("arguments[0].click();", driver.find_element_by_xpath(each_plant_xpath))
            except:
                print("Skip this part due to frame has changed")
                break
                #driver.implicitly_wait(2)
                #driver.execute_script("arguments[0].click();", driver.find_element_by_xpath(each_plant_xpath))
    
                #driver.implicitly_wait(2)
                #break # Loop over the plants ******************************
    
            driver.implicitly_wait(3)
            #plant_name = driver.find_element_by_xpath(each_plant_xpath).text
            #driver.implicitly_wait(1)  
            #print(plant_name)
            
            if is_have_right == True:
                if ws.Cells(r,'T').Value == None:
 
                    driver.implicitly_wait(3)
                    try:
                        driver.switch_to.default_content()
                        driver.switch_to.frame("gbom_stammdaten")
                        otherinfor = driver.find_element_by_xpath(otherinfor_xpath).text.upper()
                        #print(otherinfor)
                        list_other_infor.append(otherinfor)
                        driver.implicitly_wait(1)
                        description = driver.find_element_by_xpath(description_xpath).text.upper()
                        driver.implicitly_wait(1)
                        #bcode = driver.find_element_by_xpath(bcode_xpath).text.upper()
                    except:
                        description = ws.Cells(r,'D').Value.upper()
                        otherinfor = ""
                        #bcode = ""

                    otherinfor = "; ".join(list_other_infor)
                    compare_str = "; ".join([description,otherinfor ])
                    
                    driver.implicitly_wait(3)
                    ws.Cells(r,'Q').Value = compare_str
                    
                    #if ws.Cells(r,'E').Value == "Part of PT Tool" : #and (len(ws.Cells(r,'J').Value) == 10):
                        #ws.Cells(r,'I').Value = bcode
                    
                    print(compare_str)
                else:
                    compare_str = ";"
                
                compare_str_ =  compare_str   
            #try:
                driver.implicitly_wait(2)
                driver.switch_to.default_content()
                driver.switch_to.frame("gbom_positionen")
                try:
                    w = WebTable(driver.find_element_by_xpath('/html/body/table'))
                    num_of_row = w.get_row_count()
                    list_descr = w.column_data(3)
                    list_part_num = w.column_data(2)
                    list_QTY = w.column_data(4)
                except:
                    print("Switching plant due to failure to call WebTable")
                
                #print(list_QTY)
                #print(len(list_QTY))
                driver.implicitly_wait(2)
                if len(list_QTY) >1:                      
                    for ind, element in enumerate(list_QTY):
                        if element == "MUC" or element == "MUL" or str(element) == "001" :
                            pos1 = compare_str_.find(list_descr[ind].upper())
                            pos2 = list_descr[ind].upper().find('DOCUMENT')
                            pos3 = list_descr[ind].upper().find('DRAWING')      
                            if pos1 > -1 or pos2 > -1 or pos3 >-1:
                                driver.implicitly_wait(2)
                                drw_valid = is_DRW_valid(list_part_num[ind].strip()) 
                                if drw_valid == True:
                                    has_got_DRW = True
                                    list_drw_gis.append(list_part_num[ind].strip())
                                    print("LIST of DOC get from QTY, case >1",list_drw_gis)                           
                
                elif len(list_QTY) == 1:
                    element = list_QTY[0]
                    
                    if element == "MUC" or element == "MUL" or str(element) == "001" :
                        #print(element)
                        pos1 = compare_str_.find(list_descr[0].upper())
                        pos2 = list_descr[0].upper().find('DOCUMENT')
                        pos3 = list_descr[0].upper().find('DRAWING')                   
                        if pos1 > -1 or pos2 > -1 or pos3 > -1:
                            print(element)
                            driver.implicitly_wait(2)
                            drw_valid = is_DRW_valid(list_part_num[0].strip()) 
                            if drw_valid == True:
                                #print(list_part_num[0].strip())
                                has_got_DRW = True
                                list_drw_gis.append(list_part_num[0].strip())
                                print("LIST of DOC get from QTY, case =1",list_drw_gis)     
                        
                driver.implicitly_wait(2)
                driver.switch_to.default_content()
                driver.switch_to.frame("gbom_stammdaten")
                driver.implicitly_wait(2)
                drawing_gis = driver.find_element_by_xpath(drw_gis_xpath).text
                print(drawing_gis)
    
                
                if len(list_drw_gis) == 0:
                    
                    if num_of_row > 1:
    
                        if list_QTY[0] != "MUC" and list_QTY[0] != "MUL" and list_QTY[0] != "001" :
                            #print("In the case row > 1")
                            #print("test len of part number [0]", len(list_part_num[0]))
                            #print("test len of part number [1]", len(list_part_num[0]))
                            
                            if not (list_descr[0].upper() == list_descr[1].upper()):                           
                                    
                                for j in range(1,3): #Loop: Find final BOM
                                    driver.implicitly_wait(2)        
                                    pos = compare_str_.find(list_descr[0].upper())
                                            
                                    if pos > -1:
                                        try:
                                            print("pos: ",pos)
                                            driver.switch_to.default_content()
                                            driver.switch_to.frame('gbom_positionen')
                                            driver.implicitly_wait(2)
                                            driver.execute_script("arguments[0].click();", driver.find_element_by_xpath(first_partnumber_xpath))
                                            print("first cell has been clicked, case > 1 row ")
                                            
                                            #driver.implicitly_wait(2)
                                            #w = WebTable(driver.find_element_by_xpath('/html/body/table'))
                                            #num_row_after_clicked = w.get_row_count()
                                            #print("Number of row after clicked",num_row_after_clicked)
                                            
                                            driver.implicitly_wait(5)
                                            driver.switch_to.default_content()
                                            driver.switch_to.frame("gbom_stammdaten")
                                            driver.implicitly_wait(2)
                                            drawing_gis = driver.find_element_by_xpath(drw_gis_xpath).text
                                            print("DRW get from DRW box, case >1 row:",drawing_gis)
                                                                                       
                                            
                                            driver.implicitly_wait(2)
                                            drw_valid = is_DRW_valid(drawing_gis.strip())                 
                                            if drw_valid == True:
                                                list_drw_gis.append(drawing_gis.strip())
                                                print("List of DOC get from DRW box, case > 1 row:",list_drw_gis)
                                                has_got_DRW = True
                                                break #Loop: Find final BOM
                                        except:
                                            break  #Loop: Find final BOM
                                    else:
                                        break #Loop: Find final BOM
                            else:
                                if not (len(list_part_num[0]) < 2 and len(list_part_num[1]) < 2):
                                    is_in_SET = True
                                    driver.implicitly_wait(2)
                                    drw_valid = is_DRW_valid(drawing_gis.strip())                 
                                    if drw_valid == True:
                                        list_drw_gis.append(drawing_gis.strip())
                                        print("List of DOC get from DRW box, case > 1 row:",list_drw_gis)
                                        has_got_DRW = True
                                        break #Loop: Find final BOM
                        
                        else:
                            drw_valid = is_DRW_valid(list_part_num[ind].strip()) 
                            if drw_valid == True:
                                has_got_DRW = True
                                list_drw_gis.append(list_part_num[ind].strip())
                                print("LIST of DOC get from QTY, case = 1 row",list_drw_gis) 
                    
                    
                    elif num_of_row == 1:
                        
                        if list_QTY[0] != "MUC" and list_QTY[0] != "MUL" and list_QTY[0] != "001" and list_QTY[0] != "MUB" :
                            
                            for j in range(1,3): #Loop: Find final BOM
                                pos = compare_str_.find(list_descr[0].upper())
                                #print("Code access right here")
                                if pos > -1: 
                                    try:
                                        print("pos: ",pos)
                                        driver.switch_to.default_content()
                                        driver.switch_to.frame('gbom_positionen')
                                        driver.implicitly_wait(2)
                                        driver.execute_script("arguments[0].click();", driver.find_element_by_xpath(first_partnumber_xpath))
                                        print("first cell has been clicked, case = 1 row")
                                        
                                        #w = WebTable(driver.find_element_by_xpath('/html/body/table'))
                                        #num_row_after_clicked = w.get_row_count()
                                        #print("Number of row after clicked",num_row_after_clicked)
                                        
                                        driver.implicitly_wait(5)
                                        driver.switch_to.default_content()
                                        driver.switch_to.frame("gbom_stammdaten")
                                        driver.implicitly_wait(2)
                                        drawing_gis = driver.find_element_by_xpath(drw_gis_xpath).text
                                        print("DRW get from DRW box:",drawing_gis)
                                        
                                        driver.implicitly_wait(2)
                                        drw_valid = is_DRW_valid(drawing_gis.strip())                 
                                        if drw_valid == True:
                                            list_drw_gis.append(drawing_gis.strip())
                                            print("List of DOC get from DRW box, case = 1 row:",list_drw_gis)
                                            has_got_DRW = True
                                            break #Loop: Find final BOM
                                    except: 
                                        break #Loop: Find final BOM
                                else:
                                    driver.implicitly_wait(2)
                                    drw_valid = is_DRW_valid(drawing_gis.strip())                 
                                    if drw_valid == True:
                                        list_drw_gis.append(drawing_gis.strip())
                                        print("List of DOC get from DRW box, case = 1 row:",list_drw_gis)
                                        has_got_DRW = True
                                    break  #Loop: Find final BOM       
                        else:
                            driver.implicitly_wait(2)
                            drw_valid = is_DRW_valid(drawing_gis.strip())                 
                            if drw_valid == True:
                                list_drw_gis.append(drawing_gis.strip())
                                print("List of DOC get from DRW box, case = 1 row:",list_drw_gis)
                                has_got_DRW = True
                            break #Loop: Find final BOM
                    else:
                        driver.implicitly_wait(2)
                        drw_valid = is_DRW_valid(drawing_gis.strip()) 
                        #print("Is DRW valid: ",drw_valid)
                        if drw_valid == True:
                            list_drw_gis.append(drawing_gis.strip())
                            has_got_DRW = True
                            print("List of DOC get from DRW box, case < 2 rows:",list_drw_gis)
                
                else: 
                    driver.implicitly_wait(2)
                        
                    break # Loop over the plants ******************************
                
            if has_got_DRW == True: break
        
    if is_has_technical == True:            
        if has_got_DRW == True: 
            list_drw_gis = list(dict.fromkeys(list_drw_gis))
            print('List of GIS DRW after Remove Duplicate:',list_drw_gis)
            ws.Cells(r,'H').Value = "; ".join(list_drw_gis)
            
            #if ws.Cells(r,'E').Value == "Part of PT Tool":
                #ws.Cells(r,'H').AddComment("DRW from Where-Used") 
                      
                    
        if len(list_drw_gis)  == 0:
            #if not ws.Cells(r,'E').Value == "Part of PT Tool":
            if is_in_SET == True:
                ws.Cells(r,'H').Value = "Part number refer to a SET of PRODUCT"
            else:
                ws.Cells(r,'H').Value = "Cannot find DRW on GIS" 
            #else:
                #ws.Cells(r,'H').Value = "Cannot find DRW for Where-Used part"     
    
    return #----------------------------*******-------------------------------------------------------------------      


def Get_Where_Used_Infor(driver,ws,part_no,r):
    #part_no = ws.Cells(r,'C').Value.strip()
    description = ws.Cells(r,'D').Value.strip().upper()
    #ws.Cells(r,'A').Value = '.'
    has_WU = False
    
    print("Find Where-Used, number's checking:",part_no," ","at row:", r)
    
    driver.get('URL......')
    snr = driver.find_element_by_id("param_snr")
    snr.send_keys(part_no)
    driver.implicitly_wait(2)
    select = Select(driver.find_element_by_id('param_WERK'))
    select.select_by_visible_text('* all plants')
    driver.implicitly_wait(2)
    driver.execute_script("arguments[0].click();", driver.find_element_by_id("ButtonSingle"))
    driver.implicitly_wait(2)
    try:
        w1 = WebTable(driver.find_element_by_xpath("html/body/table"))
        number_rowcount = w1.get_row_count()
        print("Number of WH-U list:",number_rowcount)
        list_part_WU = w1.column_data_other(5)
        driver.implicitly_wait(2)
        list_desc = w1.column_data_other(6)
        if number_rowcount > 3:                   
            #print(list_part_WU, "case >1")
            #print(list_desc,"case >1")
            for ind, part in enumerate(list_part_WU): #can not do for loop with row = 1??????
                if len(part) > 3:
                    #print(part.strip().replace(" ", ""))
                    check_part = part.replace(" ", "").strip()
                    #print(check_part[0:2])
                    if check_part[0:2] == "36" or check_part[0:2] == "06" or check_part[0:2] == "F1" or check_part[0:2] == "F0":
                        if not list_desc[ind].upper().strip() == description:
                            print(check_part)
                            print(list_desc[ind].upper())
                            #description_WU = str(list_desc[ind].upper())
                            has_WU = True
                            break
                    if  ind == 50: break   
        
        elif number_rowcount < 3:
            print(list_part_WU, "case = 1")
            print(list_desc,"case = 1")
            check_part = list_part_WU[0].replace(" ", "").strip()
            if check_part[0:2] == "36" or check_part[0:2] == "06" or check_part[0:2] == "F1" or check_part[0:2] == "F0":
                if not list_desc[0].upper().strip() == description:
                    print(check_part)
                    print(list_desc[0].upper())
                    #description_WU = str(list_desc[0].upper())
                    has_WU = True
            
        if has_WU == True:
            ws.Cells(r,'O').Value = "'".join(["",check_part])
            #ws.Cells(r,'I').Value = description_WU
        else:
            ws.Cells(r,'O').Value = "No Where-Used Found (LV1)"
            #ws.Cells(r,'I').ClearContents()
            
        
    except Exception as e:
        print(e)
        print("No Where-Used List")
        try:
            error_text = driver.find_element_by_xpath("html/body/b").text.upper()
            ws.Cells(r,'O').Value = "No Where-Used Found"
            #ws.Cells(r,'I').ClearContents()
        except:
            error_text = "No Where-Used Found"
            ws.Cells(r,'O').Value = error_text
            #ws.Cells(r,'I').ClearContents()
    
    return 
#main() 
def main(): 
    #global list_QTY
    #Open webdriver first to control IE automation, mandatory
    #driver = webdriver.Ie(executable_path = r"C:\Users\TNP2HC\Desktop\PT Power Tool\PT-GIS tool\IEDriverServer_32bit.exe", options=opts)
    print("Opening Google Chrome.......")
    #driver = webdriver.Chrome(executable_path = r"C:\Users\TNP2HC\Desktop\Python\Python Web Scrapping\chromedriver.exe")
    driver = webdriver.Chrome(executable_path = r"C:\Users\TNP2HC\Desktop\Python\Python Web Scrapping\chromedriver.exe")
    login_GIS(driver)
    lastrow = ws.Cells(ws.Cells.Rows.Count, "H").End(-4162).Row
    r = lastrow + 1   # start working with the first non value row in column "H"

 
    while ws.Cells(r,'C').Value:
    #-------------------------------------------------------*******---------------------------------------------------------------------------      
        partnumber_pep = ws.Cells(r,'C').Value
        print("Type of products:", ws.Cells(r,'F').Value)
        #Get_DRW_Description(driver,ws,partnumber_pep,r)   #Find information on GIS for all type of product
        
        if ws.Cells(r,'E').Value == "Part of PT Tool":            
            
            
            driver.implicitly_wait(2)
            #print(len(ws.Cells(2198,'K').Value))
            if len(ws.Cells(r,'K').Value) == 10:
                partnumber_WU = ws.Cells(r,'O').Value
                Get_DRW_Description(driver,ws,partnumber_WU,r)
            else:
                Get_Where_Used_Infor(driver,ws,partnumber_pep,r)
                if len(ws.Cells(r,'K').Value) == 10:
                    partnumber_WU = ws.Cells(r,'O').Value
                    Get_DRW_Description(driver,ws,partnumber_WU,r)
        
        else:
            Get_DRW_Description(driver,ws,partnumber_pep,r)
        
        r += 1
        print("----------------------------------------")
        print("\n")
            
    print("Get information in PT-GIS Done !!!!")   
    sys.exit() 
    
    
while True:
    #global list_QTY   
    
    try:         
        main()
        
    except Exception as e:        
        print(e)
        #print('Value when error is :' + str(list_QTY))
        #os.system("TASKKILL /f /IM CHROME. EXE")
        #ws.Cells(r,'F').Value = None
        print ('Restarting!')
        time.sleep(5)
        k = ws.Cells(ws.Cells.Rows.Count, "H").End(-4162).Row + 1
        ws.Cells(k,'H').Value = "Error happen"
        lastrow_colC = ws.Cells(ws.Cells.Rows.Count, "C").End(-4162).Row
        lastrow_colH = ws.Cells(ws.Cells.Rows.Count, "H").End(-4162).Row 
        print(lastrow_colH,lastrow_colC)
        
        continue
        
sys.exit()
print("DONE!!!")
    



