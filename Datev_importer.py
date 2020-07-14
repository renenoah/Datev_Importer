
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import xlrd
from datetime import datetime




alle_belege = []  


def excel():
        loc = ("eingabe.xlsx")
        wb = xlrd.open_workbook(loc) 
        sheet = wb.sheet_by_index(0) 
        temp = 1;
        # For row 0 and column 0 
        for i in range(17,sheet.nrows): 
            if sheet.cell_value(i,3) != "" and sheet.cell_value(i,7) == "Bar":
                excel_date = int(sheet.cell_value(i,0))
                dt = datetime.fromordinal(datetime(1900, 1, 1).toordinal() + excel_date - 2)
                
                
                temp = temp +1;

                
                alle_belege.append(beleg(dt,sheet.cell_value(i,4)))
                
                
def printing_list(liste): 
    for i in liste:
        print(i.value)               
 

class beleg:  
    def __init__(self, x, y):  
        self.datum = x  # as an number 
        self.value = int(y)       



excel()


def foo():
    browser = webdriver.Firefox()
    browser.get("https://secure11.datev.de/wopl/FC/FC")
    temp = 1;
    for i in alle_belege:
        
            print ("Tag:",i.datum.day, "Befinden uns: ", temp)
            input("Bestätigung der Eingabe des oben genanten datums")
            browser.implicitly_wait(5)
            temp = temp +1;
            
            #ids = browser.find_elements_by_xpath('//*[@id]')
            #for i in ids:
            #   print(i.get_attribute('id')) 
            #  print("\n")
            elem_ausgabe = browser.find_element_by_id("ausgabe")
            elem_einnahme = browser.find_element_by_id("einnahme")
            elem_datum = browser.find_element_by_id("datum")
            elem_beleg = browser.find_element_by_id("belegtext")
            elem_button = browser.find_element_by_id("speichernButtonLarge")

            elem_einnahme.send_keys(str(i.value)+"00")
            if(i.value < 10):
                elem_beleg.send_keys("Haarberatung")
            elif(i.value < 30):
                elem_beleg.send_keys("Haarschnitt")
            else:
                elem_beleg.send_keys("Haarfärbung")

            print("\n\n\n\n")


foo()