from selenium import webdriver #Web driver activities
import logging #Activity logging
from configparser import ConfigParser #Configuration read
import os #Path
import sys #System paths
import re
import csv
import ctypes #message box
from selenium.webdriver.firefox.options import Options
import pyodbc
from selenium import webdriver
from selenium.webdriver.firefox.firefox_binary import FirefoxBinary
from selenium.webdriver.common.by import By #Element find
from selenium.webdriver.support import expected_conditions as EC #Error handling
from selenium.webdriver.support.ui import WebDriverWait #Web driver wait
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
import time #Sleep
from pathlib import Path #Path conversions
import pyautogui
from time import sleep # this should go at the top of the file
import pygetwindow
import shutil

# server='prod-di-db'
# user='tapadmin'
# password='Welcome@2021'
# database='Freedman_Reuploading'
# ConfigPath = os.path.dirname(os.path.realpath(__file__)) + '\\config.ini'
ConfigPath = r'config.ini'
logtemplate = "An exception of type {0} occurred. Arguments:\n{1!r}"
firefox_location='./'
geckodriver_path =r'Z:\Renuka\Geckodriver\geckodriver-v0.34.0-win64\geckodriver.exe'
global wait, driver, webpage_url, user_name, password, sq_1, sq_2, sq_3, sq_4,onlyfiles
global sql_server_name, sql_user_name, sql_password, sql_db, SQLconnection

#------------------------------------------------------

""" Function for reading config file """
def read_config_file():
    global webpage_url, user_name, password, sq_1, sq_2, sq_3, sq_4
    global sql_server_name, sql_user_name, sql_password, sql_db
   
    try:
        #configuration entries
        config = ConfigParser()
        config.read(ConfigPath)
        webpage_url = config.get ("Data", "webpage_URL")
        user_name = config.get ("Data", "user_name")
        password = config.get ("Data", "password") 
        sql_server_name = config.get ("SQL", "server")
        sql_user_name = config.get ("SQL", "user")
        sql_password = config.get ("SQL", "password")
        sql_db = config.get ("SQL", "database")
        return(0)
    except Exception as ex:
        print('config file read error',ex)
        return(-1)

#------------------------------------------------------

def setup():
    global  download_dir, webpage_url, driver, wait
    binary = FirefoxBinary(firefox_location)
    profile = webdriver.FirefoxProfile()
    profile.set_preference("browser.download.folderList", 2)
    profile.set_preference("browser.download.manager.showWhenStarting", False)
    # profile.set_preference("browser.download.dir", download_dir)
    profile.set_preference("browser.helperApps.neverAsk.saveToDisk", "text/css, application/vnd.openxmlformats-officedocument.wordprocessingml.document, application/octet-stream, application/word, application/wordpad, image/png, image/bmp, image/jpeg, application/pdf, text/csv, text/html, text/plain, application/docx")
    profile.set_preference("pdfjs.disabled", True)
    profile.set_preference("plugin.disable_full_page_plugin_for_types", "text/css, application/vnd.openxmlformats-officedocument.wordprocessingml.document, application/octet-stream, application/word, application/wordpad, image/png, image/bmp,, image/jpeg, application/pdf, text/csv, text/html, text/plain, application/docx")
    options = Options()
    options.profile = profile
    driver = webdriver.Firefox(options=options)
    
    wait = WebDriverWait(driver, 10)
    driver.maximize_window()
    driver.get('https://hcm.paycor.com/authentication/signin')
    time.sleep(1)
    wait = WebDriverWait(driver, 10)
    return 0

#------------------------------------------------------

def searchAndUpload(dst_id, src_id):
    try:
        sea_flg = 0
        sea_cnt = 0
        while(1):
            try:
                sea_cnt = sea_cnt + 1
                try:
                    """ Clicks Manage People """
                    xpath = "//a[@href='//HCM.paycor.com/people/?companyId=201255&clientId=157838'][contains(text(),'Manage People')]" #user name
                    element = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
                    element.click()
                    time.sleep(2)
                except:
                    try:
                        xpath = "//a[@href='//HCM.paycor.com/people/?companyId=201255&clientId=157838'][contains(text(),'Manage People')]" #user name
                        element = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
                        element.click()
                        time.sleep(2)
                    except:
                        pass
                driver.get("https://hcm.paycor.com/people/?companyId=201255&clientId=157838")
                time.sleep(2)
                paycor_employee=dst_id
                print(paycor_employee)
                xpath  = "//input[contains(@id,'search_')]" # path for search field
                driver.find_elements(By.XPATH, xpath).clear()
                element = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
                element.click()
                time.sleep(1)
                element.send_keys(paycor_employee)
                time.sleep(2)
                break
            except:
                if sea_cnt > 5:
                    sea_flg = 1
                    break
                else:
                    pass
        if sea_flg == 1:
            return(-1)

        try:
            """ XPath to find the element that contains the message "No people to display."(id not found in system),
            if path found then it writes on file  """
            xpath = '//*[text()=" No people to display."]'
            element = driver.find_element(By.XPATH,xpath)
            txt=element.text
           
            if txt.strip() == "No people to display.":
                print(txt)
                with open("not found in system.csv",'a+') as f:
                    f.write(f'{dst_id},{src_id},{txt}\n')
                return 0
        except:
            """ List all the search result """ 
            xpath = "//div[contains(@id,'personCard_1')]"
        elements = driver.find_elements(By.XPATH,xpath)
        no_of_employees = len(elements)
        time.sleep(2)
        for i in range(0,no_of_employees):
            
            xpath1 = "//div[contains(@id,'personCard_1_"+str(i)+"')]/.//div[contains(@class,'empNumber')]"
            element1 = driver.find_element(By.XPATH, xpath1).get_attribute("title")
            if str(element1).strip()=='#'+str(paycor_employee):
                #xpath3 = ""
                xpath2 = "//div[contains(@id,'personCard_1_"+str(i)+"')]/.//span[@class='name secondaryCTA']"
                element2 = wait.until(EC.element_to_be_clickable((By.XPATH, xpath2)))
                element2.click()
                break
           
        # time.sleep(10)
           
        time.sleep(2)
        
        try:
            # xpath = "//*[@id='3000'][@aria-label='Position'][@aria-expanded='false']"
            """Assignment dropdown"""
            xpath = '//*[text()="Assignment"]'
            element = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
            element.click()
            time.sleep(1)
        except:
            pass
        
        try:
            # xpath = '//*[@class="leaf-text__1llFL" and text() = "Documents"]'
            """ Documents path """
            xpath = '//*[text()="Documents"]'
            element = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
            element.click()
            time.sleep(1)
        except:
            try:
                # xpath = "/html/body/div[1]/div/div[3]/div[2]/div/div[2]/nav/ul/nav[3]/ul/li[1]/a/span/span[2]"
                xpath = '//*[@id="person-left-navigation"]/div[2]/nav/ul/nav[3]/ul/li[1]/a/span/span'
                element = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
                element.click()
                time.sleep(2)
            except:
                pass
        try:
            xpath = "//span[normalize-space()='Unfiled Documents']"
            element = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
            element.click()
            time.sleep(2)
            print("clicked on unfiled documents")
        except Exception as h:
            try:
                xpath = "//span[normalize-space()='Unfiled Documents']"
                element = driver.find_element(By.XPATH,xpath)
                element.click()
            except Exception as h:
                print("error h is:", h)
        i=0
        count_xpath = '//*[@id="table_now0asezt_body"]'
        count_xpath='//*[@id="panel_myDocuments"]/div/div[2]/section/div[2]'
        count_xpath='//*[@id="table_5zt6xg4x1_body"]/div'
        # div_element = driver.find_element(By.XPATH,count_xpath)
        """ Will find number of rows(documents) """
        xpath = '//*[@class="table-content_cell__Ei7OF"]/span[@class="style_documentName__XbpI-"]'
        child_elements = driver.find_elements(By.XPATH,xpath)
        print(len(child_elements))
        no_of_childs = len(child_elements)
        csv_file = open('./report.csv','a',newline='')
        csv_writer = csv.writer(csv_file)
        csv_file2 = open('./report_no_files.csv','a',newline='')
        csv_writer2 = csv.writer(csv_file2)
        '''if no_of_childs == 1 : 
            xpath = '/html/body/div[1]/div[2]/div[2]/div/div/div[2]/div/div[2]/section/div[2]/div/div'
            xpath = '//*[text()="No data to display.]'
            element = driver.find_element(By.XPATH,xpath)
            if element.text == 'No data to display.':
                csv_writer2.writerow([src_id,'no files'])


                try:
                    xpath = "//*[@id='Header_NavigateToEmployeeListButton']"
                    element = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
                    element.click()
                    time.sleep(2)  
                except:
                    xpath = "/html/body/div[1]/paycor-employee-navigation-bar/quick-links/button"
                    element = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
                    element.click()
                    time.sleep(2)
                        
                return 0'''
        csv_writer.writerow([src_id,no_of_childs])
        for i in range(1,no_of_childs+1):
            """ Extracts Category """
            # xpath = '/html/body/div[1]/div[2]/div[2]/div/div/div[2]/div/div[2]/section/div[2]/div['+str(i)+']/div[1]/span'
            xpath = '//*[@class="table-content_row__31sbr"]['+str(i)+']/div[1]/span' 
            element = driver.find_element(By.XPATH,xpath)
            category = element.text
            """ File name """
            # xpath = '/html/body/div[1]/div[2]/div[2]/div/div/div[2]/div/div[2]/section/div[2]/div['+str(i)+']/div[2]/span'
            xpath = '//*[@class="table-content_row__31sbr"]['+str(i)+']/div[2]/span'
            element = driver.find_element(By.XPATH,xpath)
            file_name =element.text
            """ Modified date """
            # xpath = '/html/body/div[1]/div[2]/div[2]/div/div/div[2]/div/div[2]/section/div[2]/div['+str(i)+']/div[3]'
            xpath = '//*[@class="table-content_row__31sbr"]['+str(i)+']/div[3]'
            element = driver.find_element(By.XPATH,xpath)
            date_uploaded = element.text
            """ Access level/ Visibility"""
            # xpath = '/html/body/div[1]/div[2]/div[2]/div/div/div[2]/div/div[2]/section/div[2]/div['+str(i)+']/div[4]'
            xpath = '//*[@class="table-content_row__31sbr"]['+str(i)+']/div[4]'
            element = driver.find_element(By.XPATH,xpath)
            visibility = element.text
            csv_writer.writerow([category,file_name,date_uploaded,visibility])
        csv_writer.writerow(['-----------------------------------------------'])
                 
        
    except Exception as h:
        print('employee name error',h)
        f= open('employee name error.csv',"a")
        f.write(f'{src_id},{dst_id}')
        f.write('\n')
        f.close()
        # return(-1)
        try:
            xpath = '//button[@title="Back to employee list"]'
            element = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
            element.click()
            time.sleep(2)  
        except:
            pass
        return -1
    try:
        xpath = '//button[@title="Back to employee list"]'
        element = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
        element.click()
        time.sleep(2)  
    except:
        xpath = "/html/body/div[1]/paycor-employee-navigation-bar/quick-links/button"
        element = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
        element.click()
        time.sleep(2)
    return 0

#------------------------------------------------------       

def login():
    global user_name, driver, password, wait, driver, sq_1, sq_2, sq_3, sq_4
    window_before = driver.window_handles[0]

    xpath = "//input[@id='Username']" #user name 
    element = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
    element.click()
    time.sleep(1)
    element.send_keys("tap_roadsafe50")
    time.sleep(1)
    xpath = "//input[@id='Password']" #password
    element = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
    element.click()
    time.sleep(1)
    element.send_keys("Tap@2024innov")
    time.sleep(1)    
    xpath = "//button[contains(text(),'Sign In')]" #login button
    element = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
    element.click()
    time.sleep(1)
    # time.sleep(30)
    s=input()
    print("login success")

#------------------------------------------------------

""" Function used for connecting to database """

def connectSQL():
    global sql_server_name, sql_user_name, sql_password, sql_db, SQLconnection
    try:
        SQLconnection = pyodbc.connect('Driver={SQL Server};'
                            'Server=' + sql_server_name + ';'
                            'Database=' + sql_db + ';'
                            'UID=' + sql_user_name + ';'
                            'PWD=' + sql_password + ';'
                            'Trusted_Connection=no;')
        return(0)
    except Exception as ex:
        print('SQL connection error')
        return(-1)

#------------------------------------------------------

def main():
    global SQLconnection
    read_config_file()
    connectSQL()
    setup()
    login()
    emp_ids=set()
    cnt=0
    emp_flg_start=1
    rows=[]

    select_statement = "SELECT [Employee_ID],[Last_Name],[First_Name],[Reports]"
    select_statement = select_statement + "FROM [RoadSafe].[dbo].[Emp_List_Termed];"
    # select_statement += "WHERE [Paycor_Client_ID] = '170678';"
    cursor = SQLconnection.cursor()
    cursor.execute(select_statement)
    employee_details = cursor.fetchall()
    cursor.close()
    """ First we will fetch data from DB for list and then we will loop for all employees to generate a report"""

    for i in range(0,2500):
        each_employee = employee_details[i]
        if each_employee[0] :
            src_id=each_employee[0].strip()
            src_id = str(src_id)
            dst_id = each_employee[0].strip()
            isDownloaded = each_employee[3]
            if not isDownloaded:
                res = searchAndUpload(dst_id, src_id)
                if res == 0:
                        """ Updates emp list DB after data written to report all files """
                            
                        cursor = SQLconnection.cursor()
                        sql_update = """UPDATE [RoadSafe].[dbo].[Emp_List_Termed] SET [Reports] = ?
                        WHERE  [Employee_ID] =? ;"""
                        cursor.execute(sql_update,(1,src_id))
                        SQLconnection.commit()
                        cursor.close()
                        print(src_id," : updated in Emp_List")
            
                else :
                            print('Employee error ')
                            continue
main()
