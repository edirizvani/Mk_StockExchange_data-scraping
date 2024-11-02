import asyncio
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
import time
import os
import pandas as pd
from datetime import datetime




def create_selenium_driver(download_dir):
    chrome_options = Options()
    chrome_options.add_experimental_option("prefs", {
            "download.default_directory": download_dir,  # Set download directory
            "download.prompt_for_download": False,        # Disable download prompt
            "download.directory_upgrade": True,           # Enable directory upgrade
            "safebrowsing.enabled": True,                 # Enable safe browsing
            "profile.default_content_settings.popups": 0,  # Disable popups
            "download.restrict_quota": 0                   # Disable quota restrictions
        })
    chrome_options.add_argument("--no-sandbox")  # Optional, helps with some environments
    chrome_options.add_argument("--headless")     # Optional, runs Chrome in headless mode

    service=Service(executable_path="chromedriver.exe")
    driver=webdriver.Chrome(service=service,chrome_options=chrome_options)
    return driver



def get_data_for_x_years(years,stockname):
    timestamp = datetime.now().minute

    download_dir = os.path.join("C:\\Users\\Administrator\\python-project-data-mse\\downloaded_files", stockname)
    os.makedirs(download_dir, exist_ok=True)  # Create the directory
    
    driver = create_selenium_driver(download_dir)

    current_year = datetime.now().year
    driver.get("https://www.mse.mk/mk/stats/symbolhistory/"+stockname)

    for year in range(current_year,current_year-years,-1):

        from_element=driver.find_element(By.ID,"FromDate")
        to_element=driver.find_element(By.ID,"ToDate")

        from_element.clear()
        from_element.send_keys("01.1."+str(year)) 
        to_element.clear()
        to_element.send_keys("31.12."+str(year))   

        prikazhi=driver.find_element_by_css_selector("#report-filter-container .container-end .btn")
        prikazhi.click()

        button=driver.find_element(By.ID,"btnExport")

        # button.__setattr__("download",file_name)
             
        button.click()
        # driver.execute_script(f"arguments[0].setAttribute('download', '{file_name}');", button)
    
    driver.quit()
    
    return join_data(download_dir,stockname)

def join_data(download_dir,stockname):
    dataframes_list=[]
    for s in [""]+[" (" + str(i) + ")" for i in range(1,5)]:
        downloaded_file = os.path.join(download_dir,  f"Историски податоци"+s+".xls")
       
        if os.path.exists(downloaded_file):
            print("File downloaded successfully!")
        else:
            print("File not found in the specified directory.")
            continue  # Skip to the next iteration if the file is not found

      

        tables = pd.read_html(downloaded_file)
        if tables:  # Check if any tables were found
            df_now = tables[0]
            dataframes_list.append(df_now)

    output_directory = "C:\\Users\\Administrator\\python-project-data-mse\\All_Stock_Data"

    # Concatenate all DataFrames in the list into a single DataFrame
    if dataframes_list:
        final_dataframe = pd.concat(dataframes_list, ignore_index=True)

        # Export to CSV
        output_file = os.path.join(output_directory, stockname)
        final_dataframe.to_csv(output_file, index=False, encoding='utf-8-sig')
        print(f"Data exported successfully to {output_file}")
        
        return final_dataframe
    else:
        print("No DataFrames were created.")
        return None
    

def get_stock_data(codes):
    for c in codes:
        get_data_for_x_years(10, c)  
    


download_dir = "./"
driver =  create_selenium_driver(download_dir)
driver.get("https://www.mse.mk/mk/stats/symbolhistory/ALK")

codes = [option.get_attribute("value") for option in driver.find_elements(By.CSS_SELECTOR, "#Code option")]
filtered_codes = [
    c for c in codes
    if not (c.startswith('E') or any(char.isdigit() for char in c))
]
get_stock_data(filtered_codes)
driver.quit()



 # Replace with your desired output directory
# final_df = get_data_for_x_years(5, output_directory)  # Replace 5 with the number of years you want





        

       






