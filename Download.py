import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
import requests
import pandas as pd

# Read excel file with company details
df = pd.read_excel('Input/List.xlsx')

# switch off chrome automatic bar
options = Options()
options.add_experimental_option("excludeSwitches", ["enable-automation"])
options.add_experimental_option('useAutomationExtension', False)

# read chromedriver
#driver = webdriver.Firefox(executable_path=r"C:\Users\XuanHeng\Desktop\geckodriver.exe")
driver = webdriver.Chrome(chrome_options=options, executable_path=r"chromedriver.exe")
 
driver.get("http://www.google.com")


for index, row in df.iterrows():

    # range of the years we are interested in 
    rang = list(range(row['FirstYear'],row['LastYear'] + 1))
    for x in rang:
        anno = str(x)

        # parola da cercare nella barra di ricerca
        element = driver.find_element_by_name("q");
        element.send_keys(row['CompanyName'] + ' ' + row['TypeReport'] + ' ' + anno + ' filetype:pdf');
        element.submit();

        #tutti i risultati
        results = driver.find_elements_by_xpath("//div[@class='g']//a[not(@class)]");
        
        #primo risultato
        result = results[0]

        #url del primo risultato
        url = result.get_attribute("href")

        #download e rinomina del url trovato
        try:
            r = requests.get(url, stream=True, timeout=10, verify=False, headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/74.0.3729.169 Safari/537.36'})
            chunk_size = 2000
            TypeReport = row['TypeReport']
            Nospace= TypeReport.replace(" ","")
            with open('Output/'+row['CompanyCode'] + '_' + anno + '_' + Nospace + '.pdf', 'wb') as fd:
                for chunk in r.iter_content(chunk_size):
                    fd.write(chunk)
        except:
            pass
        
        element = driver.find_element_by_name("q");
        element.send_keys(Keys.CONTROL, 'a')
        element.send_keys(Keys.BACKSPACE)
    print(row['CompanyName']+' done!')

driver.quit()