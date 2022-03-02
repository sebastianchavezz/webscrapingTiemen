from os import lstat
from selenium.webdriver.firefox.options import Options
from selenium import webdriver
from bs4 import BeautifulSoup
import time
from tqdm import tqdm

import pandas as pd
import csv

class Data:

    reloadTime : int
    
    colmuns : list = ['name','exclusief-BTW','eenheid']
    def __init__(self,reloadTime:int):
        """Needs to recieve the reload time=> faster internet and PC equals less reload time
            Needs to recieve the amount of pages per material , for wood it is 14"""
        self.reloadTime = reloadTime
        

    def get_wood(self,driver:webdriver,total_pages:int,material:str)->list:
        '''Loops through all the pages to get all the wood articles'''
        
        wood : list = []
        for page in tqdm(range(1,total_pages+1)):
            URL = {
                'wood':f'https://www.martenshout.be/nl/collection/hout?page={page}&numberOfItems=12&sortBy=title&sorting=ASC',
                'platen':f'https://www.martenshout.be/nl/collection/constructieplaten?page={page}&numberOfItems=12&sortBy=title&sorting=ASC',
                'fineer':f'https://www.martenshout.be/nl/collection/fineer?page={page}&numberOfItems=12&sortBy=title&sorting=ASC'
            }
            driver.get(URL.get(material))
            #sleep so the page can load, hardcoded: this can be a problem later on
            time.sleep(self.reloadTime)
            soup = BeautifulSoup(driver.page_source,'html.parser')
            
            for div in soup.find_all('div',class_='product-list-item col-lg-4 col-md-6 col-sm-12 mb-5'):
                temp_list = []
                name_product =div.find('span').text
                exclusief, eenheid = self.split_currencies(div.find('div',class_='price').text)
                temp_list.append(name_product)
                temp_list.append(exclusief)
                temp_list.append(eenheid)
                wood.append(temp_list)
        driver.close()
        
        return wood


        
    def save_wood_as_csv(self,data:list)->None:
        '''Saves the list as a csv file'''
        with open('wood_prices.csv', 'w',encoding='UTF8') as f:
            # using csv.writer method from CSV package
            write = csv.writer(f)
            write.writerow(self.colmuns)
            write.writerows(data)

    def save_as_excel(self,data:list,name_file:str)->None:
        """Saves directly as excel file"""
       
        df = pd.DataFrame(data,columns=self.colmuns)
        with pd.ExcelWriter(name_file) as writer:
            df.to_excel(writer) 

    def split_currencies(self,raw_text:str):
        """Splits raw prices into: inclusief btw per eenheid AND exclusief btw"""
        arr = []
        eenheid_index= 3
        exclusief_index = 1
        #getting rid of the noice
        raw_text = raw_text.replace(u'\xa0', u' ')
        arr = raw_text.split(' ')
        
        return arr[exclusief_index],arr[eenheid_index]


def main():
    #initialize the chromedriver
    options = Options()
    options.headless = True
    
    #make object from Data
    data= Data(reloadTime=3)
    #get wood 
    wood = data.get_wood(webdriver.Firefox(options=options),total_pages=14,material='wood')
    data.save_as_excel(wood,"wood_prices.xlsx")
    #get platen
    construtie_platen = data.get_wood(webdriver.Firefox(options=options),total_pages=12,material='platen')
    data.save_as_excel (construtie_platen,'plate.xlsx')
    #get fineer
    fineer = data.get_wood(webdriver.Firefox(options=options),total_pages=2,material='fineer')
    data.save_as_excel (fineer,'fineer.xlsx')

if __name__ == '__main__':
    main()


