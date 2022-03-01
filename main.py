from os import sep
from selenium.webdriver.firefox.options import Options
from selenium import webdriver
from bs4 import BeautifulSoup
import time
from tqdm import tqdm

import pandas as pd
import csv

class Data:

    all_data : list = []
    reloadTime : int
    total_pages : int

    def __init__(self,reloadTime:int, total_pages:int):
        self.reloadTime = reloadTime
        self.total_pages = total_pages

    def get_wood(self,driver:webdriver)->dict:
        '''Loops through all the pages to get all the wood articles'''
        
        for page in tqdm(range(1,self.total_pages+1)):
            URL = f'https://www.martenshout.be/nl/collection/hout?page={page}&numberOfItems=12&sortBy=title&sorting=ASC'
            driver.get(URL)
            #sleep so the page can load, hardcoded: this can be a problem later on
            time.sleep(self.reloadTime)
            soup = BeautifulSoup(driver.page_source,'html.parser')
            
            for div in soup.find_all('div',class_='product-list-item col-lg-4 col-md-6 col-sm-12 mb-5'):
                temp_list = []
                name_product =div.find('span').text
                exclusief, inclusief, eenheid = self.split_currencies_wood(div.find('div',class_='price').text)
                temp_list.append(name_product)
                temp_list.append(exclusief)
                temp_list.append(inclusief)
                temp_list.append(eenheid)
                self.all_data.append(temp_list)
        driver.close()
        return self.all_data

    def save_wood_as_csv(self)->None:
        '''Saves the list as a csv file'''
        col = ['name','exclusief-BTW','inclusief-BTW','eenheid']
        
        with open('wood_prices.csv', 'w',encoding='UTF8') as f:
            # using csv.writer method from CSV package
            write = csv.writer(f)
            write.writerow(sep)
            write.writerow(col)
            write.writerows(self.all_data)


    def split_currencies_wood(self,raw_text:str):
        """Splits raw prices into: inclusief btw per eenheid AND exclusief btw"""
        arr = []
        inclusief_index = 4
        eenheid_index= 3
        exclusief_index = 1
        #getting rid of the noice
        raw_text = raw_text.replace(u'\xa0', u' ')
        arr = raw_text.split(' ')
        
        return arr[exclusief_index],arr[inclusief_index],arr[eenheid_index]


def main():
    #initialize the chromedriver
    options = Options()
    options.headless = True
    driver = webdriver.Firefox(options=options)

    #make object from Data
    data= Data(3,14)
    _data = data.get_wood(driver)
    data.save_wood_as_csv()
    


if __name__ == '__main__':
    main()


