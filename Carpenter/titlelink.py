from selenium import webdriver
import pandas as pd
from selenium.webdriver.chrome.options import Options as ChromeOptions
import openpyxl
import time
import random


def get_pagelink(page):
	driver=webdriver.Chrome('../Drivers/chromedriver')
	url=f'https://www.truelocal.com.au/search/carpenter/australia?page={page}'
	driver.get(url)
	driver.implicitly_wait(10)
	carpenters=driver.find_elements_by_class_name('item-title')
	for carpenter in carpenters:
		link=carpenter.get_attribute('href')
		carpenterlist.append(link)
		print(carpenter.get_attribute('href'))
	driver.quit()
carpenterlist=[]
for i in range(972,973):
	print("Getting page : %d"%i)
	get_pagelink(i)
df=pd.DataFrame(carpenterlist)
print(df.head())
df.to_csv('carpenter.csv')
print(len(carpenterlist))