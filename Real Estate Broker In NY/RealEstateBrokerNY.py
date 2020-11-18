import requests
from bs4 import BeautifulSoup
import pandas as pd
def get_html(page):
	headers={'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/86.0.4240.198 Safari/537.36'}
	url =f'https://hubbiz.com/New-York,NY/search?geo_id=&near=New+York%2C+NY&page={page}&q=Real+Estate+Broker&qf=home'
	response=requests.get(url,headers)
	html=BeautifulSoup(response.content,'html.parser')
	return html

def transform_data(html):
	broker_card=html.find_all('div',class_='container-list')
	for item in broker_card:
		try:
			name=item.find('h2',{'itemprop':'name'}).text.strip()
		except:
			name=''
		try:
			cat_in=item.find('span',class_='c').text.strip()
		except:
			cat_in=''
		try:
			street=item.find('p',{'itemprop':'streetAddress'}).text.strip()
		except:
			street=''
		try:
			business_city = item.find('p',class_='business_city_result').text.strip()
		except:
			business_city=''
		try:
			address_locality = item.find('span',{'itemprop':'addressLocality'}).text.strip()
		except:
			address_locality=''
		try:
			address_region = item.find('span',{'itemprop':'addressRegion'}).text.strip()
		except:
			address_region=''
		try:
			telephone = item.find('div',class_='phone_biz_stub').text.strip()
		except:
			telephone=''


		broker={
			'name':name,
			'cat_in':cat_in,
			'street':street,
			'business_city':business_city,
			'address_locality':address_locality,
			'address_region':address_region,
			'telephone':telephone
		}
		brokerlist.append(broker)
brokerlist=[]
for i in range(1,21):
	print(f'Getting Page {i}')
	html=get_html(i)
	transform_data(html)
df=pd.DataFrame(brokerlist)
print(df.head())
df.to_csv('RealEstateBrokerNY.csv')
print(len(brokerlist))