import requests
from bs4 import BeautifulSoup
import pandas as pd
def get_html(page):
	headers={'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/86.0.4240.198 Safari/537.36'}
	url =f'https://www.indeed.com/jobs?q=python+developer&l=Los+Angeles%2C+CA&start={page}'
	response=requests.get(url,headers)
	html=BeautifulSoup(response.content,'html.parser')
	return html

def transform_data(html):
	job_cards=html.find_all('div',class_='jobsearch-SerpJobCard')
	for item in job_cards:
		title=item.find('a',class_='jobtitle').text.strip()
		company=item.find('span',class_='company').text.strip()
		location=item.find('span',class_='location').text.strip()
		try:
			salary=item.find('span',class_='salaryText').text.strip()
		except:
			salary=''
		summary=item.find('div',class_='summary').text.strip()

		job={
			'title':title,
			'company':company,
			'location':location,
			'salary':salary,
			'summary':summary
		}
		joblist.append(job)
joblist=[]
for i in range(0,661,10):
	print(f'Getting Page {i}')
	html=get_html(i)
	transform_data(html)
df=pd.DataFrame(joblist)
print(df.head())
df.to_csv('jobs.csv')
print(len(joblist))