import time
import re
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.keys import Keys
from bs4 import BeautifulSoup
import csv
from openpyxl import Workbook

driver = webdriver.Chrome()
driver.wait = WebDriverWait(driver, 5)

url = "http://www.imdb.com/search/title"
driver.get(url)
#print(driver.page_source)

element = driver.find_element_by_xpath('//*[@id="title_type-1"]')
element.click()

date_min = 2016
date_max = 2018
textbox_date_min = driver.find_element_by_name('release_date-min')
textbox_date_max = driver.find_element_by_name('release_date-max')
textbox_date_min.send_keys(date_min)
textbox_date_max.send_keys(date_max)

#select = Select(driver.find_element_by_name('user_rating-min'))

element = driver.find_element_by_xpath("//select[@name='user_rating-min']")
all_options = element.find_elements_by_tag_name("option")
for option in all_options:
    if option.get_attribute("value") == "5.0":
        option.click()
        break

element = driver.find_element_by_xpath("//select[@name='user_rating-max']")
all_options = element.find_elements_by_tag_name("option")
for option in all_options:
    if option.get_attribute("value") == "10":
        option.click()
        break


element = driver.find_element_by_xpath("//select[@name='countries']")
all_options = element.find_elements_by_tag_name("option")
for option in all_options:
    if option.get_attribute("value") == "us":
        option.click()
        break


element = driver.find_element_by_xpath("//select[@name='languages']")
all_options = element.find_elements_by_tag_name("option")
for option in all_options:
    if option.get_attribute("value") == "en":
        option.click()
        break


element = driver.find_element_by_xpath("//select[@id='search-count']")
all_options = element.find_elements_by_tag_name("option")
for option in all_options:
    if option.get_attribute("value") == "250":
        option.click()
        break

#search
element = driver.find_element_by_xpath('//*[@id="main"]/p[3]/button')
element.click()



wb = Workbook()
ws = wb.active
ws.append(['Year','Rating','MetaScore','MovieName','Director','Actors','Votes'])

count = 0
while count in range(8):

    html = driver.page_source
    soup = BeautifulSoup(html,"html.parser")
    #print(soup)

    rows = []

    all_lists = soup.find("div", attrs = {'class' : 'lister-list'})
    for entry in all_lists.find_all("div", attrs = {'class' : 'lister-item-content'}):
        #movie_name = entry.find("a").text
        #cells.append(movie_name)
        cells = []
        y = entry.find('span',{'class':'lister-item-year text-muted unbold'}).text
        regex = r"\d{4}"
        year  = ''.join(re.findall(regex, y)) #findall returns a list contains one string ex: ['2017']
        cells.append(int(year))

        rating = entry.find('strong').text
        cells.append(float(rating))

        if entry.find('div', attrs = {'class' : 'inline-block ratings-metascore'}):
            meta_score = int(entry.find('span',{'class':'metascore'}).text)
        else:
            meta_score = 0 #meta_score not present
        cells.append(meta_score)

        i = 1
        actors = []
        for a in entry.find_all('a', href=True):
            #print("Found: {} ".format(i), a.text)
            if i == 1:
                movie_name = a.text
            elif i == 2:
                if(a.text == "See full summary")  :
                    continue
                else:
                    director = a.text
            else:
                actors.append(a.text)
            i += 1

        cells.append(movie_name)
        cells.append(director)

        if not actors:
            cells.append("")
        else:
            cells.append(':'.join(actors))

        votes = entry.find('span',{'name':'nv'}).text
        v = votes.replace(',','')

        cells.append(int(v))
        rows.append(cells)

    if(count  == 0):
        element = driver.find_element_by_xpath('//*[@id="main"]/div/div[4]/a')
        #element = driver.find_element_by_xpath('//*[@id="main"]/div/div/div[4]/div/a')
    else:
        element = driver.find_element_by_xpath('//*[@id="main"]/div/div[4]/a[2]')
         #element = driver.find_element_by_xpath('//*[@id="main"]/div/div/div[4]/div/a[2]')
    element.click()
    #print("\n\n\n\n\n\n\n\ncount is ", count)
    count+=1

    '''
    with open('imdb_scraper.csv','a+') as csvfile:
        writer = csv.writer(csvfile)
        for m in rows:
            writer.writerow(m)
    '''

    for i in rows:
        ws.append(i)

ws2 = wb.copy_worksheet(ws)
ws3 = wb.copy_worksheet(ws)
ws4 = wb.copy_worksheet(ws)
ws5 = wb.copy_worksheet(ws)

wb.save("imdb.xlsx")

df = pd.read_excel('imdb.xlsx', 'Sheet')   #print(df.head(100))

writer = pd.ExcelWriter('output.xlsx', engine='xlsxwriter')


df.to_excel(writer, index=False, sheet_name='OriginalIMDB')
sorted_by_year = df.sort_values(['Year'], ascending=False)
sorted_by_year.to_excel(writer, index=False, sheet_name='SortedByYear')

sorted_by_rating = df.sort_values(['Rating'], ascending=False)
sorted_by_rating.to_excel(writer, index=False, sheet_name='SortedbyRating')

sorted_by_votes = df.sort_values(['Votes'], ascending=False)
sorted_by_votes.to_excel(writer, index=False, sheet_name='SortedByVotes')

sorted_by_meta = df.sort_values(['MetaScore'], ascending=False)
sorted_by_meta.to_excel(writer, index=False, sheet_name='SortedByMetaScore')

writer.save()
