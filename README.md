# webscraping
Automating webscraping using Python - Pandas and Selenium

Write a Python Script to scrape the website www.imdb.com using Selenium and Beautiful Soup. Script should scrape the following details
for at least latest 2000 movies and generate an Excel Sheet.

1. Name of the Movie
2. Year of Release
3. Director of Movie
4. Actors of the Movie
5. IMDB Rating
6. Metascore
7. Number of Votes

Generate the following sheets in excel through automation

1. Sheet 1 - Should contain the original scraped data
2. Sheet 2 - Sort the rows based on 'Year of Release' and the latest Year should appear on top
3. Sheet 3 - Sort the rows based on IMDB Rate from original Sheet 1 scraped data with highest rate appear on top
4. Sheet 4 - Sort the rows based on Metascore from original Sheet 1 scraped data with highest Metascore appear on top
5. Sheet 5 - Sort the rows based on Number of Votes from original Sheet 1 scraped data with highest Vote appear on top.


References used -
http://selenium-interview-questions.blogspot.in/search/label/Selenium%20Webdriver%20in%20Python
http://openpyxl.readthedocs.io/en/stable/filters.html
https://www.dataquest.io/blog/excel-and-pandas/
https://stackoverflow.com/questions/26474693/excelfile-vs-read-excel-in-pandas
