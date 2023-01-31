from bs4 import BeautifulSoup
import requests
import openpyxl

excel = openpyxl.Workbook()
sheet = excel.active
sheet.title = 'Top rated movies'
sheet.append(['Movie rank', 'Movie name', 'Year of release', 'IMDB rating'])


try:
    source = requests.get('https://www.imdb.com/chart/top/')
    source.raise_for_status()

    soup = BeautifulSoup(source.text, 'html.parser')
    movies = soup.find('tbody', class_= "lister-list").find_all('tr')
    for movie in movies:

        name = movie.find('td', class_= "titleColumn").a.text
        rank = movie.find('td', class_= "titleColumn").get_text(strip = True).split('.')[0]
        year = movie.find('td', class_= "titleColumn").span.text.strip('()')
        IMDBrating = movie.find('td', class_= "ratingColumn imdbRating").strong.text

        sheet.append([rank,name,year,IMDBrating])
       
except Exception as e:
    print(e) 

excel.save('IMDB Movie rating.xlsx')