
from bs4 import BeautifulSoup
import requests, openpyxl

excel = openpyxl.Workbook()
sheet = excel.active  # active sheet
sheet.title = 'Top rated Movies'  # change the title
print(excel.sheetnames)

sheet.append(['Movie Rank', 'Movie Name', 'Year', 'Rating'])

try:

    source = requests.get('https://www.imdb.com/chart/top/')
    source.raise_for_status()  # trow the error when incorrect/cant connect to url

    soup = BeautifulSoup(source.text, 'html.parser')  # taking html code to a variable

    movies = soup.find('tbody', class_="lister-list").find_all('tr')  # access the tr tag

    for movie in movies:
        name = movie.find('td', class_="titleColumn").a.text

        rank = movie.find('td', class_="titleColumn").get_text(strip=True).split('.')[0]

        year = movie.find('td', class_="titleColumn").span.text.strip('()')

        rating = movie.find('td', class_="ratingColumn imdbRating").strong.text

        print(rank, name, year, rating)
        sheet.append([rank, name, year, rating])


except Exception as e:
    print(e)

excel.save('movie ratings.xlsx')

