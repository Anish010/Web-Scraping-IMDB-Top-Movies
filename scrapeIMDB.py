from bs4 import BeautifulSoup
import requests, openpyxl

excel = openpyxl.Workbook()
print(excel.sheetnames)
sheet =excel.active
sheet.title = 'Top Rated Movies'
sheet.append(['Movies Rank', 'Movies Name','IMDB rating', "Year of Release"])

try:
    source = requests.get('https://www.imdb.com/chart/top/')
    #To capture error if url is wrong
    source.raise_for_status()
    
    soup = BeautifulSoup(source.text,'html.parser')
    
    movies = soup.find('tbody',class_='lister-list').find_all('tr')
    
    for movie in movies:
        name = movie.find('td',class_="titleColumn").find('a').text
        rank = movie.find('td',class_="titleColumn").get_text(strip=True).split('.')[0]
        year = movie.find('td',class_="titleColumn").find('span').text.strip('()')
        rating = movie.find('td',class_="ratingColumn").strong.text
        print(f'{rank}  {name}  {rating}  {year}')
        sheet.append([rank,name,rating,year])
    
except Exception as e:
    print(e)
    
excel.save('TopMovies.xlsx')