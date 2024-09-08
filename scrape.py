from bs4 import BeautifulSoup
import requests, openpyxl

excel = openpyxl.Workbook()
sheet = excel.active
sheet.title = 'Top rated Shows'
print(excel.sheetnames)
sheet.append(['Show Rank', 'Show Name' , 'Year of release', 'IMDB Rating'])
try:
    source = requests.get('https://www.imdb.com/chart/toptv/')
    source.raise_for_status()
    #html doc contents are taken using bs4 then parsed and stored in soup
    soup = BeautifulSoup(source.text,'html.parser')
    #find finds first tag matching the tag name with class listers-list and find_all shows list of contents
    shows = soup.find('tbody', class_='lister-list').find_all('tr')
    
    for show in shows:
        name = show.find('td', class_='titleColumn').a.text
        rank = show.find('td', class_='titleColumn').get_text(strip=True).split('.')[0]
        year = show.find('td', class_='titleColumn').span.text.strip('()')
        rating = show.find('td', class_='ratingColumn imdbRating').strong.text
            
  
        print(rank, name, year, rating)
        sheet.append([rank, name, year, rating])

except Exception as e:
    print(e)

excel.save('IMDB top rated shows.xlsx')