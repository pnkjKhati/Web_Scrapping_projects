from bs4 import BeautifulSoup
import requests, openpyxl

excel = openpyxl.Workbook()
print(excel.sheetnames)
sheet=excel.active
sheet.title = 'Top_50_sport_movies'
print(excel.sheetnames)
sheet.append(['Movie Rank','Movie Name','Year','Rating'])



try:

    source = requests.get('https://www.imdb.com/search/title/?genres=sport&title_type=feature&explore=genres')
    source.raise_for_status()
    
    soup = BeautifulSoup(source.text, 'html.parser')
    movies = soup.find('div', class_='lister-list').find_all('div', class_='lister-item mode-advanced')
    for movie in movies:
        
        
        rank = movie.find('div',class_='lister-item-content').h3.span.text.strip('.')
        if rank=='37':
            continue
        name = movie.find('div', class_='lister-item-content').h3.a.text
        year = movie.find('div', class_='lister-item-content').h3.find('span', class_='lister-item-year text-muted unbold').text.strip('()')
        rating = movie.find('div', class_='lister-item-content').div.div.get('data-value')
        print(rank, name, year,rating)
        sheet.append([rank, name, year, rating])
            
except Exception as e:
    print(e)

excel.save('IMDB Top 50 Sport Movie Ratings.xlsx')