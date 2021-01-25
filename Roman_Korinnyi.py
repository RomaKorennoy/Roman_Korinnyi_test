from bs4 import BeautifulSoup
import requests as req
import xlrd
import pandas as pd
from omdbapi.movie_search import GetMovie



def parse():

    ''' Taking information from the website'''

    titles, years = [], []
    directors, actors, awards, plots, countries = [],[],[],[],[]
    
    html = BeautifulSoup(open("imdb_most_popular_movies_dump.html"), 'lxml')
    tds =  html.find_all('td', class_ = 'titleColumn')
    
    
    for td in tds:
        a = td.find('a')
        titles.append(a.text)
        year = td.find('span', class_ = 'secondaryInfo').text
        years.append(int(year[1:5]))


    ''' Taking information from the OMDb'''

    for title in titles:
        movie = GetMovie(title=title, api_key='767f674e')
        mov_data = movie.get_data('Director', 'Actors', 'Awards', 'Plot','Country')
        directors.append(mov_data['Director'])
        actors.append(mov_data['Actors'])
        awards.append(mov_data['Awards'])
        plots.append(mov_data['Plot'])
        countries.append(mov_data['Country'])

    

 

    df = pd.DataFrame()
    df['Name'] = titles
    df['Year'] = years
    df['Director'] = directors
    df['Actors'] = actors
    df['Awards'] = awards
    df['Plot'] = plots
    df['Country'] = countries
    

    writer = pd.ExcelWriter('./habr.xlsx', engine= 'xlsxwriter')
    df.to_excel(writer, sheet_name = 'Лист1', index=False)

    writer.sheets['Лист1'].set_column('A:A', 100)
    writer.sheets['Лист1'].set_column('B:B', 30)
    writer.sheets['Лист1'].set_column('C:C', 100)
    writer.sheets['Лист1'].set_column('D:D', 100)
    writer.sheets['Лист1'].set_column('E:E', 100)
    writer.sheets['Лист1'].set_column('F:F', 100)
    writer.sheets['Лист1'].set_column('G:G', 100)

    writer.save()

parse()