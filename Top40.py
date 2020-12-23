import pandas as pd
from bs4 import BeautifulSoup
import urllib.request
from openpyxl import load_workbook

url = 'https://www.at40.com/charts/top-40-238/2020-12-12-december-12-2020-0'

source = urllib.request.urlopen(url).read()
soup = BeautifulSoup(source, 'lxml')

# Get date
title = soup.title.string.split(" ")
year = title[5]
date = title[3] + " " + title[4] + " " + year

# Get Excel file
try:
    xlsx = pd.ExcelFile('AT40 songs.xlsx')
    df = pd.read_excel(xlsx, year)
    sheetExist = True
except:
    sheetExist = False
    df = pd.DataFrame([], columns=['Song', 'Artist',
                                   'Date Added', 'Downloaded'])

dfSheets = pd.concat(pd.read_excel(xlsx, sheet_name=None), ignore_index=True)

if dfSheets.empty:
    dfSheets = pd.DataFrame([], columns=['Song', 'Artist',
                                         'Date Added', 'Downloaded'])

# Scraping
tracks = soup.find(
    'div', class_='component-container component-chartlist block').find_all('figcaption')

# Get songs, check for duplicates & append to Excel file
newSongs = []
for track in tracks:
    try:
        song = track.find('a', class_='track-title').string
        artist = track.find('a', class_='track-artist').string
    except:
        song = track.find('span', class_='track-title').string
        artist = track.find('span', class_='track-artist').string
    else:
        pass

    if not ((dfSheets['Song'] == song) & (dfSheets['Artist'] == artist)).any():
        newSongs.append(
            pd.Series([song, artist, date, None], index=df.columns))

newDf = df.append(newSongs, ignore_index=True)

# Write to Excel file
try:
    with pd.ExcelWriter('AT40 songs.xlsx') as writer:
        writer.book = load_workbook('AT40 songs.xlsx')
        if sheetExist:
            del writer.book[year]
        newDf.to_excel(writer, sheet_name=year, index=False)
        writer.book.save('AT40 songs.xlsx')
except:
    print('Error: Excel file must be closed for the program to execute\nClose the excel file and run the program again')
    exit(0)

print(newDf)
