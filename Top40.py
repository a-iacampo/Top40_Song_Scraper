import pandas as pd
from bs4 import BeautifulSoup
import urllib.request
from openpyxl import load_workbook
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox


# ---- Function call ----

# Scrape song metadata from website and saves list into Excel file
def scrapeSongs():
    url = entry.get().strip()
    path = pathEntry.get()

    # Validate input
    if not url:
        messagebox.showwarning('Warning', 'Please enter a url')
        return

    if not 'www.at40.com/charts/' in url:
        messagebox.showwarning('Warning', 'Please enter a url from AT40')
        return

    if not path:
        messagebox.showwarning('Warning', 'Please import a file')
        return

    if not path.endswith('.xlsx'):
        messagebox.showwarning(
            'Warning', 'Please import an Excel file (.xlsx)')
        return

    try:
        source = urllib.request.urlopen(url).read()
        soup = BeautifulSoup(source, 'lxml')
    except:
        messagebox.showerror(
            'Error', "This site can't be reached")
        return

    # Get date
    title = soup.title.string.split(" ")
    year = title[5]
    date = title[3] + " " + title[4] + " " + year

    # Get Excel file
    try:
        xlsx = pd.ExcelFile(path)
        df = pd.read_excel(xlsx, year)
        sheetExist = True

        if not all([item in df.columns for item in ['Song', 'Artist', 'Date Added', 'Downloaded']]):
            df = pd.DataFrame([], columns=['Song', 'Artist',
                                           'Date Added', 'Downloaded'])
    except:
        sheetExist = False
        df = pd.DataFrame([], columns=['Song', 'Artist',
                                       'Date Added', 'Downloaded'])

    try:
        dfSheets = pd.concat(pd.read_excel(
            xlsx, sheet_name=None), ignore_index=True)
    except:
        messagebox.showerror('Error', "File path not found")
        return

    if dfSheets.empty:
        dfSheets = pd.DataFrame([], columns=['Song', 'Artist',
                                             'Date Added', 'Downloaded'])

    # Scraping
    try:
        tracks = soup.find(
            'div', class_='component-container component-chartlist block').find_all('figcaption')
    except:
        messagebox.showerror(
            'Error', 'Please enter a url from the weekly charts')
        return

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
        with pd.ExcelWriter(path) as writer:
            writer.book = load_workbook(path)
            if sheetExist:
                del writer.book[year]
            newDf.to_excel(writer, sheet_name=year, index=False)
            writer.book.save(path)
    except:
        messagebox.showerror(
            'Error', "Excel file must be closed for the program to execute")
        return

    messagebox.showinfo("Done", "Chart has been successfully scraped")


# Get path to Excel file and writes path in Entry Field

def UploadAction(event=None):
    filename = filedialog.askopenfilename()
    pathEntry.configure(state="normal")
    pathEntry.delete(0, tk.END)
    pathEntry.insert(0, filename)
    pathEntry.configure(state="disabled")


# ------ GUI --------
win = tk.Tk()

tk.Label(win, text="Enter URL").grid(row=0)
win.title("Top40 Scraper")
win.geometry("500x100")
win.resizable(False, False)

entry = tk.Entry(win, width=71)
entry.grid(row=0, column=1, padx=5, pady=5)

pathEntry = tk.Entry(win, width=71)
pathEntry.grid(row=1, column=1, padx=5, pady=5)
pathEntry.configure(state="disabled")

tk.Button(win, text="Import", command=UploadAction
          ).grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
tk.Button(win, text="Scrape", command=scrapeSongs
          ).grid(row=2, column=1, sticky=tk.W, padx=5, pady=5)
tk.Button(win, text="Quit", command=win.quit).grid(
    row=2, column=1, sticky=tk.W, padx=60, pady=5)

win.mainloop()
