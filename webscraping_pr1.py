# -*- coding: utf-8 -*-
"""
Created on Wed Oct 26 14:37:58 2022

@author: Havva
"""

#if you haven't download the libraries yet, you can use the codes below to upload
#pip install requests
#pip install bs4 #-- and if you're using macOS as system, u have to use pip3 install ..... instead of pip install --

#libraries
from bs4 import BeautifulSoup
import requests, openpyxl

excel = openpyxl.Workbook() #create a workbook
sheet = excel.active  #assign first sheet on your excel
sheet.title = 'Top Rated Movies' #rename your first sheet
print(excel.sheetnames)
sheet.append(['Movie Rank', 'Movie Name', 'Year of Release', 'IMDB Rating']) 
#this appending means list's elements will be used for columns

#try and except

try:
    source = requests.get('https://www.imdb.com/chart/top/') 
    source.raise_for_status() #if link doesn't work, code gives an HTTPerror 
    
    soup = BeautifulSoup(source.text, 'html.parser') #take the parser from HTML code on the source
    
    movies = soup.find('tbody', class_="lister-list").find_all('tr') #use tag and class name to find all movie name's text  
    
    for movie in movies:
        
        name = movie.find('td', class_="titleColumn").a.text 
        #use your main tag and class name but the movie's name text is in 'a' sub-tag on HTML source hence u have to define it here
        
        rank = movie.find('td', class_="titleColumn").get_text(strip=True).split('.')[0]
        
        year = movie.find('td', class_="titleColumn").span.text.strip('()')
        
        rating = movie.find('td', class_="ratingColumn imdbRating").strong.text
        
        print(rank, name, year, rating)
        sheet.append([rank, name, year, rating]) #add all data your excel below the first row

    
except Exception as e:
    print(e)
    
excel.save('IMDB Movie Ratings.xlsx') #save your excel doc