from bs4 import BeautifulSoup
import requests
import openpyxl

excel = openpyxl.Workbook()
sheet = excel.active
sheet.title = "Top Rated Movies"
sheet.append(['Rank', 'Movie_Name', 'Release_Year', 'Duration', 'Type', 'Rating'])
# print(excel.sheetnames)

try:
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3'
    }

    url = requests.get("https://www.imdb.com/chart/top/", headers=headers)

    soup = BeautifulSoup(url.text, 'html.parser')
    
    movies = soup.find('ul',class_='ipc-metadata-list')
    # print(len(movies))

    for movie in movies:

        Movie_Name_R = movie.find('div',class_="ipc-title").a.text

        Rank, Movie_Name = Movie_Name_R.split('.', 1)
        Movie_Name = Movie_Name.strip()

        Release_Year = movie.find('div',class_="cli-title-metadata").span.text
        Duration = movie.find_all('span',class_="cli-title-metadata-item")[1].text
        Type = movie.find_all('span',class_="cli-title-metadata-item")[2].text

        Rating = movie.find('span',class_="ipc-rating-star--rating").text

        print(Rank, Movie_Name, Release_Year, Duration, Type, Rating)
        sheet.append([Rank, Movie_Name, Release_Year, Duration, Type, Rating])

except Exception as e:
    print(e)

excel.save('IMDB Movie Ratings.xlsx')