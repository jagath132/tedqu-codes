import requests
from bs4 import BeautifulSoup

url = "https://news.ycombinator.com/"
page = requests.get(url)
soup = BeautifulSoup(page.content, "html.parser")

headlines = soup.find_all("span" , class_="titleline")

for h in headlines[:5]:
    print("/n", h.text)
