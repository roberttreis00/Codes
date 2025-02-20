# Conta a quantidade de links que varreu para achar o que estou procurando
from urllib.request import urlopen, Request
from bs4 import BeautifulSoup

link_pagina_now = 'https://en.wikipedia.org/wiki/Eric_Idle'
link_para_encontrar = 'https://en.wikipedia.org/wiki/Kevin_Bacon'
tentativas = 0

headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36"
}


def conjunto_links(l):
    r = Request(l, headers=headers)
    links = list()
    open_link = urlopen(r)
    bs = BeautifulSoup(open_link, 'html.parser')
    for link in bs.find_all('a'):
        if 'href' in link.attrs and 'http' == link['href'][:4]:
            links.append(link['href'])

    return links
