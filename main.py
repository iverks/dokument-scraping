import os
import shutil
import docx2txt
import pathlib
import re
import requests
import unicodedata
import pdfplumber
import pandas as pd
from wordcloud import WordCloud
from collections import Counter
from bs4 import BeautifulSoup
from config import username, password, styremote_only, intimini_only
import matplotlib.colors as clr
import pickle

G_HEADERS = {
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET',
    'Access-Control-Allow-Headers': 'Content-Type',
    'Access-Control-Max-Age': '3600',
    'User-Agent': 'Mozilla/5.0 (X11; Ubuntu; Linux x86_64; rv:52.0) Gecko/20100101 Firefox/52.0'
    }

G_LOGINURL = "https://www.timini.no/"

def soup_all_document_links() -> list:
    def recursively_get_links(session, url):
        res = session.get(url=url, headers=G_HEADERS)
        soup = BeautifulSoup(res.content, "html.parser")
        table = soup.section.div
        links = table.find_all("a", href=re.compile("document"))
        for link in links:
            linkurl = link["href"]
            document_links.append(linkurl)
        curr_page_li = soup.find("li", class_=re.compile("nolink"))
        if not curr_page_li:
            return
        next_page = curr_page_li.next_sibling()[0]
        if next_page.text != ">>":
            recursively_get_links(session, next_page["href"])

    document_links = []
    urls = ["https://www.timini.no/Aktive+medlemmer/documents", "https://www.timini.no/Aktive+medlemmer+unntatt+fadderbarn/documents", "https://www.timini.no/Alle+medlemmer/documents", "https://www.timini.no/Alle+medlemmer+unntatt+fadderbarn/documents"]
    with requests.Session() as session:
        session.post(G_LOGINURL, {"username": username, "password": password})
        for url in urls:
            recursively_get_links(session, url)
    return document_links

def document_to_download_links(document_links):
    download_links = {}
    with requests.Session() as session:
        session.post(G_LOGINURL, {"username": username, "password": password})
        for doc_url in document_links:
            res = session.get(url = doc_url, headers=G_HEADERS)
            soup = BeautifulSoup(res.content, "html.parser")
            header = soup.find("h2", class_=re.compile("documents-link"))
            filename = header.a.text
            link = header.a["href"]
            download_links[filename] = link
    return download_links

def download_all_files(download_links: dict):
    def slugify(value, allow_unicode=False):
        """
        Taken from https://github.com/django/django/blob/master/django/utils/text.py
        Convert to ASCII if 'allow_unicode' is False. Convert spaces or repeated
        dashes to single dashes. Remove characters that aren't alphanumerics,
        underscores, or hyphens. Convert to lowercase. Also strip leading and
        trailing whitespace, dashes, and underscores.
        """
        value = str(value)
        if allow_unicode:
            value = unicodedata.normalize('NFKC', value)
        else:
            value = unicodedata.normalize('NFKD', value).encode('ascii', 'ignore').decode('ascii')
        value = re.sub(r'[^\w\s\-\.]', '', value.lower())
        return re.sub(r'[\-\s]+', '-', value).strip('-_')
    with requests.Session() as session:
        session.post(G_LOGINURL, {"username": username, "password": password})
        for filename, down_url in download_links.items():
            pdfile = session.get(url = down_url, headers=G_HEADERS)
            filename = slugify(filename, allow_unicode=True)
            file_placement = document_folder/filename
            with open(file_placement, "wb") as file:
                file.write(pdfile.content)

def iterdir_to_str_list(documents_dir: pathlib.Path):
    strs = []
    for filepath in documents_dir.iterdir():
        if styremote_only:
            if not re.match("styremøte", filepath.name):
                continue # skip file
        elif intimini_only:
            if not re.match("intimini", filepath.name):
                continue

        if filepath.suffix == ".pdf":
            strs.append(pdf_to_str(filepath))
        elif filepath.suffix == ".xlsx":
            strs.append(xlsx_to_str(filepath))
        elif filepath.suffix == ".docx":
            doctxt = docx2txt.process(str(filepath))
            strs.append(doctxt)
        elif filepath.suffix == ".txt":
            with open(filepath, "r") as txtfile:
                strs.append(txtfile.read())
        elif filepath.suffix == ".doc":
            pass
        elif filepath.suffix == ".pptx":
            pass
        elif filepath.suffix == ".ppt":
            pass
    return strs

def pdf_to_str(filepath: pathlib.Path) -> str:
    pdf = filepath.resolve()
    text = ""
    with pdfplumber.open(str(pdf)) as dafile:
        for page in dafile.pages:
            text += page.extract_text()
    return text

def xlsx_to_str(filepath: pathlib.Path) -> str:
    output = ""
    df = pd.read_excel(str(filepath))
    for _, row in df.iterrows():
        for _, cell in row.iteritems():
            if str(cell) != "nan":
                output += str(cell) + " "
    return output

def add_text_to_counter(text: str, counter: Counter):
    if not text:
        return
    # text = unicodedata.normalize('NFKC', text).encode('ascii', 'ignore').decode('ascii')
    words = text.split()
    pattern = re.compile('[\W_]+')
    words = [pattern.sub("", word) for word in words]
    words = [word.upper() for word in words if len(word) > 1 and not word.isnumeric()] 
    frequencies = Counter(words)  
    counter += frequencies  

def counter_to_wordcloud(counter: Counter, filename: str):
    wordcloud = WordCloud(width=1920, height=1080, colormap=colormap, mode="RGBA", background_color=None).generate_from_frequencies(counter)
    wordcloud.to_file(filename)

def save_to_pickle(object, pickle_file):    
    with open(pickle_file, "wb") as file:
        pickle.dump(object, file)

def get_from_pickle(pickle_file):
    with open(pickle_file, "rb") as file:
        obj = pickle.load(file)
    return obj

colormap: clr.ListedColormap = clr.ListedColormap(["#df6024ff", "#f28a21ff", "#f8d164ff", "#88ba98ff", "#f0edc9ff"])

if __name__ == "__main__":
    pickle_file = pathlib.Path("./link_cache.pickle").resolve()
    document_folder = pathlib.Path("./documents/").resolve()
    results_folder = pathlib.Path("./results/").resolve()

    # Crude caching
    refresh_all = False # Override cache
    refresh_downloads = False
    if not pickle_file.is_file() or refresh_all:
        document_links: list = soup_all_document_links()
        download_links: dict = document_to_download_links(document_links)
        save_to_pickle(object=download_links, pickle_file=pickle_file)
    else:
        download_links = get_from_pickle(pickle_file)
    
    if not document_folder.exists or refresh_all or refresh_downloads:
        shutil.rmtree(document_folder, ignore_errors=True)
        os.mkdir(document_folder)
        download_all_files(download_links)

    for styremote_only, intimini_only in [(False, False), (True, False), (False, True)]:
        total_counter = Counter()

        # Henter tekst fra alle filene, her filtreres for hvilke filer vi bryr oss om
        text_list = iterdir_to_str_list(document_folder)
        
        # Legger text til counter
        for text in text_list:
            add_text_to_counter(text, total_counter) 
        
        # bestemmer filnabn
        filename = "all_timini_files"
        if styremote_only:
            filename = "styremote_files"
        elif intimini_only:
            filename = "intimini_files"

    
        # Lagrer statistikk i en fil
        with open(results_folder/(filename+".txt"), "w", encoding="utf-8") as file:
            file.write(str(total_counter))

        # Fjerner kjedelige og vanlige ord
        with open("most_common_words.txt", "r", encoding="utf-8") as file:
            for line in file:
                del total_counter[line.strip().upper()]
        
        # Lager ordsky
        counter_to_wordcloud(total_counter, str(results_folder/(filename+".png")))
