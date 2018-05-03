import requests,time,os
import xlwings as xw
import pandas as pd
from bs4 import BeautifulSoup

wb = xw.Book.caller()
wb.sheets[0].range('L14').value = "I will start web scraping for images present in tweet URL..."

def run_web():
    filepath = os.path.dirname(os.path.abspath(__file__))
    dataset = pd.read_excel(filepath+"/Url Data.xlsx",sheet_name=0)
    image_url = dataset.iloc[:,0].values

    image_tag=[]
    i=0
    for request_url in image_url:
        page = requests.get(request_url)
        page_content = page.content
        soup = BeautifulSoup(page_content,"lxml")
        page.close()
        array=soup.findAll("img",src=True)
    
        for url_image in array:
            if "data-aria-label-part" in url_image.attrs:
                image_tag.append(url_image.attrs["src"])
                i=i+1
        time.sleep(1)
        wb.sheets[0].range('L14').value = str(i)+" image/s found."

    dataset_tweets = pd.DataFrame({"Image URLs":image_tag})
    dataset_tweets = dataset_tweets.drop_duplicates("Image URLs", keep='first')
    unique_image = dataset_tweets.iloc[:,0].values
    dataset_stats = pd.DataFrame({"No of Links":[image_url.size],"Total Images":[i],"Unique Images":[unique_image.size]},columns=["No of Links","Total Images","Unique Images"])

    filepath = os.path.dirname(os.path.abspath(__file__))
    writer = pd.ExcelWriter(filepath+"/TwitterLinkUrl.xlsx")
    dataset_tweets.to_excel(writer,sheet_name="LinkUrls",index=False)
    dataset_stats.to_excel(writer,sheet_name="Stats",index=False)
    writer.save()

    wb.sheets[0].range('L14').value = "Congratulations...!!!  Collected "+str(i)+" images, out of those "+str(unique_image.size)+" is/are unique."