import sys
import requests
from bs4 import BeautifulSoup
import json
import time 
import pandas as pd

sys.stdout.reconfigure(encoding='utf-8')

current_skip=0
data_list = []

for items in range (1,11):
    
    url = f"https://www.willamettewines.com/includes/rest_v2/plugins_listings_listings/find/?json=%7B%22filter%22%3A%7B%22%24and%22%3A%5B%7B%22filter_tags%22%3A%7B%22%24in%22%3A%5B%22site_primary_subcatid_65%22%2C%22site_primary_subcatid_47%22%2C%22site_primary_subcatid_48%22%5D%7D%7D%5D%7D%2C%22options%22%3A%7B%22limit%22%3A24%2C%22skip%22%3A{current_skip}%2C%22count%22%3Atrue%2C%22castDocs%22%3Afalse%2C%22fields%22%3A%7B%22recid%22%3A1%2C%22title%22%3A1%2C%22primary_category%22%3A1%2C%22address1%22%3A1%2C%22city%22%3A1%2C%22state%22%3A1%2C%22zip%22%3A1%2C%22url%22%3A1%2C%22isDTN%22%3A1%2C%22latitude%22%3A1%2C%22longitude%22%3A1%2C%22primary_image_url%22%3A1%2C%22qualityScore%22%3A1%2C%22rankOrder%22%3A1%2C%22weburl%22%3A1%2C%22dtn.rank%22%3A1%2C%22yelp.rating%22%3A1%2C%22yelp.url%22%3A1%2C%22yelp.review_count%22%3A1%2C%22yelp.price%22%3A1%2C%22listingudfs_object.25.value%22%3A1%7D%2C%22hooks%22%3A%5B%5D%2C%22sort%22%3A%7B%22qualityScore%22%3A-1%2C%22listingudfs_object.25.value%22%3A1%2C%22sortcompany%22%3A1%7D%7D%7D&token=924e809607c2dae2783ea441bf65cd53"
    response = requests.get(url)
    data=response.json()

    count=data['docs']['count']

    i=0

    for item in range(0,24):    
        try:
            title=(data['docs']["docs"][i]['title'])
        except KeyError as e:
            title="Not available"

        try:
            address="{} {} {} {}".format(data['docs']["docs"][i]['address1'], data['docs']["docs"][i]['city'], data['docs']["docs"][i]['state'], data['docs']["docs"][i]['zip'])
        except KeyError as e:
            address="Not available"

        try:
            web=(data['docs']["docs"][i]['weburl'])
        except KeyError as e:
            web="Not available"

        new_url="https://www.willamettewines.com/"+data['docs']["docs"][i]['url']
        response = requests.get(new_url)
        soup = BeautifulSoup(response.content, 'html.parser')

        div_tag = soup.find('div', class_='contentRender contentRender_11 contentRender_type_widget contentRender_name_plugins_listings_detail')

        
        script_tag = div_tag.find('script')
        script_content = script_tag.string    
        start = script_content.find('{') 
        json_string = script_content[start:-2]      
        json_data = json.loads(json_string)

        try:
            Email=(json_data["data"]["contact_email"])
        except KeyError as e:
            Email="Not available"
            
        try:
            Phone=(json_data["data"]["phone"])
        except KeyError as e:
            Phone="Not available"

        item = {
        "title": title,
        "address": address,
        "web": web,
        "email": Email,
        "phone": Phone,
        }

        data_list.append(item)
        i+=1
        current_skip+=1

df = pd.DataFrame(data_list)


df.to_excel('Winery.xlsx', index=False)
