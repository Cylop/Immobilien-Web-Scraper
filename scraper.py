import time
import datetime
import requests
from bs4 import BeautifulSoup

import xlsxwriter

class Scraper:

    site_url = "https://www.immobilienscout24.at"
    flat_list_elem, flat_list_class = ("ol", "_7kVQE") #HTML Element der Wohnungsliste, Class der Wohnungliste
    flat_list_elem_tag, flat_list_elem_class = ("li", "_13KwO") #List Items HTML Element & Class

    address_html_tag, address_html_class = ("address", "_3_-0F") #Addresse HTML Tag, Class
    inserat_link_tag, inserat_link_class = ("a", "_3h-3Y") #Link zum Inserat

    inserat_headline_html, inserat_headline_class = ("h2", "_1Fx4t") #Headline des Inserates

    tags_list_html, tags_list_class = ("ul", "_6PgMT") #Tags des Inserates
    tags_li_html, tags_li_class = ("li", "_1Wcg4") #Einzelne Tags

    data_html, data_class = ("dd", "_2Lq0R") #Zimmer
    allowed_data = ["Zimmer", "Fläche"]
    price_classes = "_2Lq0R ixii3"

    url_curr_site = ""
    curr_site = 0

    def __init__(self, delay, base_url):
        self.delay = delay
        self.url = base_url

    def start_scraping(self):
        self.curr_site +=1
        if self.curr_site > 1:
            self.url_curr_site = "seite-" + str(self.curr_site)

            self.flat_list_elem_class = "_2Ozl0"
            self.inserat_link_class = "_3bbKs"
            self.address_html_class = "zKKpm"
            self.inserat_headline_class = "_2l9wB"

            print("Url: " + self.url + self.url_curr_site)
        if self.curr_site >= 100:
            return False
        
        response = requests.get(url=self.url + self.url_curr_site) #Response
        print("Status: " + str(response.status_code))
        if response.status_code == 200: 
            soup = BeautifulSoup(response.content, 'html.parser')

            field_flat_list = soup.find(self.flat_list_elem, self.flat_list_class) #Liste der Wohnungen auf der Seite
            print("Count: ", len(field_flat_list))
            flat_items = field_flat_list.find_all(self.flat_list_elem_tag, self.flat_list_elem_class) #Wohnungs Items in der Liste

            result_list = []
            for flat_item in flat_items: #Single Inserate
                headline = flat_item.find(self.inserat_headline_html, self.inserat_headline_class).text # Headline
                address = flat_item.find(self.address_html_tag, self.address_html_class).text #Selektierte Addresse
                link = self.site_url + flat_item.find(self.inserat_link_tag, self.inserat_link_class)['href'] #Link des Inserates
                tags = []
                for tag in flat_item.find_all(self.tags_li_html, self.tags_li_class):
                    tags.append(tag.text) #Tags hinzufügen

                data_items_li = flat_item.find_all(self.data_html, self.data_class) #Daten wie Preis, Fläche, Zimmer
                data_items = {}
                for data_item in data_items_li:
                    child_span = data_item.find("span", "_2IRwG") #Span für die Bezeichnung
                    if child_span and child_span.text in self.allowed_data:
                        data_items[child_span.text] = data_item.find(text=True, recursive=False)
                
                data_items["Preis"] = flat_item.find_all(self.data_html, self.price_classes)[0].text

                print("Headline: ", headline)
                print("Address: ", address)
                print("Data: ", data_items)
                print("Link: ", link)

                result_list.append(
                    [
                        headline,
                        address,
                        data_items,
                        link
                    ]
                )

            return result_list
        
        return False

base_url = "https://www.immobilienscout24.at/regional/oesterreich/wohnung-mieten/" #Request URL

scraper = Scraper(2000, base_url)

result = []
succ = True
while succ:
    curr_result = scraper.start_scraping()

    if curr_result == False or len(curr_result) == 0:
        succ = False
        break
    else:
        result += curr_result

    time.sleep(0.5)

print("Ende der Suche erreicht!")
print("Schreibe Daten nun in das Excel File!")


workbook = xlsxwriter.Workbook("Wohnungen_" + str(datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")) + "_export.xlsx")
worksheet = workbook.add_worksheet()

worksheet.write(0, 0, "#")
worksheet.write(0, 1, "Überschrift")
worksheet.write(0, 2, "Adresse")
worksheet.write(0, 3, "Zimmer")
worksheet.write(0, 4, "Fläche")
worksheet.write(0, 5, "Preis")
worksheet.write(0, 6, "Link")

row = 1
for headline, address, data_items, link in (result):
    worksheet.write(row, 0, row)
    worksheet.write(row, 1, headline)
    worksheet.write(row, 2, address)
    if "Zimmer" in data_items:
        worksheet.write(row, 3, data_items["Zimmer"])
    if "Fläche" in data_items:
        worksheet.write(row, 4, data_items["Fläche"])
    if "Preis" in data_items:
        worksheet.write(row, 5, data_items["Preis"])
    worksheet.write(row, 6, link)

    row += 1     


workbook.close()
                

