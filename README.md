from bs4 import BeautifulSoup   
import requests
import openpyxl
import parsel
import traceback
import os
from datetime import datetime
import re
from selenium import webdriver
import time
from selenium.webdriver.common.by import By
excel_file_path = ("loctiteproducts_output.xlsx")
workbook = openpyxl.load_workbook(excel_file_path)
sheet = workbook['Sheet']
excel = openpyxl.Workbook()
new_sheet = excel.active
sheet.title = "January"
new_sheet.append(['URL','SKUS','Title','Description' ,'Feature','Meta Title','Meta Descriptions','Meta Keywords'])
image = excel.create_sheet("Image")
image.append(['URL','SKUS', 'Image Link', 'Image Name'])
pdf = excel.create_sheet("PDF")
pdf.append(['URL','SKUS','PDF Link 1','PDF Name 1',	'PDF Link 2',	'PDF Name 2',	'PDF Link 3',	'PDF Name 3',	'PDF Link 4',	'PDF Name 4',	'PDF Link 5'])
video_sheet = excel.create_sheet("Video")
video_sheet.append(['URL','Taxonomy','End Taxonomy', 'Title', 'SKUS','Attribute Name 1','Attribute Value 1'])
row = 2
row_su = 2
row_sus =2
row_incs = 2
row_inc = 2
add_spec_rows = 2
max_row = sheet.max_row
start_time = datetime.now()
current_date = start_time.strftime('%d-%m-%y')
row_suks = 2
driver = webdriver.Chrome()
inupt_files = os.getcwd() +"\\HTML Files\\"
if not os.path.isdir(inupt_files):
    os.mkdir(inupt_files)
def file_name_checker(name):
    for char in ['@','$','%','&','\\','/',':','*','?','"',"'",'<','>','|','~','`','#','^','+','=','{','}','[',']',';','!']:
        if char in name:
            name = name.replace(char, "__")
        return name
    
try:
    for index,row in enumerate(sheet.iter_rows(min_row=2, max_row = max_row, values_only=True)):
        print(f"-------- Processing row {index+1} out of {max_row} rows -------------")
        url = row[5]
        print(url)
        response = requests.get(url)
        driver.get(url)
        y = 100
        for timer in range(0,50):
            driver.execute_script("window.scrollTo(0, "+str(y)+")")
            y += 100
            time.sleep(0.10)
        time.sleep(2)
        soup_selenium = BeautifulSoup(driver.page_source)
        soup = BeautifulSoup(response.content,"html.parser")
        print(f"****************************************{response}*********************************************")
        x_path = parsel.Selector(response.text)
        tox_cate =" | ".join([i for i in x_path.xpath('//nav[@class="breadcrumb"]//ol//li//a//span[@class="breadcrumb__linkTitle breadcrumb__link-text"]//text()').getall() if i.strip()!=''])
        features =x_path.xpath('//div[@class="product__benefits"]//ul/li//text()').getall()
        end_level = tox_cate.split("|")[-1].strip()
        title= " ".join([i.strip() for i in x_path.xpath("//div[@class='product__title']//h1//text()").getall() if i != ''])
        print(title)
        skus = title.split(",")[0].strip()
        if skus  != "":
            try:
                with open(f"{inupt_files}File__{file_name_checker(str(skus ))}_{current_date}.html", "w+",encoding='utf-8') as f:
                    f.write(str(soup))
                    f.close()
            except:
                script_error = f"Script Error \n {str(traceback.format_exc())}"
        else:
            try:
                with open(f"{inupt_files}File__{file_name_checker(str(index))}_{current_date}.html", "w+",encoding='utf-8') as f:
                    f.write(str(soup))
                    f.close()
            except:
                script_error = f"Script Error \n {str(traceback.format_exc())}"
        description = "\n".join([i.strip() for i in x_path.xpath("//div[@class='text__base']//p//text()").getall() if i.strip() != ''])       
        meta_tltle = " ".join([i.strip() for i in x_path.xpath("//meta[@property='og:title']//@content").getall() if i.strip() != ''])
        meta_descriptions = " ".join([i.strip() for i in x_path.xpath("//meta[@property='og:description']//@content").getall() if i.strip() != '']) 
        meta_keywords = " ".join([i.strip() for i in x_path.xpath("//meta[@name='keywords']//@content").getall() if i.strip() != ''])
        oz_variant = x_path.xpath('//div[@class="product__packageSizesList"]//a')
        for oz_loop in oz_variant:
            var_oz = oz_loop.xpath(".//p//text()").get()
            new_sheet.cell(row_su,1).value = url
            new_sheet.cell(row_su,2).value = tox_cate
            new_sheet.cell(row_su,3).value = end_level
            new_sheet.cell(row_su,4).value = title
            new_sheet.cell(row_su,5).value = var_oz
            new_sheet.cell(row_su,6).value = description
            new_sheet.cell(row_su,7).value = meta_tltle
            new_sheet.cell(row_su,8).value = meta_descriptions
            new_sheet.cell(row_su,9).value = meta_keywords
            colms = 10
            for i_f in features:
                print(i_f)
                features_loop = i_f
                new_sheet.cell(row_su,colms).value = features_loop 
                colms+=1
            row_su +=1
        oz_variant = x_path.xpath('//div[@class="product__packageSizesList"]//a')
        for oz_loop in oz_variant:
            var_oz = oz_loop.xpath(".//p//text()").get()
            img = x_path.xpath('//picture[@class="image__picture "]/img[@class="image__img gallery__thumbnail-image"]')
            columns =3
            if img != []:
                for index , images_url in enumerate(img):
                    each_img = f"{images_url.xpath('.//@src').get('')}?wid=1600&fit=fit%2C1&qlt=90&align=0%2C0&hei=1600"
                    im_name = f'{title}_{index+1}.jpg'
                    image.cell(row = row_incs,column = 1).value = url.strip()
                    image.cell(row = row_incs,column = 2).value = title.strip()
                    image.cell(row = row_incs,column = 3).value = var_oz
                    image.cell(row = row_incs, column = columns).value = each_img
                    image.cell(row = row_incs, column = columns+1).value = im_name.strip()
                    columns +=2
                row_incs += 1 
        # header_list = [i.strip() for i in x_path.css('.parametric-table.generic-product tr.header-row.labels-row > th.header-col div.th-wrapper  div > span::text').getall()]
        oz_variant = x_path.xpath('//div[@class="product__packageSizesList"]//a')
        for oz_loop in oz_variant:
            var_oz = oz_loop.xpath(".//p//text()").get()
            pdf_link = x_path.css(".product__documents a")
            col_num = 3
            if pdf_link!= []:
                for index ,product_pdf_link in enumerate(pdf_link):    
                    product_pdf_links =product_pdf_link.xpath(".//@href").get()
                    try:
                        product_pdf_links1 =product_pdf_link.xpath(".//@data-adobe-analytics").jmespath("name").get()
                    except:
                        product_pdf_links1s =product_pdf_link.xpath(".//@data-adobe-analytics").get()
                        product_pdf_links1 = re.findall(r'name\": \"(.*?)\"}', product_pdf_links1)
                        print(product_pdf_links1)
                        print(product_pdf_links1s)
                    pdf_n = f'{str(product_pdf_links1)}_{index+1}.pdf'
                    pdf.cell(row = row_inc,column = 1).value = url
                    pdf.cell(row = row_inc,column = 2).value = title.strip()
                    pdf.cell(row = row_inc,column = 3).value = var_oz 
                    pdf.cell(row = row_inc, column = col_num).value = product_pdf_links 
                    pdf.cell(row = row_inc, column = col_num+1).value = pdf_n
                    col_num += 2
                row_inc += 1
        oz_variant = x_path.xpath('//div[@class="product__packageSizesList"]//a')
        for oz_loop in oz_variant:
            var_oz = oz_loop.xpath(".//p//text()").get()
            video_link = soup_selenium.select('video.s7videoelement source')
            col = 4
            if video_link != []:
                for index, video_url in enumerate(video_link):
                    video_urls = video_url['src']
                    vi_link = f'https://www.loctiteproducts.com{video_urls}'
                    print(video_urls)
                    v = f'{title}_{index+1}.mp4'
                    video_sheet.cell(row = row_suks,column = 1).value = url
                    video_sheet.cell(row =row_suks,column = 2).value = title.strip()
                    video_sheet.cell(row =row_suks,column = 3).value = var_oz
                    video_sheet.cell(row =row_suks,column =col).value = v.strip()
                    video_sheet.cell(row =row_suks,column = col+1).value =vi_link
                    col +=2
                row_suks +=1
        excel.save('loctiteproducts_data.xlsx')
except:
    print(str(traceback.format_exc()))  

try:
    excel.save('loctiteproducts_data.xlsx')
except:
    input("Please close the output_data file")
    excel.save('loctiteproducts_data.xlsx')
    
    # python-code
learning
