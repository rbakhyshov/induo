# -*- coding: utf-8 -*-
#import requests
#import csv, openpyxl as xl
#from PIL import Image
#from googletrans import Translator

import datetime


HEADERS = {
    "accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9",
    "user-agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.106 Safari/537.36"
}

PIC_URL = 'https://disk.yandex.ru/d/pILO6IQJQP4GQA'

def cur_time():
    return datetime.datetime.now().strftime("%H:%M:%S")

def getImage (url):
    if url == '':
        return
    try:
        response = requests.get(url, stream=True).raw
        img = Image.open(response)
        img.save('./pic_orig/' + url.split('/')[-1], 'jpeg')
        newimg = img.resize((80,120))
        newimg.save('./pic/' + url.split('/')[-1], 'jpeg')

    except:
        return

def save_xls (file_name, items_data, withImg):

    wb = xl.Workbook()
    ws = wb.worksheets[0]
    xlrow = 1

    for item in items_data:
        ws.cell(xlrow, 1, item["itemDomain"])
        ws.cell(xlrow, 2, item["itemID"])
        ws.cell(xlrow, 3, item["itemBrand"])
        ws.cell(xlrow, 4, item["itemName"])
        ws.cell(xlrow, 5, item["itemPrice"])
        if withImg == 'Y':
            try:
                xlImg = xl.drawing.image.Image('./pic/' + item["itemImageURL"].split('/')[-1])
                xlImg.anchor = 'F' + str(xlrow)
                ws.add_image(xlImg)
            except:
                pass
        ws.cell(xlrow, 7, item["itemSizesStr"])
        ws.cell(xlrow, 8, item["itemDesc"])
        ws.cell(xlrow, 9, item["itemDescRU"])
        ws.cell(xlrow, 10).value = '=HYPERLINK("' + PIC_URL + "/" + item["itemImageURL"].split('/')[-1] + '", "Фото")'
        ws.cell(xlrow, 11, item["itemURL"])
        ws.row_dimensions[xlrow].height = 90

        xlrow = xlrow + 1

    ws.column_dimensions['H'].width = 55
    ws.column_dimensions['I'].width = 55
    #ws.row_dimensions['H'].alignment.wrap_text=True
    #ws.row_dimensions['I'].alignment = Alignment (wrap_text = True)
    cur_time = datetime.datetime.now().strftime("%Y_%m_%d_%H_%M")
    if withImg == 'Y':
        file_name = f'./output/{file_name}_{cur_time}_pic.xlsx'
    elif withImg == 'N':
        file_name = f'./output/{file_name}_{cur_time}_no_pic.xlsx'

    wb.save(file_name)

def save_csv (file_name, items_data):

    cur_time = datetime.datetime.now().strftime("%Y_%m_%d_%H_%M")
    with open(f"./output/{file_name}_{cur_time}.csv", "w") as file:
        writer = csv.writer(file)

        writer.writerow(

            (
                "itemDomain",
                "itemID",
                "itemBrand",
                "itemName",
                "itemPrice",
                "itemSizesStr",
                "itemDesc",
                "itemDescRU",
                "itemImageURL",
                "itemURL"
            )
        )

    for item in items_data:
        with open(f"./output/{file_name}_{cur_time}.csv", "a") as file:
            writer = csv.writer(file)

            writer.writerow(
                (
                    item["itemDomain"],
                    item["itemID"],
                    item["itemBrand"],
                    item["itemName"],
                    item["itemPrice"],
                    item["itemSizesStr"],
                    item["itemDesc"],
                    item["itemDescRU"],
                    item["itemImageURL"],
                    item["itemURL"]
                )
            )

def save_data (file_name, items_data, output, withImg):
    if output == 'xls':
        save_xls (file_name, items_data, withImg)
    elif output == 'csv':
        save_csv(file_name, items_data)

def translate_bulk(data, lang):
    translator = Translator()

    tr_data = []
    for d in data:
        tr_data.append(d["itemDesc"])
    res = translator.translate(tr_data, dest='ru', src = lang)

    idx = 0
    for r in res:
        data[idx]["itemDescRU"] = r.text
        idx += 1
    return data

def get_domain_RU (str):

    str = str.replace ('Women', 'Женское').replace('Men', 'Мужское').replace('Unisex', 'Унисекс')
    str = str.replace('Clothing', 'Одежда').replace('Shoes', 'Обувь')
    str = str.replace ('Bags', 'Сумки').replace('Accessories', 'Аксессуары')
    return str

def make_stat (file_name):
    data = []
    print(file_name)
    wb = xl.load_workbook(f'./result/{file_name}')
    ws = wb.worksheets[0]
    xlrow = 2
    max_row = ws.max_row
    while xlrow <= max_row:
        if file_name[0] == 'c':
            boutique = "coltorti"
        elif file_name[0] == 'j':
            boutique = 'julian'
        domain = ws.cell(xlrow, 1).value
        brand = ws.cell(xlrow, 3).value
        price = ws.cell(xlrow, 5).value

        data.append({"domain": domain, "brand": brand, "price": price, "boutique": boutique})

        xlrow += 1
        if xlrow%100 == 0:
            print(xlrow)

    wb = xl.Workbook()
    ws = wb.worksheets[0]
    xlrow = 1
    for d in data:
        ws.cell(xlrow, 1, d["domain"])
        ws.cell(xlrow, 2, d["brand"])
        ws.cell(xlrow, 3, d["price"])
        ws.cell(xlrow, 4, d["boutique"])
        xlrow += 1
    wb.save('stat.xlsx')