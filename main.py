from selenium import webdriver
import requests
from bs4 import BeautifulSoup
import pandas as pd
import fsspec
import xlsxwriter
import os
import time
import threading


def search(card):
    def citilink(card, data, headers):
        print(card + ': Parsing citilink')
        url = link_dict[card]['citilink']
        try:
            html = requests.get(url, headers=headers, timeout=5)
        except:
            # stop program to not miss error
            print('citilink ошибка, не удалось получить данные по ссылке')
            print('Для продолжения нажмите любую клавишу')
            os.system("pause")
        soup = BeautifulSoup(html.text, 'html.parser')
        items = soup.findAll('div',
                             class_="product_data__gtm-js product_data__pageevents-js ProductCardHorizontal js--ProductCardInListing js--ProductCardInWishlist")
        for item in items:
            try:
                try:
                    name = item.find('div', class_='ProductCardHorizontal__header-block').find('a').text
                except:
                    name = ('noname')
                try:
                    link = ('https://www.citilink.ru') + item.find(
                        'div', class_='ProductCardHorizontal__header-block').find('a').get('href')
                except:
                    link = ('nolink')
                try:
                    price = item.find('span',
                                      class_='ProductCardHorizontal__price_current-price js--ProductCardHorizontal__price_current-price').text
                    price = price.replace('₽', '')
                    price = price.replace(' ', '')
                    price = price.replace('\xa0', '')
                    price = int(price)
                except:
                    price = ('noprice')
            except:
                # stop program to not miss error, same in other functions
                print('citilink ошибка, не удалось обработать данные')
                print('Для продолжения нажмите любую клавишу')
                os.system("pause")
            data.append([name, link, price])
        print(card + ': citilink ready')
        return data

    def xpert(card, data, headers):
        print(card + ': Parsing xpert')
        url = link_dict[card]['xpert']
        try:
            html = requests.get(url, headers=headers, timeout=5)
        except:
            print('xpert ошибка, не удалось получить данные по ссылке')
            print('Для продолжения нажмите любую клавишу')
            os.system("pause")
        soup = BeautifulSoup(html.text, 'html.parser')
        soup = soup.find('td', class_='content_wrapper_td padd_3')
        items = soup.findAll('tr', bgcolor='#ffffff')
        for item in items:
            try:
                try:
                    name = item.find('td', align='left').find('a').text
                except:
                    # first object not item, maybe i should add [1::] into cycle =)
                    continue
                try:
                    link = ('https://www.xpert.ru') + item.find('a').get('href')
                except:
                    link = ('nolink')
                try:
                    price = item.find('td', nowrap='nowrap').text
                    price = price.replace(' ', '')
                    price = price.replace('руб.', '')
                    price = int(price)
                except:
                    price = ('noprice')
            except:
                print('xpert ошибка, не удалось обработать данные')
                print('Для продолжения нажмите любую клавишу')
                os.system("pause")
            data.append([name, link, price])
        print(card + ': xpert ready')
        return data

    def kns(card, data, headers):
        print(card + ': Parsing kns')
        url = link_dict[card]['kns']
        try:
            html = requests.get(url, headers=headers, timeout=5)
        except:
            print('kns ошибка, не удалось получить данные по ссылке')
            print('Для продолжения нажмите любую клавишу')
            os.system("pause")
        soup = BeautifulSoup(html.text, 'html.parser')
        items = soup.findAll('div', class_='item col-6 col-md-4 col-xl-4 border bg-white')
        for item in items:
            try:
                try:
                    name = item.find('a', class_='name d-block').get('title')
                except:
                    name = ('noname')
                try:
                    link = ('https://www.kns.ru') + item.find('a', class_='name d-block').get('href')
                except:
                    link = ('nolink')
                try:
                    price = item.find('span', class_='price my-1').text
                    price = price.replace('руб.', '')
                    price = price.replace(' ', '')
                    price = price.replace('\xa0', '')
                    price = int(price)
                except:
                    price = ('noprice')
            except:
                print('kns ошибка, не удалось обработать данные')
                print('Для продолжения нажмите любую клавишу')
                os.system("pause")
            data.append([name, link, price])
        print(card + ': kns ready')
        return data

    def xcom(card, data, headers):
        print(card + ': Parsing xcom')
        url = link_dict[card]['xcom']
        try:
            html = requests.get(url, headers=headers, timeout=5)
        except:
            print('xcom ошибка, не удалось получить данные по ссылке')
            print('Для продолжения нажмите любую клавишу')
            os.system("pause")
        soup = BeautifulSoup(html.text, 'html.parser')
        items = soup.findAll('div', class_='catalog_item catalog_item--tiles')
        for item in items:
            try:
                try:
                    name = item.find('a', class_='catalog_item__name catalog_item__name--tiles').get('title')
                except:
                    name = ('noname')
                try:
                    link = ('https://www.xcom-shop.ru') + item.find(
                        'a', class_='catalog_item__name catalog_item__name--tiles').get('href')
                except:
                    link = ('nolink')
                try:
                    price = item.find('div', class_='catalog_item__new_price').text
                    price = price.replace('₽', '')
                    price = price.replace(' ', '')
                    price = price.replace('\xa0', '')
                    price = int(price)
                except:
                    price = ('noprice')
            except:
                print('xcom ошибка, не удалось обработать данные')
                print('Для продолжения нажмите любую клавишу')
                os.system("pause")
            data.append([name, link, price])
        print(card + ': xcom ready')
        return data

    def flash(card, data, headers):
        print(card + ': Parsing flash')
        url = link_dict[card]['flash']
        try:
            html = requests.get(url, headers=headers, timeout=5)
        except:
            print('flash ошибка, не удалось получить данные по ссылке')
            print('Для продолжения нажмите любую клавишу')
            os.system("pause")
        soup = BeautifulSoup(html.text, 'html.parser')
        items = soup.findAll('div', class_='listing-grid__cell')
        for item in items:
            try:
                try:
                    name = item.find('a', class_='product-card__title').text
                except:
                    name = ('noname')
                try:
                    link = ('https://flashcom.ru') + item.find('a', class_='product-card__title').get('href')
                except:
                    link = ('nolink')
                try:
                    try:
                        price = item.find('span', class_='product-card__price-current').text
                    except:
                        price = item.find('span', class_='product-card__price-new').text
                    price = price.replace('₽', '')
                    price = price.replace(' ', '')
                    price = price.replace('\xa0', '')
                    price = int(price)
                except:
                    price = ('noprice')
            except:
                print('flash ошибка, не удалось обработать данные')
                print('Для продолжения нажмите любую клавишу')
                os.system("pause")
            data.append([name, link, price])
        print(card + ': flash ready')
        return data

    def topcomp(card, data, headers):
        print(card + ': Parsing topcomp')
        url = link_dict[card]['topcomp']
        try:
            html = requests.get(url, headers=headers, timeout=5)
        except:
            print('topcomp ошибка, не удалось получить данные по ссылке')
            print('Для продолжения нажмите любую клавишу')
            os.system("pause")
        soup = BeautifulSoup(html.text, 'html.parser')
        items = soup.findAll('div', class_='col-xs-12 col-sm-4 col-md-4 col-lg-3 item')
        for item in items:
            try:
                try:
                    name = item.find('a', class_='item-name').get('title')
                except:
                    name = ('noname')
                try:
                    link = ('https://topcomputer.ru') + item.find('a', class_='item-name').get('href')
                except:
                    link = ('nolink')
                try:
                    price = item.find('span', class_='item-price').text
                    price = price.replace('₽', '')
                    price = price.replace(' ', '')
                    price = price.replace('\xa0', '')
                    price = int(price)
                except:
                    price = ('noprice')
            except:
                print('topcomp ошибка, не удалось обработать данные')
                print('Для продолжения нажмите любую клавишу')
                os.system("pause")
            data.append([name, link, price])
        print(card + ': topcomp ready')
        return data

    def regard(card, data, headers):
        print(card + ': Parsing regard')
        url = link_dict[card]['regard']
        try:
            html = requests.get(url, headers=headers, timeout=5)
        except:
            print('regard ошибка, не удалось получить данные по ссылке')
            print('Для продолжения нажмите любую клавишу')
            os.system("pause")
        soup = BeautifulSoup(html.text, 'html.parser')
        items = soup.findAll('div', class_='Card_wrap__2fsLE Card_listing__LaohM ListingRenderer_listingCard__XhvNd')
        for item in items:
            try:
                try:
                    name = item.find('h6').get('title')
                except:
                    name = ('noname')
                try:
                    link = 'https://www.regard.ru' + item.find(
                        'a', class_='CardText_link__2H3AZ link_black').get('href')
                except:
                    link = ('nolink')
                try:
                    price = item.find('span', class_='CardPrice_price__1t0QB Card_price__2Q9vg').text
                    price = price.replace('₽', '')
                    price = price.replace(' ', '')
                    price = price.replace('\xa0', '')
                    price = int(price)
                except:
                    price = ('noprice')
            except:
                print('regard ошибка, не удалось обработать данные')
                print('Для продолжения нажмите любую клавишу')
                os.system("pause")
            data.append([name, link, price])
        print(card + ': regard ready')
        return data

    def selen(card, data):
        def dns(card, html, data):
            print(card + ': Parsing dns-shop')

            soup = BeautifulSoup(html, 'html.parser')
            # in html code find a block with the cards we need
            items = soup.findAll('div', class_='catalog-product ui-button-widget')

            for item in items:
                try:
                    try:
                        link = 'https://www.dns-shop.ru' + item.find(
                            'a', class_='catalog-product__name ui-link ui-link_black').get('href')
                    except:
                        link = ('nolink')
                    try:
                        name = item.find('a', class_='catalog-product__name ui-link ui-link_black').find('span').text
                    except:
                        name = ('noname')
                    try:
                        price = item.find('div', class_='product-buy__price').text
                        price = price.replace('₽', '')
                        price = price.replace(' ', '')
                        price = price.replace('\xa0', '')
                        price = int(price)
                    except:
                        price = ('noprice')
                except:
                    # stop program to not miss error
                    print('dns-shop ошибка, не удалось обработать данные')
                    print('Для продолжения нажмите любую клавишу')
                    os.system("pause")
                data.append([name, link, price])
            print(card + ': dns-shop ready')
            return data

        def onlinetrade(card, html, data):

            print(card + ': Parsing onlinetrade')

            soup = BeautifulSoup(html, 'html.parser')

            # many blocks with similar names, specify the desired one
            cards_block = soup.find('div', class_='content__mainColumn content__itemsLoadingCover js__itemsLoadingCover')
            items = cards_block.findAll('div', class_='indexGoods__item')

            for item in items:
                try:
                    try:
                        name = item.find('img').get('alt')
                        name = name.replace('\xa0', ' ')
                    except:
                        name = ('noname')
                    try:
                        link = 'https://onlinetrade.ru' + item.find(
                            'a', class_='indexGoods__item__image js__indexGoods__item__image').get('href')
                    except:
                        link = ('nolink')
                    try:
                        price = item.find('span', class_='price regular js__actualPrice').text
                    except(AttributeError):
                        price = item.find('span', class_='price js__actualPrice').text
                    except:
                        price = ('noprice')

                    price = price.replace('₽', '')
                    price = price.replace(' ', '')
                    price = price.replace('\xa0', '')
                    price = int(price)

                except:
                    print('onlinetrade ошибка, не удалось обработать данные')
                    print('Для продолжения нажмите любую клавишу')
                    os.system("pause")
                data.append([name, link, price])
            print(card + ': onlinetrade ready')
            return data

        # settings for browser
        options = webdriver.FirefoxOptions()
        # line lower don't work in current browser version, only older version
        options.set_preference('dom.webdriver.enabled', False)
        # background mode
        # options.headless = True
        browser = webdriver.Firefox(options=options)

        browser.get(link_dict[card]['dns'])
        # sleep to let scripts work
        time.sleep(6)
        html = browser.page_source

        dns(card, html, data)
        original_window = browser.current_window_handle
        # new tab
        browser.switch_to.new_window('tab')
        browser.get(link_dict[card]['onlinetrade'])
        # sleep to let scripts work
        time.sleep(6)
        html = browser.page_source
        onlinetrade(card, html, data)

        # close tab and original window
        browser.close()
        browser.switch_to.window(original_window)
        browser.quit()

        return data

    data = []
    headers = headers = (
        {'User-Agent': 'Mozilla/5.0 (X11; Ubuntu; Linux x86_64; rv:52.0) Gecko/20100101 Firefox/103.0.1'}
    )

    citilink(card, data, headers)
    xpert(card, data, headers)
    kns(card, data, headers)
    xcom(card, data, headers)
    flash(card, data, headers)
    topcomp(card, data, headers)
    regard(card, data, headers)

    # use selenium for sites that are protected from parsing with requests
    print('Обработка dns и onlinetrade займёт некоторое время, это нормально')
    selen(card, data)

    if (card == '3060'):
        global data_3060
        data_3060 = tuple(data)
    elif (card == '6700xt'):
        global data_6700xt
        data_6700xt = tuple(data)
    print(card + ': Ready')


# dict for sites and links
link_dict = ({
    '3060': {
        'citilink': 'https://www.citilink.ru/catalog/videokarty/?pf=discount.any%2Crating.any%2C9368_29nvidiad1d1geforced1rtxd13060&f=discount.any%2Crating.any%2C9368_29nvidiad1d1geforced1rtxd13060%2Cavailable.all',
        'xpert': 'https://www.xpert.ru/products.php?category_id=1042',
        'kns': 'https://www.kns.ru/multi/catalog/komplektuyuschie/videokarty/_graficheskij-protsessor_nvidia-geforce-rtx-3060/_v-nalichii/',
        'xcom': 'https://www.xcom-shop.ru/catalog/kompyuternye_komplektyyuschie/videokarty/videokarta/?prop_7215%5Bvalue%5D=58080&o=p',
        'flash': 'https://flashcom.ru/market/videokartyi/filter_video_cpu-NVIDIA%20GeForce%20RTX%203060/',
        'topcomp': 'https://topcomputer.ru/katalog/290-videokarty/?arrFilter_P1_MIN=2318&arrFilter_P1_MAX=771650&arrFilter_944_3002761195=Y&set_filter=Y',
        'regard': 'https://www.regard.ru/catalog/1013/videokarty?q=eyJzb3J0IjpbIm9yZGVyQnlQcmljZSIsImFzYyJdLCJieUNoYXIiOnsiNjEiOnsiZXhjZXB0IjpmYWxzZSwidmFsdWVzIjpbMjEwXX0sIjY3Ijp7InZhbHVlcyI6WzI0NzQ2XSwiZXhjZXB0IjpmYWxzZX19fQ',
        'dns': 'https://www.dns-shop.ru/catalog/17a89aab16404e77/videokarty/?f[mv]=zyyhm',
        'onlinetrade': 'https://www.onlinetrade.ru/catalogue/videokarty-c338/?selling[]=1&price1=1999&price2=787700&graphic_processor[]=NVIDIA%20GeForce%20RTX%203060&advanced_search=1&rating_active=0&special_active=1&selling_active=1&producer_active=1&price_active=0&proizvoditel_vid_active=0&line_active=1&graphic_processor_active=1&naznachenie_active=1&memory_size_active=1&memory_type_active=1&bus_active=0&dop_pitanie_active=1&cooling_mode_active=1&ventilyatori_active=1&rekom_power_active=1&podsvetka_active=0&dlina_active=0&low_profile_active=0&kol_slots_active=0&sockets_active=0&srok_garantii_active=0&cat_id=338&per_page=30&page=0&sort=price-asc&page=0'
    },
    '6700xt': {
        'citilink': 'https://www.citilink.ru/catalog/videokarty/?pf=discount.any%2Crating.any%2C9368_29amdd1d1radeond1rxd16700xt&f=discount.any%2Crating.any%2C9368_29amdd1d1radeond1rxd16700xt%2Cavailable.all&sorting=price_asc',
        'xpert': 'https://www.xpert.ru/products.php?category_id=1047',
        'kns': 'https://www.kns.ru/multi/catalog/komplektuyuschie/videokarty/_graficheskij-protsessor_rx-6700-xt/_v-nalichii/',
        'xcom': 'https://www.xcom-shop.ru/catalog/kompyuternye_komplektyyuschie/videokarty/videokarta/?prop_7215%5Bvalue%5D=58188&o=p',
        'flash': 'https://flashcom.ru/market/videokartyi/filter_video_cpu-AMD%20Radeon%20RX%206700%20XT/',
        'topcomp': 'https://topcomputer.ru/katalog/290-videokarty/?sort=price&order=asc&arrFilter_P1_MIN=2204&arrFilter_P1_MAX=733700&arrFilter_944_651289571=Y&set_filter=Y',
        'regard': 'https://www.regard.ru/catalog/1013/videokarty?q=eyJzb3J0IjpbIm9yZGVyQnlQcmljZSIsImFzYyJdLCJieUNoYXIiOnsiNjEiOnsidmFsdWVzIjpbMjA5XSwiZXhjZXB0IjpmYWxzZX0sIjY3Ijp7InZhbHVlcyI6WzI0Nzg4XSwiZXhjZXB0IjpmYWxzZX19fQ',
        'dns': 'https://www.dns-shop.ru/catalog/17a89aab16404e77/videokarty/?f[mv]=11xign',
        'onlinetrade': 'https://www.onlinetrade.ru/catalogue/videokarty-c338/?selling[]=1&price1=1999&price2=970700&graphic_processor[]=AMD%20Radeon%20RX%206700%20XT&advanced_search=1&rating_active=0&special_active=1&selling_active=1&producer_active=1&price_active=0&proizvoditel_vid_active=0&line_active=1&graphic_processor_active=1&naznachenie_active=1&memory_size_active=1&memory_type_active=1&bus_active=0&dop_pitanie_active=1&cooling_mode_active=1&ventilyatori_active=1&rekom_power_active=1&podsvetka_active=0&dlina_active=0&low_profile_active=0&kol_slots_active=0&sockets_active=0&srok_garantii_active=0&cat_id=338'
    }
})

start_time = time.time()

# for each card in link_dict, we give own thread
threads = []
for card in link_dict:
    thr = threading.Thread(target=search, args=(card,))
    # writing threads that we started
    threads.append(thr)
    thr.start()

# wait until all threads complete their work
for thr in threads:
    thr.join()

columns = ['name', 'link', 'price']

df_data_3060 = pd.DataFrame(data_3060, columns=columns)
df_data_6700xt = pd.DataFrame(data_6700xt, columns=columns)

# path for saving excel file
path = os.environ
path = path['USERPROFILE'] + '\Desktop\\testMulti11.xlsx'

with pd.ExcelWriter(path) as writer:
    df_data_3060.to_excel(writer, sheet_name='3060', index=False)
    df_data_6700xt.to_excel(writer, sheet_name='6700xt', index=False)

# calculate working time
finish = time.time()
print(finish - start_time)

print('The End')


