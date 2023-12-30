import os
from datetime import datetime
from dateutil.relativedelta import relativedelta
import json
import xlwings as xw
import shutil
import time
from playwright.sync_api import sync_playwright
import pandas as pd

current_dir = os.getcwd()

current_date = datetime.now()
month_ago_date = current_date - relativedelta(months=1)
two_week_ago_date = current_date - relativedelta(weeks=2)
week_ago_date = current_date - relativedelta(weeks=1)

 # Форматирование даты в нужном формате (год, месяц, день)
formatted_date = '{dt.year}_{dt.month}_{dt.day}'.format(dt = current_date)

last_month_date = month_ago_date.strftime("%Y-%m-%d")
last_two_week_date = two_week_ago_date.strftime("%Y-%m-%d")
last_week_date = week_ago_date.strftime("%Y-%m-%d")

# Создание имени файла с использованием текущей даты
file_name_fbo = "stocks_and_movement_products-report.xlsx"

dates = {last_week_date: 'supplier-goods-week',
          last_two_week_date: 'supplier-goods-two-weeks',
            last_month_date: 'supplier-goods-month'}

file_names = {
    'неделя': 'supplier-goods-week',
    '2 недели': 'supplier-goods-two-weeks',
    'месяц': 'supplier-goods-month'
    }   

warehouses = {
    'Москва, МО и Дальние регионы': ['ЖУКОВСКИЙ_РФЦ', 'ПЕТРОВСКОЕ_РФЦ', 'ПУШКИНО_1_РФЦ', 'ПУШКИНО_2_РФЦ', 'ХОРУГВИНО_РФЦ'],
    'Центр': ['ТВЕРЬ_РФЦ'],
    'Юг': ['АДЫГЕЙСК_РФЦ', 'ВОРОНЕЖ_МРФЦ', 'НОВОРОССИЙСК_МРФЦ', 'Ростов_на_Дону_РФЦ'],
    'Урал': ['Екатеринбург_РФЦ_НОВЫЙ'],
    'Поволжье': ['Казань_РФЦ_НОВЫЙ', 'НИЖНИЙ_НОВГОРОД_РФЦ', 'САМАРА_РФЦ'],
    'Сибирь': ['КРАСНОЯРСК_МРФЦ', 'Новосибирск_РФЦ_НОВЫЙ'],
    'Санкт-Петербург и СЗО': ['СПБ_БУГРЫ_РФЦ', 'Санкт_Петербург_РФЦ', 'СПБ_ШУШАРЫ_РФЦ']
    }

multi_index = [
    ('', 'Артикул')
]

def save_cookies(context):
    with open("cookies.json", "w") as f:
        f.write(json.dumps(context.cookies()))


def create_multi_index() -> list:
    multi_index.extend([('', f'Количество заказанного товара {timeframe}') for timeframe in file_names.keys()])
    multi_index.append(('', 'Итого по складам'))
    for region in warehouses.keys():
        multi_index.append((region, f'Остатки {region}'))
        multi_index.extend([(region, f'Количество заказанного товара {timeframe}') for timeframe in file_names.keys()])
        multi_index.extend([(region, f'{region} поставка {timeframe}') for timeframe in file_names.keys()])
    multi_index.extend([('Мой склад', f'Количество заказанного товара {timeframe}') for timeframe in file_names.keys()])
    return multi_index


def download_fbo_data(page):
    page.goto('https://seller.ozon.ru/app/analytics/fulfillment-reports/stocks-and-movement-products-to-ozon-warehouses')
    
    with page.expect_download() as download_info:
        page.locator('xpath=//span[text()="Скачать"]').click()
        download = download_info.value
        dir_name = current_dir + r'\tables' + rf'\{download.suggested_filename}'
        # Сохранение файла
        download.save_as(dir_name)
    wingsbook = xw.Book(dir_name)
    wingsapp = xw.apps.active
    wingsbook.save(dir_name)
    wingsapp.quit()
    fbo_df = pd.read_excel(dir_name, skiprows=3, usecols=['Артикул','Доступный к продаже товар', 'Название склада'])

    table = fbo_df.pivot_table(index='Артикул', columns='Название склада', values='Доступный к продаже товар', aggfunc='sum', fill_value=0)
    table['Итого по складам'] = table.loc[:, table.columns != "Артикул продавца"].sum(axis=1)

    table.to_excel(dir_name)
    
def dowload_file(page, file_name):
    try:
        with page.expect_download() as download_info:
            page.locator('xpath=//span[text()="Скачать"]').click()
            download = download_info.value
            dir_name = current_dir + r'\tables' + rf'\{file_name}.csv'
            # Сохранение файла
            download.save_as(dir_name)
            return True
    except:
        return False
def download_orders_data(page, url, prefix='') -> None:
    for date, file_name in dates.items():
        file_name = file_name+prefix
        data = {"subFilter":{"period":"customDates","dateRange":{"dateFrom":int(datetime.strptime(date, "%Y-%m-%d").timestamp()),"dateTo":int(current_date.timestamp())}}}
        formatted_data = str(data).replace("'", '"')
        page.goto(url+formatted_data)

        time.sleep(1)        
        page.reload()
        time.sleep(2)        
        for _ in range(1, 6):
            if dowload_file(page, file_name):
                break
            else:
                time.sleep(1)
                page.goto(url+formatted_data)
        else:
            print(f'Не удалось скачать файл {file_name}') 

        order_df = pd.read_csv(current_dir + r'\tables' + rf'\{file_name}.csv', usecols=['Артикул', 'Кластер отгрузки'], sep=';')
    
        
        # Агрегирование данных: подсчет количества элементов для каждой комбинации "Артикул продавца" и "Склад"
        aggregated_data = order_df.groupby(['Артикул', 'Кластер отгрузки']).size().reset_index(name='шт.')
        
        aggregated_data.columns = ['Артикул', 'Склад', 'шт..1']
        # Вывод таблицы
        aggregated_data.to_excel(current_dir+r'\tables'+rf'\{file_name}.xlsx', index=False)


def table_collection():
    main_df = pd.read_excel('Планирование поставок.xlsx', skiprows=1)
    fbo_df = pd.read_excel(current_dir+r'\tables'+rf'\{file_name_fbo}', engine='openpyxl')
    main_df['Артикул'] = main_df['Артикул'].astype(str)
    fbo_df['Артикул'] = fbo_df['Артикул'].astype(str)

    def _calculate(row, timeframe:str, region:str):
        if row[f'Остатки {region}'] < row[f'Количество заказанного товара {timeframe} {region}']:
            return row[f'Количество заказанного товара {timeframe} {region}'] 
        return 0 

    main_df['Итого по складам'] = main_df['Артикул'].map(fbo_df.set_index('Артикул')['Итого по складам'])

    # Обработка данных и суммирование остатков
    for region, warehouse_list in warehouses.items():
        fbo_df[f'Остатки {region}'] = fbo_df[warehouse_list].sum(axis=1)
        main_df[f'Остатки {region}'] = main_df['Артикул'].map(fbo_df.set_index('Артикул')[f'Остатки {region}'])

    for timeframe, file_name in file_names.items():
        order_df = pd.read_excel(current_dir+r'\tables'+rf'\{file_name}-my-warehouse.xlsx')
        order_df['Артикул'] = order_df['Артикул'].astype(str)
        sum_df = order_df.groupby('Артикул')['шт..1'].sum().reset_index()
        main_df[f'Количество заказанного товара {timeframe} Мой склад'] = main_df['Артикул'].map(sum_df.set_index('Артикул')['шт..1'])
        main_df[f'Количество заказанного товара {timeframe} Мой склад'].fillna(0, inplace=True)

        order_df = pd.read_excel(current_dir+r'\tables'+rf'\{file_name}.xlsx')
        order_df['Артикул'] = order_df['Артикул'].astype(str)
        sum_df = order_df.groupby('Артикул')['шт..1'].sum().reset_index()
        main_df[f'Количество заказанного товара {timeframe}'] = main_df['Артикул'].map(sum_df.set_index('Артикул')['шт..1'])
        
        for region in warehouses.keys():
            filtered_df = order_df[order_df['Склад']==region]
            sum_df = filtered_df.groupby('Артикул')['шт..1'].sum().reset_index()
            main_df[f'Количество заказанного товара {timeframe} {region}'].fillna(0, inplace=True)
            main_df[f'Количество заказанного товара {timeframe} {region}'] =  main_df['Артикул'].map(sum_df.set_index('Артикул')['шт..1'])

            main_df[f'{region} поставка {timeframe}'] = main_df.apply(_calculate, args=(timeframe, region), axis=1)

    main_df.fillna(0, inplace=True)
    main_df.columns = pd.MultiIndex.from_tuples(create_multi_index(), names=['Регион', 'Поле'])
    main_df.to_excel('Планирование поставок итоговый.xlsx')
    

def main():   
    try: 
        if not os.path.exists('tables'):
            os.mkdir('tables')
        else:
            shutil.rmtree('tables')
            os.mkdir('tables')
            
        with sync_playwright() as p:
            browser = p.chromium.launch(
                headless=False,
                args=[
                    '--ignore-certificate-errors',
                    '--disable-blink-features=AutomationControlled'
                ],  
            )

            context = browser.new_context()
            
            page = context.new_page()
            page.set_viewport_size({'width': 1920, 'height': 1080})
        
            page.goto('https://seller.ozon.ru/app/dashboard/main')
            page.click('text=Войти')
            time.sleep(3)
            
            # Загрузка cookies, если они есть
            if os.path.exists('cookies.json'):
                with open('cookies.json', 'r') as f:
                    cookies = json.loads(f.read())
                for cookie in cookies:
                    context.add_cookies([cookie])
            else:
                while True:
                    try:
                        page.wait_for_selector('.index_tooltip_3lnsW', timeout=2000)
                        time.sleep(13)
                        break
                    except:
                        pass
                save_cookies(context)
        
            download_fbo_data(page)
            print('FBO скачан')
            
            download_orders_data(page, 'https://seller.ozon.ru/app/analytics/fulfillment-reports/operation-orders-fbo?filter=')
            download_orders_data(page, 'https://seller.ozon.ru/app/analytics/fulfillment-reports/operation-orders-fbs?filter=', prefix='-my-warehouse')
            print('Заказы скачаны')

            table_collection()  
        
    except Exception as ex:
        print('Ошибка:',ex)
    finally:        
        print('Скрипт завершён')    
        input('Нажмите Enter для выхода\n')

    
if __name__ == "__main__":
    main()
    