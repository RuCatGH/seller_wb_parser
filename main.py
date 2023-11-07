import os
from datetime import datetime
from dateutil.relativedelta import relativedelta
import json
import shutil

from playwright.sync_api import sync_playwright
import pandas as pd

current_date = datetime.now()
new_date = current_date - relativedelta(months=1)

 # Форматирование даты в нужном формате (год, месяц, день)
formatted_date = '{dt.year}_{dt.month}_{dt.day}'.format(dt = current_date)

last_month_date = new_date.strftime("%d.%m.%Y")

# Создание имени файла с использованием текущей даты
file_name_fbo = f"report_{formatted_date}.xlsx"

dates = {'Последняя неделя': 'supplier-goods-week.xlsx',
          'Последние две недели': 'supplier-goods-two-weeks.xlsx',
            'Текущий месяц': 'supplier-goods-month.xlsx'}

file_names = {
    'неделя': 'supplier-goods-week.xlsx',
    '2 недели': 'supplier-goods-two-weeks.xlsx',
    'месяц': 'supplier-goods-month.xlsx'
    }   

warehouses = {
    'Центральный': ['Тула', 'Подольск', 'Электросталь', 'Коледино', 'Белые Столбы'],
    'Северо-Западный': ['Санкт-Петербург'],
    'Южный': ['Краснодар', 'Невинномысск'],
    'Уральский': ['Екатеринбург'],
    'Приволжский': ['Казань'],
    'Сибирский': ['Новосибирск']
    }

multi_index = [
    ('', 'Артикул поставщика'),
    ('', 'Артикул продавца'),
]

# Функция для сохранения cookies в файл
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
    return multi_index


def download_fbo_data(page):
    # Скачивание остатков по модели FBO
    page.goto("https://seller.wildberries.ru/analytics/warehouse-remains")

    settings_button = page.locator('xpath=//span[text()="Настройка таблицы"]')
    settings_button.click()

    # Выбираем галочку
    overlay = page.locator('[data-name="Overlay"]').first
    overlay.click()

    checkbox = page.locator('[class^="checkboxContainer_"]').all()[2]
    checkbox.click()
    page.wait_for_timeout(1000)
    
    save_button = page.locator('xpath=//span[text()="Сохранить"]')
    save_button.click()
    page.wait_for_timeout(1000)

    with page.expect_download(timeout=25000) as download_info:
        load_button = page.locator('xpath=//span[text()="Выгрузить в Excel"]')
        load_button.click()

    download = download_info.value
        
    # Сохранение файла
    download.save_as(os.getcwd()+r'\tables'+ rf'\{download.suggested_filename}')
    

    page.wait_for_timeout(1000)

    

def download_orders_data(page) -> None:
    for date in dates.keys():
        page.goto('https://seller.wildberries.ru/analytics/sales')
        page.reload()
        page.wait_for_selector('.OrdersTableView')
        page.wait_for_timeout(3000)
        try:
            delete_buttons = page.locator('xpath=//*[starts-with(@class, "DeleteButton")]//button[@type="button"]').all()
            for delete_button in delete_buttons:
                delete_button.click()
                page.wait_for_timeout(3000)
        except Exception as _:
            pass

        date_filter = page.locator('[data-name="SimpleInput"]').locator('xpath=following-sibling::button')
        date_filter.click()

        page.wait_for_timeout(100)
        if date != 'Текущий месяц':
            last_week = page.locator(f'xpath=//span[text()="{date}"]')
            last_week.click()
        else:
            input_start = page.locator('[id="startDate"]')
            input_start.fill(last_month_date)
            page.wait_for_timeout(1000)
            input_end = page.locator('[id="endDate"]')
            input_end.fill(current_date.strftime("%d.%m.%Y"))
        page.wait_for_timeout(200)

        save = page.locator('xpath=//span[text()="Сохранить"]')
        save.click()
        page.wait_for_timeout(1000)

        page.reload()

        page.wait_for_timeout(1000)
        
        with page.expect_download(timeout=25000) as download_info:
            save_xlsx = page.locator('xpath=//span[text()="Выгрузить в Excel"]')
            save_xlsx.click()

        download = download_info.value
            
        # Сохранение файла
        download.save_as(os.getcwd()+r'\tables'+rf'\{download.suggested_filename}')

        page.wait_for_timeout(2000)
        
        rename_xlsx_file(date)

def rename_xlsx_file(date: str) -> None:
    files = os.listdir(os.getcwd()+r'\tables')
    for file_name in files:
        if file_name.startswith("supplier-goods-26475"):
            # Переименовываем файл
            new_file_name = dates[date]
            os.rename(os.getcwd()+r'\tables'+fr'\{file_name}', os.getcwd()+r'\tables'+fr'\{new_file_name}')
            break

def table_collection():
    main_df = pd.read_excel('Планирование поставок.xlsx', skiprows=1)
    fbo_df = pd.read_excel(os.getcwd()+r'\tables'+rf'\{file_name_fbo}')
    main_df['Артикул продавца'] = main_df['Артикул продавца'].astype(str)
    fbo_df['Артикул продавца'] = fbo_df['Артикул продавца'].astype(str)

    def _calculate(row, timeframe:str, region:str):
        if row[f'Остатки {region}'] < row[f'Количество заказанного товара {timeframe} {region}']:
            return row[f'Количество заказанного товара {timeframe} {region}']
        else:
            return 0

    main_df['Итого по складам'] = main_df['Артикул продавца'].map(fbo_df.set_index('Артикул продавца')['Итого по складам'])

    # Обработка данных и суммирование остатков
    for region, warehouse_list in warehouses.items():
        fbo_df[f'Остатки {region}'] = fbo_df[warehouse_list].sum(axis=1)
        main_df[f'Остатки {region}'] = main_df['Артикул продавца'].map(fbo_df.set_index('Артикул продавца')[f'Остатки {region}'])

    for timeframe, file_name in file_names.items():
        order_df = pd.read_excel(os.getcwd()+r'\tables'+rf'\{file_name}', skiprows=1).iloc[:, [5, 10, 12]]
        order_df['Артикул продавца'] = order_df['Артикул продавца'].astype(str)
        sum_df = order_df.groupby('Артикул продавца')['шт..1'].sum().reset_index()
        main_df[f'Количество заказанного товара {timeframe}'] = main_df['Артикул продавца'].map(sum_df.set_index('Артикул продавца')['шт..1'])
        
        for region, warehouse_list in warehouses.items():
            filtered_df = order_df[order_df['Склад'].isin(warehouse_list)]
            sum_df = filtered_df.groupby('Артикул продавца')['шт..1'].sum().reset_index()
            main_df[f'Количество заказанного товара {timeframe} {region}'] = main_df['Артикул продавца'].map(sum_df.set_index('Артикул продавца')['шт..1'])
            main_df[f'{region} поставка {timeframe}'] = main_df.apply(_calculate, args=(timeframe, region), axis=1)

    main_df.columns = pd.MultiIndex.from_tuples(create_multi_index(), names=['Регион', 'Поле'])
    main_df.fillna(0, inplace=True)
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
            page.goto('https://seller.wildberries.ru/')
            # Авторизация
            if os.path.exists('cookies.json'):
                with open('cookies.json', 'r') as f:
                    cookies = json.loads(f.read())
                for cookie in cookies:
                    context.add_cookies([cookie])
            else:
                while True:
                    try:
                        page.wait_for_selector('.Logo__img', timeout=2000)
                        break
                    except:
                        pass
                save_cookies(context)
            
            download_fbo_data(page)
            print('FBO скачан')
            download_orders_data(page)
            print('Заказы скачаны')
            page.close()
            browser.close()

        table_collection()  
        
    except Exception as ex:
        print('Ошибка:',ex)
    finally:        
        print('Скрипт завершён')    
        input('Нажмите Enter для выхода\n')

    
if __name__ == "__main__":
    main()
    