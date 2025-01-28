from selenium import webdriver
from selenium.webdriver.common.by import *
from time import sleep
from selenium.webdriver.common.keys import Keys
import xlwt
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from multiprocessing import Process
import multiprocessing
from datetime import datetime
from selenium.webdriver.chrome.options import Options
import time
import os
import math
from enum import Enum


class StoreOption(Enum):
    COSTCO = 1
    WHOLESALE = 2

store_option = StoreOption.COSTCO
# Change to True to hide chrome browsers
is_headless = False
# Add Product's to list here
product_list = ["Milk", "Chicken"]
chrome_options = Options()
chrome_options.add_argument("--headless")
date_str = datetime.today().strftime('%Y-%m-%d')
file_path = "C:/Users/16138/Documents/wholesale-prices " + date_str + ".xls"
stores = [
    ["Ontario", "Ottawa", "Ottawa"],
    ["Ontario", "Mississauga", "Mississauga"],
    ["British Columbia", "Victoria", "Victoria"],
    ["Quebec", "Quebec City", "Club Entrepot - Quebec City"],
    ["Alberta", "Calgary", "Calgary North"],
    ["Newfoundland and Labrador", "St. John's", "St. John's"]
]
num_threads = 6


def get_stores():
    if is_headless:
        driver = webdriver.Chrome(options=chrome_options)
    else:
        driver = webdriver.Chrome()
    driver.get("https://www.wholesaleclub.ca/")
    sleep(5)
    select_options = driver.find_elements(By.XPATH, "//select[@name='provinceSelect']/option")
    count = 0
    stores = []
    for option in select_options:
        if count == 0:
            count = count + 1
            continue
        province = option.text

        option.click()
        city_element = driver.find_elements(By.XPATH, "//select[@name='citySelect']/option")
        city_count = 0

        for city in city_element:
            if city_count == 0:
                city_count = city_count + 1
                continue
            city_text = city.text
            city.click()

            store_element = driver.find_elements(By.XPATH, "//select[@name='storeSelect']/option")
            is_first = True
            for store in store_element:
                if is_first:
                    is_first = False
                    continue
                store_text = store.text

                stores.append([province, city_text, store_text])
                store.click()
    driver.close()
    return stores


def scrape_store(store, return_dict):
    excel_data = []

    if is_headless:
        driver = webdriver.Chrome(options=chrome_options)
    else:
        driver = webdriver.Chrome()
    driver.get("https://www.wholesaleclub.ca/")
    province = store[0]
    city = store[1]
    store_name = store[2]
    WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "//select[@name='provinceSelect']/option[text()='" + province + "']"))).click()

    excel_data.append([province + " " + city + " " + store_name])

    city_name_split = store[1].split("'")
    city_name = city_name_split[0]
    city_element = driver.find_element(By.XPATH, "//select[@name='citySelect']/option[contains(text(),'" + city_name + "')]")
    city_element.click()

    store_name_split = store[2].split("'")
    store_name = store_name_split[0]
    store_element = driver.find_element(By.XPATH, "//select[@name='storeSelect']/option[contains(text(),'" + store_name + "')]")
    store_element.click()

    shop_button = driver.find_element(By.XPATH, "//button[text()='Shop']")
    shop_button.click()

    columns = ["Brand", "Product Name", "Size", "Price", "Unit", "Factored Price", "Factored Unit"]
    excel_data.append(columns)

    is_first_search = True
    for product in product_list:
        search_bar = driver.find_element(By.XPATH, "//input")
        excel_data.append([product])
        if not is_first_search:
            WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.XPATH, "//button[@type='reset']"))).click()
        is_first_search = False
        search_bar.send_keys(product)
        search_bar.send_keys(Keys.ENTER)
        sleep(2)
        product_elements = driver.find_elements(By.XPATH, "//div[@class='product-tracking']")

        for product_element in product_elements:
            try:
                brand = product_element.find_element(By.XPATH, ".//span[@class='product-name__item product-name__item--brand']").text
            except:
                brand = "N/A"
            try:
                name = product_element.find_element(By.XPATH, ".//span[@class='product-name__item product-name__item--name']").text
            except:
                name = "N/A"
            try:
                size = product_element.find_element(By.XPATH, ".//span[@class='product-name__item product-name__item--package-size']").text
            except:
                size = "N/A"
            try:
                price = product_element.find_element(By.XPATH,".//span[@class='price__value selling-price-list__item__price selling-price-list__item__price--now-price selling-price-list__item__price--__value']").text
            except:
                price = "N/A"
            try:
                unit = product_element.find_element(By.XPATH,".//span[@class='price__unit selling-price-list__item__price selling-price-list__item__price--now-price selling-price-list__item__price--__unit']").text
            except:
                unit = "N/A"
            try:
                factored_price = product_element.find_element(By.XPATH,".//span[@class='price__value comparison-price-list__item__price__value']").text
            except:
                factored_price = "N/A"
            try:
                factored_unit = product_element.find_element(By.XPATH, ".//span[@class='price__unit comparison-price-list__item__price__unit']").text
            except:
                factored_unit = "N/A"

            scraped_row = [brand, name, size, price, unit, factored_price, factored_unit]
            excel_data.append(scraped_row)
    driver.close()

    return_dict[store_name] = excel_data


def main():
    start_time = time.time()
    #stores = get_stores()
    jobs = []
    manager = multiprocessing.Manager()
    return_dict = manager.dict()
    for store in stores:
        process = Process(target=scrape_store, args=(store, return_dict,))
        jobs.append(process)

    num_of_batches = int(math.ceil(len(jobs) / num_threads))
    job_index = 0
    for batch_num in range(num_of_batches):
        running_jobs = []
        for cur_job in range(num_threads):
            if job_index == len(jobs):
                break
            job = jobs[job_index]
            job.start()
            running_jobs.append(job)
            job_index = job_index + 1

        for running_job in running_jobs:
            running_job.join()

    book = xlwt.Workbook()
    for store_name, data in return_dict.items():

        sheet = book.add_sheet(store_name)
        for row_index in range(len(data)):
            row = data[row_index]
            excel_row = sheet.row(row_index)
            for col_index in range(len(row)):
                col = row[col_index]
                excel_row.write(col_index, col)
    try:
        os.remove(file_path)
    except:
        print("")

    book.save(file_path)
    total_seconds = time.time() - start_time

    print("Finished in: " + str(total_seconds))


if __name__ == "__main__":
    main()
