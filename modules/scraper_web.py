from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import WebDriverException
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import StaleElementReferenceException, NoSuchElementException, NoSuchWindowException
from modules.helpers import *
from modules.const.settings import SETTINGS
from modules.const.colors import fore
import re
import time
import json
import xlsxwriter

def scrape(args):
    '''
    Scrapes the results and puts them in the excel spreadsheet.

    Parameters:
            args (object): CLI arguments
    '''
    thrshhold = 4.5
    if args.pages is not None:
        SETTINGS["PAGE_DEPTH"] = args.pages
    SETTINGS["BASE_QUERY"] = args.query
    SETTINGS["PLACES"] = args.places.split(',')
    
    # Created driver and wait
    driver = webdriver.Chrome()
    wait = WebDriverWait(driver, 15)

    # Set main box ID name
    BOX_CLASS = "center_col"

    # Initialize workbook / worksheet
    workbook = xlsxwriter.Workbook(
        f'./Output/Web_ScrapedData_Business_{args.query} {args.places}.xlsx')
    worksheet = workbook.add_worksheet()

    # Headers and data
    data = {
         "name": "",
         "address": "",
         "phone": "",
         "Website": "",
         "rating": ""
     }
    headers = generate_headers(args, data)
    print_table_headers(worksheet, headers)

    # Start from second row in xlsx, as first one is reserved for headers
    row = 1

    # Remember scraped addresses to skip duplicates
    addresses_scraped = {}

    start_time = time.time()

    for place in SETTINGS["PLACES"]:
        # Go to the index page
        driver.get('https://www.google.com/search?hl=en')

        # Build the query string
        query = "{0} {1}".format(SETTINGS["BASE_QUERY"], place)
        print(f"{fore.GREEN}Moving on to {place}{fore.RESET}")

        # Fill in the input and press enter to search
        q_input = driver.find_element_by_name("q")
        q_input.send_keys(query, Keys.ENTER)
        try:
            driver.find_element_by_xpath(
                "//*[@id='Odp5De']/div/div/div[2]/div[1]/div[2]/g-more-link/a").click()
        except:
            print('No business found!')
        try:
            w = wait.until(
                EC.presence_of_element_located((By.ID, BOX_CLASS))
            )
        except:
            print('not found')
            continue

        while True:
            # Get all the results boxes
            boxes = driver.find_elements_by_xpath(
                "//div[contains(@class, 'uMdZh tIxNaf')]")
            print(f'found : {len(boxes)}')
            # Loop through all boxes and get the info from it and store into an excel
            for box in boxes:
                try:
                    rating = box.find_element_by_class_name('yi40Hd').text
                except:
                    rating = 0
                if float(rating) <= thrshhold:
                    # Just get the values, add only after we determine this is not a duplicate (or duplicates should not be skiped)
                    name = box.find_element_by_class_name("OSrXXb").text
                    # print(f'name: {name}')
                    #box.click()
                    cls = box.find_element_by_class_name("rllt__details")
                    try:
                      address = cls.find_element_by_xpath(".//div[3]").text
                    except:
                       address = ''
                    # print(f'address: {address}')
                    scraped = address in addresses_scraped
                    print(f'rating: {rating}')
                    if scraped and args.skip_duplicate_addresses:
                        print(f"{fore.WARNING}Skipping {name} as duplicate by address{fore.RESET}")
                    else:
                        try:
                           phone = cls.find_element_by_xpath(".//div[4]").text
                        except:
                           phone = ''
                        # print(f'Phone: {phone}')
                        if scraped:
                            addresses_scraped[address] += 1
                            print(f"{fore.WARNING}Currently scraping on{fore.RESET}: {name}, for the {addresses_scraped[address]}. time")
                        else:
                            addresses_scraped[address] = 1
                            print(f"{fore.GREEN}Currently scraping on{fore.RESET}: {name}")
                             
                        # Only if user wants to get the URL to, get it
                        if args.scrape_website:
                            try:
                                url = box.find_element_by_class_name(
                                    "Q7PwXb").get_attribute("href")
                            except:
                                url = ''
                            # print(f'link: {url}')
                        if name != '':
                            data["name"] = name
                            mean_data = f'{address} · {phone}'
                            p_data = mean_data.split('·')
                            add = ""
                            for item in p_data:
                                item = item.strip()
                                if len(item) > 0 and item.startswith('Open') == False and item.startswith('Closes') == False and item.startswith('In') == False and item[0].isdigit() == False:
                                    if item.startswith('+1') == False or item.startswith('(') == True:
                                        data["address"] = item
                                        add = item
                                    elif item.startswith('+1') == True or item.startswith('(') == True:
                                        data["phone"] = item
                            data['Website'] = url
                            data["rating"] = rating
                            # If additional output is requested
                            if add != "Temporarily closed":
                                print(data)
                                if args.verbose:
                                    print(json.dumps(data, indent=1))
                                write_data_row(worksheet, data, row)
                                row += 1
            try:
                driver.find_element_by_xpath(
                "//*[@id='pnnext']").click()
            except:
                break
            time.sleep(5)
        print("-------------------")

    workbook.close()
    driver.close()

    end_time = time.time()
    elapsed = round(end_time-start_time, 2)
    print(f"{fore.GREEN}Done. Time it took was {elapsed}s{fore.RESET}")