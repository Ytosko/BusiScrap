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

    # Set main box class name
    BOX_CLASS = "ecceSd"

    # Initialize workbook / worksheet
    workbook = xlsxwriter.Workbook(
        f'./Output/map_ScrapedData_Business_{args.query} {args.places}.xlsx')
    worksheet = workbook.add_worksheet()

    # Headers and data
    data = {
        "name": "",
        "phone": "",
        "address": "",
        "website": "",
        "rating": "",
        "email": ""
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
        driver.get(SETTINGS["MAPS_INDEX"])

        # Build the query string
        query = "{0} {1}".format(SETTINGS["BASE_QUERY"], place)
        print(f"{fore.GREEN}Moving on to {place}{fore.RESET}")

        # Fill in the input and press enter to search
        q_input = driver.find_element_by_name("q")
        q_input.send_keys(query, Keys.ENTER)
        
        # Wait for the results page to load. If no results load in 10 seconds, continue to next place
        try:
            w = wait.until(
                EC.presence_of_element_located((By.CLASS_NAME, BOX_CLASS))
            )
            els = driver.find_elements(By.CSS_SELECTOR, '.TFQHme')
            driver.execute_script("arguments[0].scrollIntoView();", els[-1])
        except:
            print('not found')
            continue
        # Loop through pages and results
        def infinite_scroll(driver):
            number_of_elements_found = 0
            while True:
                els = driver.find_elements(By.CSS_SELECTOR, '.TFQHme')
                if number_of_elements_found == len(els):
                    # Reached the end of loadable elements
                    break
                try:
                    driver.execute_script("arguments[0].scrollIntoView();", els[-1])
                    number_of_elements_found = len(els)
        
                except StaleElementReferenceException:
                    # Possible to get a StaleElementReferenceException. Ignore it and retry.
                    pass
                time.sleep(3)

        for _ in range(0, SETTINGS["PAGE_DEPTH"]):
            infinite_scroll(driver)
            # Get all the results boxes
            boxes = driver.find_elements_by_xpath("//div[contains(@class, 'bfdHYd Ppzolf OFBs3e')]")
            print(f'found : {len(boxes)}')
            # Loop through all boxes and get the info from it and store into an excel
            for box in boxes:
                rating = box.find_element_by_class_name('MW4etd').text
                if float(rating) <= thrshhold:
                    # Just get the values, add only after we determine this is not a duplicate (or duplicates should not be skiped)
                    name = box.find_element_by_class_name(
                        "fontHeadlineSmall").find_element_by_xpath(".//span[1]").text
                    print(f'name: {name}')
                    #box.click()
                    cls = box.find_elements_by_class_name("W4Efsd")
                    try:
                      address = cls[len(cls)-2].find_element_by_xpath(".//span[2]//jsl//span[2]").text
                    except:
                       address = ''
                    print(f'address: {address}')
                    scraped = address in addresses_scraped
                    print(f'rating: {rating}')
                    if scraped and args.skip_duplicate_addresses:
                        print(f"{fore.WARNING}Skipping {name} as duplicate by address{fore.RESET}")
                    else:
                        try:
                           phone = cls[len(cls)-1].find_element_by_xpath(".//span[2]//jsl//span[2]").text
                        except:
                           phone = ''
                        print(f'Phone: {phone}')
                        if scraped:
                            addresses_scraped[address] += 1
                            print(f"{fore.WARNING}Currently scraping on{fore.RESET}: {name}, for the {addresses_scraped[address]}. time")
                        else:
                            addresses_scraped[address] = 1
                            print(f"{fore.GREEN}Currently scraping on{fore.RESET}: {name}")
                             
                        # Only if user wants to get the URL to, get it
                        if args.scrape_website:
                            url = box.find_element_by_class_name(
                                "Rwjeuc").find_element_by_xpath(
                                ".//a").get_attribute("href")
                            print(f'link: {url}')
                            website, email = get_website_data(url)
                            if website is not None:
                                data["website"] = website
                            if email is not None:
                                data["email"] = ','.join(email)
    
                        data["name"] = name
                        data["address"] = address
                        data["phone"] = phone
                        data["rating"] = rating
                        print(data)
                        # If additional output is requested
                        if args.verbose:
                            print(json.dumps(data, indent=1))
    
                        write_data_row(worksheet, data, row)
                        row += 1
            # Wait for the next page to load
            # time.sleep(5)
        print("-------------------")

    workbook.close()
    driver.close()

    end_time = time.time()
    elapsed = round(end_time-start_time, 2)
    print(f"{fore.GREEN}Done. Time it took was {elapsed}s{fore.RESET}")