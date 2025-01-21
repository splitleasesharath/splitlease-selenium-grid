import io
import os.path
import re
import shutil
import sys
import time
from datetime import datetime, timedelta
import geocoder
import pandas as pd
import pytz
from PIL import Image
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
import requests
import logging
import traceback
# fix for issues on 1/20/2023
# selenium 3 not compatible with python 3
# stackover flow link with info: https://stackoverflow.com/questions/65323114/robotframework-choose-file-causes-attributeerror-module-base64-has-no-attri
import base64
base64.encodestring = base64.encodebytes
# end fix

# list of computers available
computers = {'ECOM-BRM-LENOVO': "120",
             'ECOM-BRM-LPTTOP': "121",
             'ECOM-COO-LPT': "122",
             'ECOM-DET-DATASCI': "123",
             'ECOM-DET-CUSTOMERDEV': "124",
             'ECOM-DET-CRAIGSLIST': "125",
             'ECOM-DET-DEVELOPER': "126",
             'ECOM-DET-DEVELOPER2': "127",
             'ECOM-DET-PRODUCT': "128",
             'ECOM-DET-SUPPORT': "129",
             'ECOM-NYC-LAPTOP': "130",
             'Vlad': "150",
             'ECOM-DET-DESIGNER': "131",
             'ECOM-DET-ADVISOR': "132",
             'ECOM-WBLM-LPTTOP': "133",
             'ECOM-WBLM-COO': "134",
             'ECOM-DET-SmallDELL': "135",
             'mini-craigslistposter': "136",
             'ECOM-DET-LargeLaptop2': "137",
             'ECOM-DET-LargeLaptop': "138",
             'ECOM-LATITUDE1': "139",
             'ECOM-LATITUDE2': "140",
             'ECOM-LATITUDE3': "142",
             'ECOM-LATITUDE4': "141",
             'ECOM-LATITUDE5': "143",
             'ECOM-DET-LARGELAPTOP3': "144",
             'ECOM-LATITUDE6': "146",
             'ECOM-LATITUDE8': "147",
             'ECOM-LATITUDE9': "148",
             'ECOM-THIN1': "149",
             'ECOM-THIN2': "150",
             'ECOM-THIN3': "151",
             'ECOM-THIN4': "152",
             'ECOM-THIN5': "153",
             'ECOM-THIN6': "154",
             'ECOM-IDEAPAD9': "155",
             'ECOM-FRED-DEVELOPER3': "203",
             'ECOM-IDEAPAD4': "156",
             'ECOM-IDEAPAD5': "157",
             'ECOM-DET-LargeLaptop5': "158",
             'ECOM-Largegateway7': "159",
             'ECOM-IBM-LPT1': "160",
             'ECOM-DELL-LPT1': "161",
             'ECOM-DELL-LPT2': "162",
             'Laptop-SGIN1': "163",
             'Laptop-SGIN2': "164",
             'LILBLUE14': "167",
             'LILBLUE9': "168",
             'Split-Lease-5': "400"
             }


def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.dirname(__file__)
    return os.path.join(base_path, relative_path)


""" Send Vlad a slack dm"""


def send_slack_dm(txtToSend):
    slack_token = "xoxb-719141409925-2913593088630-STaPzBgOjKww3XWYWSkPtsMb"
    data = {
        'token': slack_token,
        'channel': 'U02GB3NEP6H',  # User ID.
        'as_user': True,
        'text': txtToSend
    }
    requests.post(url='https://slack.com/api/chat.postMessage',
                  data=data)


# For accessing google drive/spreadsheet
SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']

""" Manages login for google services """


def log_in():
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists(resource_path('token.json')):
        creds = Credentials.from_authorized_user_file(resource_path('token.json'), SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(resource_path('credentials.json')
                                                             , SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open(resource_path('token.json'), 'w') as token:
            token.write(creds.to_json())
    return creds


""" Downloads images from google drive"""


def uploadImages(driver, post_data):
    creds = log_in()
    service = build('drive', 'v3', credentials=creds)
    print(post_data[-1])
    folder_id = post_data[-1].split('?')[0].split('/')[-1]
    results = service.files().list(
        pageSize=10, q=f"'{folder_id}' in parents", fields="nextPageToken, files(id, name)").execute()

    items = results.get('files', [])
    item_ids = []
    if not items:
        print('No files found.')
    else:
        for item in items:
            extension = os.path.splitext(item['name'])[1]
            request = service.files().get_media(fileId=item['id'])
            fh = io.BytesIO()
            downloader = MediaIoBaseDownload(fh, request)
            done = False
            while done is False:
                status, done = downloader.next_chunk()
            fh.seek(0)
            byteImg = Image.open(fh)
            if byteImg.mode in ("RGBA", "P"):
                byteImg = byteImg.convert("RGB")
            byteImg.save('currentImg' + extension)
            driver.find_element(By.XPATH, '//*[@id="uploader"]/form/input[3]').send_keys('currentImg' + extension)
            os.remove('currentImg' + extension)


""" Cleans the data to send into forms"""


def fixed_keys(keys_to_send):
    return re.split('[^a-zA-Z]', keys_to_send)[0]


""" Pulls the tasks for the bot from the spreadsheet"""


def pull_tasks():
    # The ID and range of a sample spreadsheet.
    SAMPLE_SPREADSHEET_ID = '1eBxVRTIfnHO1miRg6grcNrprMGDZ7dvIGrEoFNDTxio'
    SAMPLE_RANGE_NAME = '2:1000'

    creds = log_in()
    service = build('sheets', 'v4', credentials=creds)
    # Call the Sheets API
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                range=SAMPLE_RANGE_NAME).execute()
    tasks = result.get('values', [])
    return tasks


""" Sets up the browser settings """


def set_up_browser(machine_name):
    driver_options = Options()
    driver_options.add_argument('--no-sandbox')
    driver_options.add_argument('--start-maximized')
    # driver_options.add_argument('--start-fullscreen')
    driver_options.add_argument('--single-process')
    driver_options.add_argument('--disable-dev-shm-usage')
    driver_options.add_argument("--incognito")
    # driver_options.add_experimental_option("detach", True)
    driver_options.add_argument('--disable-blink-features=AutomationControlled')
    driver_options.add_experimental_option('useAutomationExtension', False)
    driver_options.add_experimental_option("excludeSwitches", ["enable-automation"])
    driver_options.add_argument("disable-infobars")
    driver_options.set_capability("browserVersion", str(machine_name))
    # print('machine name  .' + str(machine_name) + '.')
    # print('reached driver.Remote()')
    print('connecting-->>>')
    driver = webdriver.Remote(
        command_executor='http://192.168.196.159:4444',
        options=driver_options
    )
   #  print('david debug reached driver.Remote() end')
    #print('david debug going to sleep, please shut me down ASAP and remove the sleep for prod run next time')
    #time.sleep(100000) # TODO REMOVE!!!!!
    return driver


""" Reposts an expired listing """


def repost(listing_data, driver):
    if len(listing_data) < 6:
        print(
            f'Listing: {listing_data} is invalid. The code requires a row to have a task, link, email, password,'
            f'host name, and the computer used to make the post. In that order.')

    # Go to the link #driver.get(listing_data[0])
    driver.get('https://accounts.craigslist.org/login/')

    # Log in
    driver.find_element(By.ID, 'inputEmailHandle').send_keys(listing_data[2])
    driver.implicitly_wait(1.5)
    driver.find_element(By.ID, 'inputPassword').send_keys(listing_data[3])
    driver.implicitly_wait(1.5)
    driver.find_element(By.ID, 'login').click()

    # Go to the old listing
    driver.get(listing_data[1])

    lat = driver.find_element(By.ID, 'map').get_attribute("data-latitude")
    long = driver.find_element(By.ID, 'map').get_attribute("data-longitude")
    category = driver.find_element(By.CSS_SELECTOR, '.category p').text
    category = category.replace('>', "")
    category = category.replace('<', "")
    category = category.strip()
    g = geocoder.mapquest([lat, long], method='reverse', key='b7bow6CgalFYwE56sSxA4JT6BpOGqsHU')
    location = g.osm['addr:city'] + ', ' + g.osm['addr:state']
    # Repost
    driver.find_element(By.CSS_SELECTOR, '.managebtn').click()
    driver.implicitly_wait(3)
    driver.find_element(By.CSS_SELECTOR, '.submit-button').click()

    # Publish
    driver.find_element(By.CSS_SELECTOR, '.button').click()

    # Get new link
    driver.implicitly_wait(5)
    link = driver.find_element(By.XPATH, '//*[@id="new-edit"]/div/div/ul/li[2]/a').get_attribute('href')

    # Update account stats
    update_stats(listing_data, driver)

    # Close browser
    driver.quit()

    # Return updated listing
    curr_time = datetime.now(pytz.timezone('America/New_York')).strftime("%H:%M")
    today_date = datetime.now(pytz.timezone('America/New_York')).strftime("%m/%d")
    host = listing_data[4]
    output = [host, listing_data[1], 'Repost', category, link, location,
              today_date, curr_time, listing_data[5], '-', '-', '-', listing_data[2]]

    # tell slack which machine reposted
    import requests
    import json
    webhookZap = "https://hooks.zapier.com/hooks/catch/9700515/bz4aegb?computername="
    print(listing_data[5])
    webhook_url = webhookZap + str(listing_data[5]) + '    repost'
    requests.post(webhook_url, headers={'Content-Type': 'application/json'})



    return output


""" Gets post data from a sheet"""


def get_post_data(link):
    # The ID and range of a sample spreadsheet.
    SAMPLE_SPREADSHEET_ID = '1eBxVRTIfnHO1miRg6grcNrprMGDZ7dvIGrEoFNDTxio'
    SAMPLE_RANGE_NAME = 'PostData!2:3000'

    creds = log_in()
    service = build('sheets', 'v4', credentials=creds)
    # Call the Sheets API
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                range=SAMPLE_RANGE_NAME).execute()
    tasks = result.get('values', [])
    df = pd.DataFrame(tasks)
    df = df[df[0] == link]
    df = df.replace([''], [None])
    return df.values.tolist()[0]


""" Posts a new listing """


def post(listing_data, driver):
    # Go to the link
    driver.get('https://accounts.craigslist.org/login/')

    # Log in
    # print(listing_data[2])
    driver.find_element(By.ID, 'inputEmailHandle').send_keys(listing_data[2])
    driver.implicitly_wait(1.5)
    # print(listing_data[3])
    driver.find_element(By.ID, 'inputPassword').send_keys(listing_data[3])
    driver.implicitly_wait(1.5)
    driver.find_element(By.ID, 'login').click()
    # print('It clicked on login')

    update_stats(listing_data, driver)

    # Go back to the main page
    driver.get('https://newyork.craigslist.org/')

    # click on create a post
    WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, 'a.cl-thumb-anchor.cl-goto-post'))).click()

    # Get the data for the post
    post_data = get_post_data(listing_data[1])

    # Select location
    # driver.find_element(By.XPATH, f"//*[contains(text(), '{post_data[1]}')]").click()

    try:
        driver.find_element(By.XPATH, f"//*[contains(text(), '{post_data[1]}')]").click()
    except NoSuchElementException:
        print(f"Could not find location option: {post_data[1]}")
        # Optionally interact with the submit button div as a fallback
        try:
            submit_button = driver.find_element(By.CSS_SELECTOR, '.submit_button .pickbutton')
            submit_button.click()
        except Exception as e:
            print(f"Error interacting with fallback submit button: {e}")
            return

    # Deal with another location question
    try:
        if driver.find_element(By.CSS_SELECTOR, '.label').text == 'choose the location that fits best:':
            if post_data[2]:
                driver.find_element(By.XPATH, f"//*[contains(text(), '{post_data[2]}')]").click()
            else:
                driver.find_element(By.XPATH, f"//*[contains(text(), 'bypass this step')]").click()
    except NoSuchElementException:
        pass

    # Select the housing category
    driver.find_element(By.XPATH, f"//*[contains(text(), 'housing offered')]").click()

    # If category is given use that category
    if post_data[3]:
        category = post_data[3]
        driver.find_element(By.XPATH, f"//*[contains(text(), '{post_data[3]}')]").click()
    # If not, use the one with the least active posts
    else:
        accounts = get_account_data()
        accounts = accounts[accounts['Email'] == listing_data[2]]
        category = accounts[['Active listings in rooms & shares',
                             'Active listings in vacation rentals',
                             'Active listings in sublets & temporary']].apply(pd.to_numeric).idxmin(axis=1)
        category = category.tolist()[0].split(' ')[-1]
        driver.find_element(By.XPATH, f"//*[contains(text(), '{category}')]").click()

    # Fix Categories
    if category == 'shares':
        category = 'rooms & shares'
    if category == 'temporary':
        category = 'sublets & temporary'
    if category == 'rentals':
        category = 'vacation rentals'

    # Set the title
    driver.find_element(By.CSS_SELECTOR, '#PostingTitle').send_keys(post_data[4])

    # Set description
    driver.find_element(By.CSS_SELECTOR, '#PostingBody').send_keys(post_data[6])

    # Set Zip
    driver.find_element(By.CSS_SELECTOR, '#postal_code').send_keys(post_data[7])

    # Set sqft if we have info
    if post_data[8]:
        driver.find_element(By.CSS_SELECTOR, '.surface_area .json-form-input').clear()
        driver.find_element(By.CSS_SELECTOR, '.surface_area .json-form-input').send_keys(post_data[8])

    # Cats
    if post_data[16] == 'TRUE':
        driver.find_element(By.CSS_SELECTOR, '.variant-checkbox .pets_cat').click()

    # Dogs
    if post_data[17] == 'TRUE':
        driver.find_element(By.CSS_SELECTOR, ".variant-checkbox .pets_dog").click()

    # Furnished
    if post_data[18] == 'TRUE':
        driver.find_element(By.CSS_SELECTOR, ".variant-checkbox .is_furnished").click()
    # No smoking
    if post_data[19] == 'TRUE':
        driver.find_element(By.CSS_SELECTOR, ".variant-checkbox .no_smoking").click()

    # Wheel
    if post_data[20] == 'TRUE':
        driver.find_element(By.CSS_SELECTOR, ".variant-checkbox .wheelchaccess").click()

    # Air
    if post_data[21] == 'TRUE':
        driver.find_element(By.CSS_SELECTOR, ".variant-checkbox .airconditioning").click()

    # EV
    if post_data[22] == 'TRUE':
        driver.find_element(By.CSS_SELECTOR, ".variant-checkbox .ev_charging").click()

    # Available on:
    if post_data[23]:
        driver.find_element(By.CSS_SELECTOR, '.hasDatepicker').send_keys(post_data[23])

    # Street and City
    if post_data[24] or post_data[25]:
        driver.find_element(By.CSS_SELECTOR, '.variant-checkbox .show_address_ok').click()
        if post_data[24]:
            driver.find_element(By.CSS_SELECTOR, '.xstreet0 .json-form-input').send_keys(post_data[24])
        if post_data[25]:
            driver.find_element(By.CSS_SELECTOR, '.city .json-form-input').send_keys(post_data[25])

    # Rent
    if post_data[5]:
        driver.find_element(By.CSS_SELECTOR, '.short-input .json-form-input').send_keys(post_data[5])

    if category == 'vacation rentals':

        # Laundry
        driver.find_element(By.CSS_SELECTOR, '#ui-id-3-button .ui-selectmenu-text').click()
        driver.find_element(By.CSS_SELECTOR, '#ui-id-3-menu').send_keys(fixed_keys(post_data[11]))
        try:
            driver.find_element(By.CSS_SELECTOR, '#ui-id-3-menu').send_keys(Keys.ENTER)
        except:
            pass

        # Apt type
        if post_data[26]:
            driver.find_element(By.CSS_SELECTOR, '#ui-id-1-button .ui-selectmenu-text').click()
            driver.find_element(By.CSS_SELECTOR, '#ui-id-1-menu').send_keys(fixed_keys(post_data[26]))
            try:
                driver.find_element(By.CSS_SELECTOR, '#ui-id-1-menu').send_keys(Keys.ENTER)
            except:
                pass

        # Parking
        driver.find_element(By.CSS_SELECTOR, '#ui-id-4-button .ui-selectmenu-text').click()
        driver.find_element(By.CSS_SELECTOR, '#ui-id-4-menu').send_keys(fixed_keys(post_data[12]))
        try:
            driver.find_element(By.CSS_SELECTOR, '#ui-id-4-menu').send_keys(Keys.ENTER)
        except:
            pass

        # Bedrooms
        driver.find_element(By.CSS_SELECTOR, '#ui-id-5-button .ui-selectmenu-text').click()
        driver.find_element(By.CSS_SELECTOR, '#ui-id-5-menu').send_keys(post_data[13])
        try:
            driver.find_element(By.CSS_SELECTOR, '#ui-id-5-menu').send_keys(Keys.ENTER)
        except:
            pass

        # Bathrooms
        driver.find_element(By.CSS_SELECTOR, '#ui-id-6-button .ui-selectmenu-text').click()
        driver.find_element(By.CSS_SELECTOR, '#ui-id-6-menu').send_keys(post_data[14])
        try:
            driver.find_element(By.CSS_SELECTOR, '#ui-id-6-menu').send_keys(Keys.ENTER)
        except:
            pass

        # Rent Period
        if post_data[15]:
            driver.find_element(By.CSS_SELECTOR, '#ui-id-1-button .ui-selectmenu-text').click()
            driver.find_elements(By.XPATH, f"//*[contains(text(), '{post_data[15]}')]")[1].click()

    if category == 'rooms & shares' or category == 'sublets & temporary':

        # Laundry
        driver.find_element(By.CSS_SELECTOR, '#ui-id-5-button .ui-selectmenu-text').click()
        driver.find_element(By.CSS_SELECTOR, '#ui-id-5-menu').send_keys(fixed_keys(post_data[11]))
        try:
            driver.find_element(By.CSS_SELECTOR, '#ui-id-5-menu').send_keys(Keys.ENTER)
        except:
            pass

        # Private Room
        driver.find_element(By.CSS_SELECTOR, '#ui-id-2-button .ui-selectmenu-text').click()
        driver.implicitly_wait(5)
        if post_data[9] == "TRUE":
            driver.find_elements(By.XPATH, f"//*[text() = 'private room']")[2].click()
        else:
            driver.find_elements(By.XPATH, f"//*[text() = 'room not private']")[1].click()

        # Private Bath
        driver.find_element(By.CSS_SELECTOR, '#ui-id-4-button .ui-selectmenu-text').click()
        driver.implicitly_wait(5)
        if post_data[10] == "TRUE":
            driver.find_elements(By.XPATH, f"//*[text() = 'private bath']")[2].click()
        else:
            driver.find_elements(By.XPATH, f"//*[text() = 'no private bath']")[1].click()

        # Apt type
        if post_data[26]:
            driver.find_element(By.CSS_SELECTOR, '#ui-id-3-button .ui-selectmenu-text').click()
            driver.find_element(By.CSS_SELECTOR, '#ui-id-3-menu').send_keys(fixed_keys(post_data[26]))
            try:
                driver.find_element(By.CSS_SELECTOR, '#ui-id-3-menu').send_keys(Keys.ENTER)
            except:
                pass

        # Parking
        driver.find_element(By.CSS_SELECTOR, '#ui-id-6-button .ui-selectmenu-text').click()
        driver.find_element(By.CSS_SELECTOR, '#ui-id-6-menu').send_keys(fixed_keys(post_data[12]))
        try:
            driver.find_element(By.CSS_SELECTOR, '#ui-id-6-menu').send_keys(Keys.ENTER)
        except:
            pass

        # Bedrooms
        driver.find_element(By.CSS_SELECTOR, '#ui-id-7-button .ui-selectmenu-text').click()
        driver.find_element(By.CSS_SELECTOR, '#ui-id-7-menu').send_keys(post_data[13])
        try:
            driver.find_element(By.CSS_SELECTOR, '#ui-id-7-menu').send_keys(Keys.ENTER)
        except:
            pass

        # Bathrooms
        driver.find_element(By.CSS_SELECTOR, '#ui-id-8-button .ui-selectmenu-text').click()
        driver.find_element(By.CSS_SELECTOR, '#ui-id-8-menu').send_keys(post_data[14])
        try:
            driver.find_element(By.CSS_SELECTOR, '#ui-id-8-menu').send_keys(Keys.ENTER)
        except:
            pass

        # Rent Period
        if post_data[15]:
            driver.find_element(By.CSS_SELECTOR, '#ui-id-1-button .ui-selectmenu-text').click()
            # driver.find_element(By.XPATH, f"//option[contains(text(), '{post_data[15]}')]").click()
            print(driver.find_elements(By.XPATH, f"//*[contains(text(), '{post_data[15]}')]"))
            driver.find_elements(By.XPATH, f"//*[contains(text(), '{post_data[15]}')]")[1].click()

    # Submit
    WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, '.submit-button'))).click()

    # Approve Location
    # Deal with New Jersey
    WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, '.bigbutton'))).click()
    try:
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, '.medium-pickbutton+ .medium-pickbutton'))).click()
    except:
        pass

    # Set Up image upload button
    WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, '#classic'))).click()

    # Upload images
    uploadImages(driver, post_data)

    # Submit
    WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, 'button#doneWithImages.bigbutton'))).click()

    # Post
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, '.button'))).click()

    #tell slack which machine posted
    import requests
    import json
    webhookZap = "https://hooks.zapier.com/hooks/catch/9700515/bz4aegb?computername="
    print(listing_data[5])
    webhook_url = webhookZap + str(listing_data[5])
    requests.post(webhook_url, headers={'Content-Type': 'application/json'})




    # Update spreadsheet
    host = listing_data[4]
    link = driver.find_element(By.XPATH, '//*[@id="new-edit"]/div/div/ul/li[2]/a').get_attribute('href')
    curr_time = datetime.now(pytz.timezone('America/New_York')).strftime("%H:%M")
    location = f"{post_data[2]}, {post_data[1].capitalize()}" if post_data[2] is not None else post_data[1].capitalize()
    today_date = datetime.now(pytz.timezone('America/New_York')).strftime("%m/%d")
    output = [host, listing_data[1], 'Post', category, link, location,
              today_date, curr_time, listing_data[5], '-', '-', '-', listing_data[2]]

    # Update account stats
    update_stats(listing_data, driver)

    # Close browser
    driver.quit()

    return output


""" Renews the listing """


def renew(listing_data, driver):
    # Go to the link
    driver.get('https://accounts.craigslist.org/login/')

    # Log in
    driver.find_element(By.ID, 'inputEmailHandle').send_keys(listing_data[2])
    driver.implicitly_wait(1.5)
    driver.find_element(By.ID, 'inputPassword').send_keys(listing_data[3])
    driver.implicitly_wait(1.5)
    driver.find_element(By.ID, 'login').click()

    # Go to the old listing
    driver.get(listing_data[1])

    # Get listing data
    lat = driver.find_element(By.ID, 'map').get_attribute("data-latitude")
    long = driver.find_element(By.ID, 'map').get_attribute("data-longitude")
    category = driver.find_element(By.CSS_SELECTOR, '.category p').text
    category = category.replace('>', "")
    category = category.replace('<', "")
    category = category.strip()
    g = geocoder.mapquest([lat, long], method='reverse', key='b7bow6CgalFYwE56sSxA4JT6BpOGqsHU')
    location = g.osm['addr:city'] + ', ' + g.osm['addr:state']
    curr_time = datetime.now(pytz.timezone('America/New_York')).strftime("%H:%M")
    today_date = datetime.now(pytz.timezone('America/New_York')).strftime("%m/%d")
    host = listing_data[4]
    title = driver.find_element(By.XPATH, '//*[@id="titletextonly"]').text

    try:
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable(
                (By.XPATH, '//*[@id="manage-posting"]/div[1]/table/tbody/tr[6]/td[1]/div/form/input[3]'))).click()
    except:
        print(f'Error, check if {listing_data[1]} is already reposted')

    time.sleep(3)
    update_stats(listing_data, driver)

    new_link = driver.find_elements(By.XPATH, f"//*[contains(text(), '{title}')]")[0].get_attribute('href')
    driver.quit()

    # Return updated listing
    output = [host, listing_data[1], 'Renew', category, new_link, location,
              today_date, curr_time, listing_data[5], '-', '-', '-', listing_data[2]]

    # tell slack which machine posted
    import requests
    import json
    webhookZap = "https://hooks.zapier.com/hooks/catch/9700515/bz4aegb?computername="
    print(listing_data[5])
    webhook_url = webhookZap + str(listing_data[5]) + '    renewal'
    requests.post(webhook_url, headers={'Content-Type': 'application/json'})

    return output


""" Updates the results spreadsheet with the new data after posting or reposting """


def update(listing_data_updated):
    # The ID and range of a sample spreadsheet.
    SAMPLE_SPREADSHEET_ID = '1eBxVRTIfnHO1miRg6grcNrprMGDZ7dvIGrEoFNDTxio'
    SAMPLE_RANGE_NAME = 'Results!2:1000'
    creds = log_in()
    service = build('sheets', 'v4', credentials=creds)
    body = {
        'values': listing_data_updated
    }

    # Update new sheet
    result = service.spreadsheets().values().append(
        spreadsheetId=SAMPLE_SPREADSHEET_ID, range=SAMPLE_RANGE_NAME,
        valueInputOption='RAW', body=body).execute()
    print('{0} cells appended to the new sheet.'.format(result
                                                        .get('updates')
                                                        .get('updatedCells')))

    # Update old sheet
    old_SPREADSHEET_ID = '1Juftetgo4c2i9SmFkAE9QvE0fVQdpQDZIKziQ8EJSjw'
    old_RANGE_NAME = 'Enumerated Postings!2:1000'
    result = service.spreadsheets().values().append(
        spreadsheetId=old_SPREADSHEET_ID, range=old_RANGE_NAME,
        valueInputOption='RAW', body=body).execute()
    print('{0} cells appended to the old sheet.'.format(result
                                                        .get('updates')
                                                        .get('updatedCells')))


""" This one runs the program """


def main():
    print('about to pull tasks')
    tasks = pull_tasks()
    print('pulled tasks, now iterating through the list')
    for row_num, task in enumerate(tasks):
        # print('len task          ', len(task))
        if len(task) not in [6, 7]:
            print(f'Problem with Row {row_num} in the Tasks tab')
        else:
            try:
                if len(task) == 7:
                    time_col = 6
                    post_time = datetime.strptime(task[time_col], '%m/%d/%Y %H:%M:%S')
                    print('Waiting until ' + task[time_col] + ' for next task...')
                    time.sleep((post_time - datetime.now()).seconds)
                    print('Now posting!')
                machine_col = 5

                # gets stuck at setupbrowser
                driver = set_up_browser(computers[task[machine_col]])
                #print('david debug:', 'reached set up finished')
                print(task[0])
                if task[0] == 'Repost':
                    update([repost(task, driver)])
                if task[0] == 'Post':
                    update([post(task, driver)])
                if task[0] == 'Renew':
                    #print('david debug:', 'reached renew')
                    update([renew(task, driver)])
                    #print('david debug:', 'reached renew finished')

            except Exception as exception:
                traceback.print_exc()
                print(f'Failed to execute Row {row_num} in the Tasks tab')
                logger.error(exception)
                send_slack_dm(f"Task {task} just threw an error \n {traceback.format_exc()}")
                try:
                    driver.quit()
                except:
                    pass


""" Pulls currently known data for accounts in use """


def get_account_data():
    # The ID and range of a sample spreadsheet.
    SAMPLE_SPREADSHEET_ID = '1eBxVRTIfnHO1miRg6grcNrprMGDZ7dvIGrEoFNDTxio'
    'AccountData!2:1000'

    creds = log_in()
    service = build('sheets', 'v4', credentials=creds)
    # Call the Sheets API
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                range='AccountData!1:1000').execute()
    data = result.get('values')
    df = pd.DataFrame(data[1:], columns=data[0])
    df = df.replace([''], [None])
    return df


""" Updates the sheet with the number of current active, flagged, and expired listings """


def update_stats(listing_data, driver):
    # The ID and range of a sample spreadsheet.
    SAMPLE_SPREADSHEET_ID = '1eBxVRTIfnHO1miRg6grcNrprMGDZ7dvIGrEoFNDTxio'
    SAMPLE_RANGE_NAME = 'AccountData!1:1000'
    creds = log_in()
    service = build('sheets', 'v4', credentials=creds)
    driver.get('https://accounts.craigslist.org/login/home')
    time.sleep(3)
    driver.refresh()
    email = listing_data[2]
    total_posts = len(driver.find_elements(By.CSS_SELECTOR, '.gc')) / 2
    active_listings = len(driver.find_elements(By.CSS_SELECTOR, '.active .gc')) / 2
    removed_listings = len(driver.find_elements(By.CSS_SELECTOR, '.removed .gc')) / 2
    expired_listings = len(driver.find_elements(By.CSS_SELECTOR, '.expired .gc')) / 2
    active_listings_categories = [el.text.strip() for el in driver.find_elements(By.CSS_SELECTOR, '.areacat.active')]

    count_rooms_shares = 0
    count_vacation_rentals = 0
    count_sublets_temporary = 0
    for element in active_listings_categories:
        if 'rooms & shares' in element:
            count_rooms_shares += 1
        if 'vacation rentals' in element:
            count_vacation_rentals += 1
        if 'sublets & temporary' in element:
            count_sublets_temporary += 1
    update_for_accounts = [email, total_posts, active_listings,
                           count_rooms_shares, count_vacation_rentals,
                           count_sublets_temporary, expired_listings, removed_listings]
    df = get_account_data()
    if email not in df.Email.values:
        body = {
            'values': [update_for_accounts]
        }
        service.spreadsheets().values().append(
            spreadsheetId=SAMPLE_SPREADSHEET_ID, range=SAMPLE_RANGE_NAME,
            valueInputOption='RAW', body=body).execute()
    else:
        df.loc[df['Email'] == email, ['Email', 'Total Posts', 'Active Listings',
                                      'Active listings in rooms & shares',
                                      'Active listings in vacation rentals',
                                      'Active listings in sublets & temporary', 'Number of expired listings',
                                      'Times flagged']] = update_for_accounts
        data = [df.columns.values.tolist()]
        data.extend(df.values.tolist())
        value_range_body = {"values": data}
        service.spreadsheets().values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=SAMPLE_RANGE_NAME,
                                               valueInputOption='RAW', body=value_range_body).execute()


# Run the code
logging.basicConfig(filename='test.log', level=logging.DEBUG,
                    format='%(asctime)s %(levelname)s %(name)s %(message)s')
logger = logging.getLogger(__name__)
try:
    main()
except Exception as bigErr:
    send_slack_dm(f"Error {traceback.format_exc()} just occurred")
    logger.error(bigErr)