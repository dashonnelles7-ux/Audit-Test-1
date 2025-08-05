import time
import requests
import re
import json
import os
import traceback


def read_json_cookies(file_path):

    with open(file_path, "r") as file:
        try:
            cookies_dict = json.load(file)
            print(cookies_dict)
        except json.JSONDecodeError:
            print("Error: JSON file is corrupted, will be initialized with empty object")
            cookies_dict = {}
        except FileNotFoundError as e:
            traceback.print_exc()
            time.sleep(5)
            raise e

    print("Cookies loaded into session")

    return cookies_dict


def check_cookies_work(cookies_dict, file_path):

    session = requests.Session()

    session.cookies.update(cookies_dict)

    dashboard_url = "https://app.sellerboard.com/en/dashboard/"

    try:
        for i in range(3):

            dashboard_response = session.get(dashboard_url)

            if dashboard_response.status_code == 200:

                if "Dashboard" in dashboard_response.text:
                    print("Successfully received the dashboard page!")
                    print()
                    #print(dashboard_response.text)
                    return cookies_dict

                else:
                    print(f"Error loading dashboard. Status code:: {dashboard_response.status_code}")
                    print()
                    cookies_dict = login_seller()
                    return write_json_cookies(cookies_dict, file_path)

            else:
                print("Dashboard response status code: ",dashboard_response.status_code)
                time.sleep(5)
                continue

    except requests.exceptions.TooManyRedirects as e:
        print("Too many redirects encountered. Please check the URL or cookies.")
        print("Starting to update cookies")
        print(f"Error details: {e}")
        cookies_dict = login_seller()
        return write_json_cookies(cookies_dict, file_path)


def write_json_cookies(cookies_dict, file_path):

    with open(file_path, "w") as file:
        json.dump(cookies_dict, file)

    print(f"Cookies updated in {file_path}")

    return cookies_dict


def login_seller():

    session = requests.Session()

    headers = {
        'User-Agent': "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        'Referer': 'https://app.sellerboard.com/',
        'Accept-Language': 'en-US,en;q=0.9'
    }


    sellerboard_url = "https://sellerboard.com/"
    response = session.get(sellerboard_url)

    login_page_url = "https://app.sellerboard.com/en/auth/login/"
    response = session.get(login_page_url, headers=headers)
    print("response.status_code ",response.status_code)

    csrf_key_pattern = r'(?<=<input\ type="hidden"\ name=").*(?="\ value=")'
    csrf_value_pattern = r'(?<="\ value=").*(?="/>)'

    csrf_key = re.search(csrf_key_pattern, response.text)
    csrf_value = re.search(csrf_value_pattern, response.text)

    if csrf_key and csrf_value:
        csrf_key = csrf_key.group()
        csrf_value = csrf_value.group()
        print(f"CSRF Key: {csrf_key}")
        print(f"CSRF Value: {csrf_value}")
    else:
        print("CSRF not found!")
        exit()

    email = str(os.getenv('SB_USERNAME'))
    password = str(os.getenv('SB_PASSWORD'))

    login_url = "https://app.sellerboard.com/en/auth/login/"
    payload = {
        'login': email,
        'password': password,
        'keeploggedin': '1',
        csrf_key: csrf_value,
        'selling_partner_id': '',
        'amazon_callback_uri': '',
        'amazon_state': '',
        'jdq': '1'
    }

    for i in range(20):

        login_response = session.post(login_url, data=payload)
        #login_response = session.post(login_url, data=payload)

        if login_response.status_code == 200 and "Dashboard" in login_response.text:
            print("Successful login")
            # for cookie in session.cookies:
            #     print(f"{cookie.name} = {cookie.value}")
            break

        elif login_response.status_code == 429:
            print(login_response)
            time.sleep(900)
            continue


        else:
            print("Login error!")
            print("Status code:", login_response.status_code)
            time.sleep(1)
            print(login_response.text)


    dashboard_url = "https://app.sellerboard.com/en/dashboard/"
    dashboard_response = session.get(dashboard_url)

    if dashboard_response.status_code == 200:
        print("Successfully received the dashboard page!")
        cookies_dict = session.cookies.get_dict()

    else:
        print("Error getting dashboard page!")
        print("Status code:", dashboard_response.status_code)
        print("Server response:", dashboard_response.text)

    print('LOGIN')

    return cookies_dict


def start_log_in(file_path):

    cookies_dict = read_json_cookies(file_path)

    cookies_dict = check_cookies_work(cookies_dict, file_path)

    return  cookies_dict

