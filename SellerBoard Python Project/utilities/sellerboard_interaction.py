from io import BytesIO
import openpyxl
import json
import msvcrt
import sys
import os

sys.path.append(
    "C:\\Program Files\\ZennoLab\\RU\\ZennoPoster Pro V7\\7.7.0.0\\Progs\\Projects\\AMZ Professional\\SellerBoard Python Project\\utilities")
sys.path.append(
    "D:\\PyCharm Projects\\SellerBoard Python Project\\utilities")

import re
import requests
import os
import time
from dotenv import load_dotenv
from datetime import datetime, timedelta, timezone
from tables import Table
from log_in import start_log_in
# from logger_config import logger

import xlrd

load_dotenv()

user_agent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"

headers = {
    'User-Agent': "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    'Referer': 'https://app.sellerboard.com/',
    'Accept-Language': 'en-US,en;q=0.9'
}


def switch_account(table_sb_acc_id, data_token_key, data_token_value, session):
    headers = {
        'User-Agent': user_agent,
        # 'Sellerboard-Account-Id': sb_acc_id,
        # 'Sellerboard-User-Id': sb_user_id,
        # 'X-Requested-With': 'XMLHttpRequest',
        "origin": "https://app.sellerboard.com",
        "pragma": "no-cache",
        "sec-ch-ua": '"Chromium";v="128", "Not;A=Brand";v="24", "Opera";v="114"',
        "sec-ch-ua-mobile": "?0",
        "sec-ch-ua-platform": '"Windows"',
        "sec-fetch-dest": "empty",
        "sec-fetch-mode": "cors",
        "sec-fetch-site": "same-origin",
        "accept": "*/*",
        "accept-encoding": "gzip, deflate, br, zstd",
        "accept-language": "ru-RU,ru;q=0.9,en-US;q=0.8,en;q=0.7",
        "cache-control": "no-cache",
        "content-type": "application/x-www-form-urlencoded; charset=UTF-8",
        "priority": "u=1, i"
    }

    post_data = {
        "account": table_sb_acc_id,
        data_token_key: data_token_value
    }
    # post_data = {
    #     "account": table_sb_acc_id,
    #     "QTBxQ0JaQWdxWGM0cUgzUFAzelJRZz09": "OXJuSmdJbVEvVUptVTM3M0NORXpFdz09"
    # }

    switch_account_url = "https://app.sellerboard.com/en/setup/switchAccount"

    switch_account_response = session.post(switch_account_url, data=post_data, headers=headers)

    switchAccount_resp = switch_account_response.text

    if switch_account_response.status_code == 200:
        print("Successfully switched to account!")


    else:
        print("Error switching to account!")
        print("Status code:", switch_account_response.status_code)
        time.sleep(10)

        switchAccount_resp = None

    return switchAccount_resp


def products(count_switch, session):
    while count_switch != 2:
        dashboard_url = "https://app.sellerboard.com/en/export"

        headers = {
            'host': 'app.sellerboard.com',
            'sec-ch-ua': '"Chromium";v="130", "Google Chrome";v="130", "Not?A_Brand";v="99"',
            'sec-ch-ua-mobile': '?0',
            'sec-ch-ua-platform': '"Windows"',
            'upgrade-insecure-requests': '1',
            'user-agent': user_agent,
            'accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7',
            'sec-fetch-site': 'same-origin',
            'sec-fetch-mode': 'navigate',
            'sec-fetch-user': '?1',
            'sec-fetch-dest': 'document',
            'referer': 'https://app.sellerboard.com/en/export',
            'accept-encoding': 'gzip, deflate, br, zstd',
            'accept-language': 'ru-RU,ru;q=0.9,en-US;q=0.8,en;q=0.7',
            'priority': 'u=0, i',

        }

        dashboard_response = session.get(dashboard_url, headers=headers)

        dashboard_resp = dashboard_response.text

        if dashboard_response.status_code == 200:

            print("Successfully received the dashboard page!")

            csrf_token_match = re.search(r'"CSRF_TOKEN":"(.*?)"', dashboard_resp)
            if csrf_token_match:
                csrf_token = csrf_token_match.group(1)

            csrf_token_key_match = re.search(r'"CSRF_TOKEN_KEY":"(.*?)"', dashboard_resp)
            if csrf_token_key_match:
                csrf_token_key = csrf_token_key_match.group(1)

            print("csrf_token", csrf_token)
            print("csrf_token_key", csrf_token_key)
            print()
            break

    return csrf_token_key, csrf_token, session


def prepare_variables(switchAccount_resp, table_sb_acc_id):
    dashboard_session_id_match = re.search(r'data-dashboardSessionId\s*=\s*"(.*?)"', switchAccount_resp)
    sb_acc_id_match = re.search(r'(?<=ACCOUNT_ID":").*?(?=")', switchAccount_resp)
    sb_user_id_match = re.search(r'(?<=USER_ID":").*?(?=")', switchAccount_resp)
    data_token_key_match = re.search(r'(?<=data-tokenKey=").*?(?=")', switchAccount_resp)
    data_token_value_match = re.search(r'(?<=data-token=").*?(?=")', switchAccount_resp)

    dashboard_session_id = dashboard_session_id_match.group(1) if dashboard_session_id_match else None
    sb_acc_id = sb_acc_id_match.group(0) if sb_acc_id_match else None
    sb_user_id = sb_user_id_match.group(0) if sb_user_id_match else None
    data_token_key = data_token_key_match.group(0) if data_token_key_match else None
    data_token_value = data_token_value_match.group(0) if data_token_value_match else None

    print("dashboard_session_id:", dashboard_session_id)
    print("sb_acc_id:", sb_acc_id)
    print("sb_user_id:", sb_user_id)
    print("data_token_key:", data_token_key)
    print("data_token_value:", data_token_value)

    change_flag = False
    if table_sb_acc_id != sb_acc_id:
        print("Could not change an account!")

    else:
        print("Account changed successfully!")
        change_flag = True

    return dashboard_session_id, sb_acc_id, sb_user_id, data_token_key, data_token_value, change_flag


def export(session, sb_acc_id, sb_user_id, func_period):
    print("Start entriesExport")
    url = "https://app.sellerboard.com/en/export/createReport"

    headers = {
        'User-Agent': user_agent,
        'Sellerboard-Account-Id': sb_acc_id,
        'Sellerboard-User-Id': sb_user_id,
        'X-Requested-With': 'XMLHttpRequest',
        "origin": "https://app.sellerboard.com",
        "pragma": "no-cache",
        "sec-ch-ua": '"Chromium";v="128", "Not;A=Brand";v="24", "Opera";v="114"',
        "sec-ch-ua-mobile": "?0",
        "sec-ch-ua-platform": '"Windows"',
        "sec-fetch-dest": "empty",
        "sec-fetch-mode": "cors",
        "sec-fetch-site": "same-origin",
        "accept": "*/*",
        "accept-encoding": "gzip, deflate, br, zstd",
        "accept-language": "ru-RU,ru;q=0.9,en-US;q=0.8,en;q=0.7",
        "cache-control": "no-cache",
        "content-type": "application/x-www-form-urlencoded; charset=UTF-8",
        "priority": "u=1, i"
    }

    session.headers.update(headers)

    # payload = {
    #     f'{data_token_key}': data_token_value,
    #
    #     "type": "stock",
    #
    # }

    start_unix_timestamp, end_unix_timestamp = func_period()

    payload = {
        'from': start_unix_timestamp,
        'to': end_unix_timestamp,
        'type': 'DashboardGoods',
        'format': 'xls',
        'name': 'Dashboard by product',
        'withDetails': 'false'
    }

    entries_resp = None

    for i in range(6):

        response = session.post(url, data=payload)

        if response.status_code == 200:

            entries_resp = response.json()

            # print(entries_resp)

            if entries_resp == '':
                continue

            return session, entries_resp

        else:
            print(f"{response.status_code} in entriesExport")

            time.sleep(5)

    return session, entries_resp


def fill_start_time_to_table(spreadsheet_id, start_col_num):
    sheet_name = "Acc"
    id_table = '1xVzaBNyUO0fk_Ue712hWrrfPU6sQljkpbnjVdZCGiKs'  # for working
    # id_table = '14pgfr20kfwhCY_h2Xocn1uVks8LTmE0YLbsibQ8INzI' # for testing
    t = Table(id_table)

    # sheet = t.read_range(sheet_name, id_table)
    # print("sheet", sheet.get_all_values())

    data = t.read_range(sheet_name, id_table)
    # all_rows = data.get_all_values()
    all_rows = t.try_table_operation(lambda: data.get_all_values())
    # print(len(all_rows))

    for i in range(len(all_rows)):
        row = all_rows[i]

        if spreadsheet_id in row[2]:
            current_date_time = datetime.now().strftime("%d.%m.%Y %H:%M")
            t.try_table_operation(lambda: data.update_cell(i + 1, start_col_num, current_date_time))


def get_spreadsheet_id(url):
    match = re.search(r'/d/([a-zA-Z0-9-_]+)', url)
    if match:
        return match.group(1)
    else:
        return None


def ttl_len_rows(sheet_name, spreadsheet_id):
    sheet_name = "SellerBoard"

    # t = Table('14pgfr20kfwhCY_h2Xocn1uVks8LTmE0YLbsibQ8INzI') # table for testing
    t = Table(spreadsheet_id)  # table for working

    sheet = t.read_range(sheet_name, spreadsheet_id)

    ttl_rows = t.try_table_operation(lambda: sheet.get_all_values())
    print("ttl_rows ", len(ttl_rows))

    return len(ttl_rows)


def report_status(report, session):
    print()
    print("Start export status")

    for attempt in range(1, 31):
        print("attempt", attempt)

        url = "https://app.sellerboard.com/en/export/reportStatus"

        payload = {
            "format": "xls",
            "report": report,
            "type": "DashboardGoods",

        }

        response_post = session.post(url, data=payload, timeout=(60, 120))

        if response_post.status_code == 200:

            try:
                response_data = response_post.json()

                if response_data:

                    try:

                        if response_data["report_status"] == "finished":
                            # print(response_data["report_status"])
                            print()
                            return True, session

                        else:
                            print(response_data["report_status"])
                            time.sleep(5)
                            continue

                    except KeyError:
                        print()
                        print("Key 'report_status' not found in response_data. Retrying...")
                        time.sleep(5)
                        continue

                else:
                    # print(response_data["report_status"])
                    time.sleep(5)
                    continue
            except requests.exceptions.JSONDecodeError:
                print()
                print("ERROR JSON")
                time.sleep(3)

        else:
            print(response_post.status_code)
            time.sleep(3)
            continue

        time.sleep(3)

    return False, session


def export_report_download(report, session):
    print()
    print("Start export report download")

    for attempt in range(1, 10):

        print("attempt", attempt)
        url = "https://app.sellerboard.com/en/export/reportDownload"

        # payload = {
        #     f'{data_token_key}': data_token_value,
        #     'dashboardSessionId': dashboard_session_id,
        #     'task_id': task_id,
        #     'format': 'xls',
        #     'entryType': 'product'
        # }

        payload = {
            "format": "xls",
            "id": report,
            "type": "DashboardGoods",

        }

        response_post = session.post(url, data=payload, timeout=(30, 120))

        if response_post.status_code == 200:

            try:
                # content_type = response_post.headers.get('Content-Type')
                # print("content_type: ", content_type)
                print()
                response_data = response_post.json()
                print(response_data)

                if response_data:

                    redirect = response_data.get("redirect")
                    message = response_data.get("message")

                    if redirect:
                        return response_data
                    elif message == 'Report file is missing':
                        print("Message from server: Report file is missing")
                        continue

                return response_data

            except requests.exceptions.JSONDecodeError:

                print("Failed to decode JSON. Response from server:")

                response_data = None

        else:
            print(f"Error: status {response_post.status_code}")


def read_accounts(file_path):
    accounts = {}

    # file_path = "sellerboard-accounts.txt"
    with open(file_path, 'r', encoding='utf-8') as file:
        for line in file:
            parts = line.strip().split(';')
            if len(parts) == 3:
                acc_name = parts[0].strip()
                table_sb_acc_id = parts[1].strip()
                url = parts[2].strip()
                accounts[acc_name] = {
                    "table_sb_acc_id": table_sb_acc_id,
                    "url": url
                }

    print(accounts)
    return accounts


def get_data_from_download(json_res, session, sheet_name, spreadsheet_id, delete_rows_func):
    all_accounts_sheet_name = "Acc"
    all_accounts_id_table = os.getenv("ACCOUNTS_SPREADSHEET_ID")
    all_accounts_table = Table(all_accounts_id_table)
    all_accounts_sheet = all_accounts_table.read_range(all_accounts_sheet_name)

    def fill_table(all_rows):

        def fill_end_time_to_table():
            end_col_num = 7

            all_accounts_rows = all_accounts_table.try_table_operation(lambda: all_accounts_sheet.get_all_values())

            # print(len(all_rows))
            spreadsheet_id_col = 2

            for row_i, row in enumerate(all_accounts_rows):

                if spreadsheet_id in row[spreadsheet_id_col]:
                    current_date_time = datetime.now().strftime("%d.%m.%Y %H:%M")
                    all_accounts_table.try_table_operation(
                        lambda: all_accounts_sheet.update_cell(row_i + 1, end_col_num, current_date_time))
                    break

        sheet_name = "SellerBoard"

        print("START FILL TABLE")

        t = Table(spreadsheet_id)  # table for working
        sheet = t.read_range(sheet_name, spreadsheet_id)

        ttl_rows = t.try_table_operation(lambda: sheet.get_all_values())

        if len(ttl_rows) == 1 and ttl_rows == [[]]:
            start_row = len(ttl_rows)
        else:
            start_row = len(ttl_rows) + 1

        # all_rows.extend(t.empty_last_rows(sheet, all_rows))

        total_rows_in_sheet = sheet.row_count

        empty_rows = total_rows_in_sheet - len(ttl_rows)

        add_empty_rows = max(0, len(all_rows) - empty_rows)

        t.try_table_operation(lambda: sheet.add_rows(add_empty_rows + 1))

        t.update_range(sheet, all_rows, start_row)

        fill_end_time_to_table()

    def get_filtered_rows():
        filtered_rows = []

        start = check_if_empty_sheet(spreadsheet_id)  # not in products
        headers_added = False
        if start != 0:
            delete_rows_func(spreadsheet_id)

        len_all_rows = ttl_len_rows(sheet_name, spreadsheet_id)

        for row_i, row in enumerate(sheet_rows[start:]):
            if not headers_added and start == 0:  # add headers if empty sheet
                month_number = "Month number"
                formatted_date = " Current date"
                headers_added = True

            else:
                current_date = datetime.now()
                month_number = datetime.strptime(row[0], "%m/%d/%Y").month
                # formatted_date = current_date.strftime("%d/%m/%Y")
                formatted_date = f"{current_date.day}/{current_date.month}/{current_date.year}"
                formatted_date = f"=LET(x;SPLIT(C{row_i + len_all_rows + 1};\"/\");DATE(INDEX(x;3);INDEX(x;1);INDEX(x;2)))"

            for i in range(5, len(row)):
                if row[i] == '':
                    row[i] = "0"

            filtered_rows.append([month_number] + [formatted_date] + row)
        # print("filtered_rows", filtered_rows)
        return filtered_rows

    for i in range(5):

        url = json_res['redirect']
        response_get = session.get(url)

        if response_get.status_code == 200:
            content_type = response_get.headers.get('Content-Type')

            print('Successfully downloaded')
            sheet_rows = []
            if 'application/vnd.ms-excel' in content_type:
                workbook = xlrd.open_workbook(file_contents=response_get.content)
                sheet = workbook.sheet_by_index(0)
                sheet_rows = [sheet.row_values(row_i) for row_i in range(sheet.nrows)]

            elif 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' in content_type:
                print("\nXLSX FILE FOUND\n")
                workbook = openpyxl.load_workbook(BytesIO(response_get.content))
                sheet = workbook.active  # Get first sheet
                sheet_rows = [list(row) for row in sheet.iter_rows(values_only=True)]

            if sheet_rows:
                filtered_rows = get_filtered_rows()

                fill_table(filtered_rows)

                update_local_file(spreadsheet_id)
                return 1



        else:
            print("Status code", response_get.status_code)
        time.sleep(3)


def check_if_empty_sheet(spreadsheet_id):
    sheet_name = "SellerBoard"

    # spreadsheet_id = '14pgfr20kfwhCY_h2Xocn1uVks8LTmE0YLbsibQ8INzI' # table for testing

    t = Table(spreadsheet_id)  # table for working
    sheet = t.read_range(sheet_name, spreadsheet_id)

    # all_rows = data.get_all_values()
    all_rows = t.try_table_operation(lambda: sheet.get_all_values())
    # print(len(all_rows))

    if len(all_rows) == 1 and all_rows == [[]]:
        return 0
    else:
        return 1


def main_frame(base_dir, cookies_file, delete_rows_func, export_period_func, time_difference_hours: int = 12,
               difference_days: int = 1):
    acc_count_ready = 0
    start_col_num = 6

    base_dir = os.path.normpath(os.path.join(os.path.dirname(__file__), base_dir))
    cookies_dict = start_log_in(os.path.join(base_dir, cookies_file))
    accounts = read_accounts(os.path.join(base_dir, "sellerboard-accounts.txt"))
    spreadsheet_ids_to_update = find_not_actual_accounts(time_difference_hours, difference_days)

    for acc_name, details in accounts.items():
        if acc_name in ["ALONA _NASALSKA", "Oleksandr Mikhailov"]:
            continue

        table_sb_acc_id = details["table_sb_acc_id"]
        url = details["url"]

        if not any(spreadsheet_id in url for spreadsheet_id in spreadsheet_ids_to_update):
            continue

        print(f"Account Name: {acc_name}")
        print(f"Table SB Account ID: {details}")
        print(f"URL: {url}")
        spreadsheet_id = get_spreadsheet_id(url)
        print("spreadsheet_id: ", spreadsheet_id)
        print("=" * 50)

        fill_start_time_to_table(spreadsheet_id, start_col_num)
        print(f"Get account {acc_name}")

        sheet_name = "SellerBoard"
        count_switch = 0

        session = requests.Session()
        session.cookies.update(cookies_dict)

        data_token_key, data_token_value, session = products(count_switch, session)

        for i in range(6):
            switchAccount_resp = switch_account(table_sb_acc_id, data_token_key, data_token_value, session)
            dashboard_session_id, sb_acc_id, sb_user_id, data_token_key, data_token_value, change_flag = prepare_variables(
                switchAccount_resp, table_sb_acc_id
            )
            if change_flag:
                break
            else:
                time.sleep(10)

        if not change_flag:
            continue  # switch to next account

        print(f"SWITCH ACCOUNT {acc_name}")

        attempts = 26
        for attempt in range(1, attempts):
            print("attempt :", attempt)
            try:
                session, entries_resp = export(session, sb_acc_id, sb_user_id, export_period_func)

                if "report" in entries_resp:
                    report = entries_resp["report"]
                    print(report)

                    flag_status, session = report_status(report, session)
                    if flag_status:
                        json_res = export_report_download(report, session)

                        if json_res is not None:
                            cou = get_data_from_download(
                                json_res, session, sheet_name, spreadsheet_id, delete_rows_func
                            )
                            if cou:
                                acc_count_ready += 1
                                print("work accounts count", acc_count_ready)
                                print()
                            break
                    else:
                        print("flag_status false, trying again ")
                        continue
                time.sleep(5)

            except requests.exceptions.RequestException as e:
                print(f"Attempt {attempt} of {attempts} failed: ConnectionError - {e}")
                time.sleep(5)
                continue

    print(f"{len(accounts)} accounts from sellerboard-accounts.txt")
    print("Work accounts count", acc_count_ready)


def download_entries_export(data_token_key, data_token_value, dashboard_session_id, task_id, session):
    print()
    print("Start download entries export")

    for attempt in range(1, 10):

        print("attempt", attempt)
        url = "https://app.sellerboard.com/en/dashboard/downloadEntriesExport"

        payload = {
            f'{data_token_key}': data_token_value,
            'dashboardSessionId': dashboard_session_id,
            'task_id': task_id,
            'format': 'xls',
            'entryType': 'product'
        }

        response_post = session.post(url, data=payload, timeout=(30, 120))

        if response_post.status_code == 200:

            try:
                # content_type = response_post.headers.get('Content-Type')
                # print("content_type: ", content_type)
                print()
                response_data = response_post.json()
                print(response_data)

                if response_data:

                    redirect = response_data.get("redirect")
                    message = response_data.get("message")

                    if redirect:
                        return response_data
                    elif message == 'Report file is missing':
                        print("Message from server: Report file is missing")
                        continue

                return response_data

            except requests.exceptions.JSONDecodeError:

                print("Failed to decode JSON. Response from server:")

                response_data = None

        else:
            print(f"Error: status {response_post.status_code}")


def manage_products_or_planner(base_dir, cookies_file,
                               products_func, entriesExport_func, export_status_func,
                               get_data_from_download_func, sheet_name, time_difference_hours: int = 12,
                               difference_days: int = 1):
    acc_count_ready = 0

    base_dir = os.path.normpath(os.path.join(os.path.dirname(__file__), base_dir))
    cookies_dict = start_log_in(os.path.join(base_dir, cookies_file))
    accounts = read_accounts(os.path.join(base_dir, "sellerboard-accounts.txt"))

    spreadsheet_ids_to_update = find_not_actual_accounts(time_difference_hours, difference_days)

    for acc_name, details in accounts.items():
        if acc_name in ["ALONA _NASALSKA", "Oleksandr Mikhailov"]:
            continue
        table_sb_acc_id = details["table_sb_acc_id"]
        url = details["url"]
        if not any(spreadsheet_id in url for spreadsheet_id in spreadsheet_ids_to_update):
            continue
        print(f"Account Name: {acc_name}")
        print(f"Table SB Account ID: {details}")
        print(f"URL: {url}")
        spreadsheet_id = get_spreadsheet_id(url)
        print("Account's spreadsheet_id: ", spreadsheet_id)
        print("=" * 50)

        print(f"Get account {acc_name}")

        count_switch = 0

        session = requests.Session()
        session.cookies.update(cookies_dict)

        data_token_key, data_token_value, session = products_func(count_switch, session)

        for i in range(6):

            switchAccount_resp = switch_account(table_sb_acc_id, data_token_key, data_token_value, session)

            dashboard_session_id, sb_acc_id, sb_user_id, data_token_key, data_token_value, change_flag = prepare_variables(
                switchAccount_resp, table_sb_acc_id)

            if change_flag:
                break
            else:
                time.sleep(10)
        if not change_flag:
            continue  # switch to next account

        print(f"SWITCH ACCOUNT {acc_name}")

        attempts = 26

        for attempt in range(1, attempts):
            print("attempt :", attempt)
            print()

            try:

                session, entries_resp = entriesExport_func(session, data_token_key, data_token_value, sb_acc_id,
                                                           sb_user_id)

                print("entries_resp ", entries_resp)

                if entries_resp is not None and entries_resp['task_status'] == 'pending':

                    task_id = entries_resp["task_id"]

                    flag_status, session = export_status_func(data_token_key, data_token_value, task_id, session)

                    if flag_status:
                        json_res = download_entries_export(data_token_key, data_token_value, dashboard_session_id,
                                                           task_id, session)

                        if json_res is not None:

                            cou = get_data_from_download_func(json_res, session, sheet_name, spreadsheet_id)

                            if cou:
                                acc_count_ready += 1
                                print("work account count", acc_count_ready)
                                print()
                            break

                    else:
                        print()
                        print("flag_status false, trying again ")
                        continue
                time.sleep(5)

            except requests.exceptions.RequestException as e:
                print(f"Attempt {attempt} of {attempts} failed: ConnectionError - {e}")
                # logger.exception(f"Account {acc_name} returned exception after {attempt} attempt(s)")

                time.sleep(5)  # Delay between attempts for network stability
                continue

    print("Work account count", acc_count_ready)


def read_json_atomic(file_path, max_wait_time_fl: float = 10):
    max_wait_time = round(max_wait_time_fl)
    # print(f"read_json_atomic start {max_wait_time_fl} {file_path}")
    count = 0
    shall_continue = True
    while shall_continue:
        try:
            with open(file_path, 'r+') as editable_file:
                msvcrt.locking(editable_file.fileno(), msvcrt.LK_NBLCK, 1)

                json_value = json.load(editable_file)

                return json_value

        except (OSError, BlockingIOError):
            if max_wait_time_fl != -1:
                count += 1
                shall_continue = count < max_wait_time
            time.sleep(1)

    # print(f"read_json_atomic finish {max_wait_time} {file_path}")


def update_file_atomic(file_path, data, max_wait_time_fl: float = 10):
    max_wait_time = round(max_wait_time_fl)
    # print(f"update_file_atomic start {len(file_path)} {data}")
    count = 0
    shall_continue = True
    while shall_continue:
        try:
            with open(file_path, 'r+') as editable_file:
                msvcrt.locking(editable_file.fileno(), msvcrt.LK_NBLCK, 1)

                editable_file.seek(0)
                json.dump(data, editable_file, ensure_ascii=False, indent=4)
                editable_file.truncate()
                # print(f"update_file_atomic return {len(file_path)} {data}")

            return

        except (OSError, BlockingIOError):
            if max_wait_time != -1:
                count += 1
                shall_continue = count < max_wait_time
            time.sleep(1)


def find_not_actual_accounts(time_difference_hours: int = 12, difference_days: int = 1) -> list:
    time_difference_seconds = time_difference_hours * 60 * 60
    updating_local_file_name = "accounts_updated.json"
    accounts_local_updatings_json = read_json_atomic(updating_local_file_name)

    current_datetime = datetime.now()
    not_actual_spreadsheet_ids = []
    for spreadsheet_id, date_update in accounts_local_updatings_json.items():
        time_delta = current_datetime - datetime.strptime(date_update, "%d.%m.%Y %H:%M")

        # If the date of update was more than time_difference_hours ago and more than 1 day ago
        if time_delta.seconds > time_difference_seconds or time_delta.days >= difference_days:
            not_actual_spreadsheet_ids.append(spreadsheet_id)

    return not_actual_spreadsheet_ids


def update_local_file(spreadsheet_id):
    updating_local_file_name = "accounts_updated.json"
    accounts_local_updatings_json = read_json_atomic(updating_local_file_name)
    accounts_local_updatings_json.update({
        spreadsheet_id: datetime.now().strftime("%d.%m.%Y %H:%M")
    })
    update_file_atomic(updating_local_file_name, accounts_local_updatings_json)
