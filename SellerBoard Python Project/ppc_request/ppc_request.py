try:
    import openpyxl
    from io import BytesIO
    import sys
    import os
    import traceback

    sys.path.append(
        "C:\\Program Files\\ZennoLab\\RU\\ZennoPoster Pro V7\\7.7.0.0\\Progs\\Projects\\AMZ Professional\\SellerBoard Python Project\\utilities")
    sys.path.append(
        "D:\\PyCharm Projects\\SellerBoard Python Project\\utilities")

    from gspread.exceptions import APIError
    import time
    import requests
    import re
    import json
    import os
    from datetime import datetime, timedelta, timezone
    from dateutil.relativedelta import relativedelta

    file_path = "sellerboard-accounts.txt"
    import xlrd  # intstall
    from tables import Table
    from log_in import start_log_in

    from sellerboard_interaction import (switch_account, prepare_variables,
                                         download_entries_export,
                                         user_agent)
except Exception as e:
    traceback.print_exc()
    time.sleep(5)
    raise e
pages_counter = 1

t = Table(str(os.getenv("PPC_DATA_SPREADSHEET_ID")))


def main():
    def delete_sheets_from_table():
        spreadsheet = t.spreadsheet
        worksheets = spreadsheet.worksheets()
        accounts = read_accounts()

        for acc_name, details in accounts.items():
            # print(f"Account Name: {acc_name}")

            for worksheet in worksheets:

                if worksheet.title.startswith(acc_name) and "ACoS" not in worksheet.title:
                    found = False
                    # time.sleep(3)
                    # t.try_table_operation(lambda: t.spreadsheet.add_worksheet(title=new_sheet_name_2, rows=100, cols=10))
                    try:
                        t.try_table_operation(lambda: spreadsheet.del_worksheet(worksheet))
                        # spreadsheet.del_worksheet(worksheet)
                        print(f"Sheet '{worksheet.title}' was successfully deleted.")
                        found = True
                    except APIError:
                        pass
                    if not found:
                        print(f"Sheets starting with '{acc_name}' were not found.")

    def read_accounts():
        accounts = {}

        file_path = "sellerboard-accounts.txt"
        with open(file_path, 'r', encoding='utf-8') as file_lines:
            for line in file_lines:
                parts = line.strip().split(';')
                if len(parts) == 2:

                    acc_name = parts[0].strip()
                    table_sb_acc_id = parts[1].strip()
                    accounts[acc_name] = {
                        "table_sb_acc_id": table_sb_acc_id,
                        # "url": url
                    }

        print(accounts)
        return accounts

    def sheet_names():
        current_date = datetime.now()

        current_month = current_date.month

        previous_month = current_month - 1 if current_month > 1 else 12
        previouss_month = current_month - 2 if current_month > 2 else 11

        first_date_range = f"26.{previouss_month:02}-25.{previous_month:02}"
        second_date_range = f"26.{previous_month:02}-25.{current_month:02}"

        new_sheet_name_1 = f"{acc_name} {first_date_range}"
        new_sheet_name_2 = f"{acc_name} {second_date_range}"

        months_objs = {
            0: new_sheet_name_1,
            1: new_sheet_name_2
        }

        return months_objs

    def dashboard():
        data_token_key, data_token_value = None, None
        dashboard_start_attempts = 2
        for _ in range(dashboard_start_attempts):
            dashboard_url = "https://app.sellerboard.com/en/dashboard/"

            headers = {'Referer': dashboard_url}
            dashboard_response = session.get(dashboard_url, headers=headers)
            dashboard_resp = dashboard_response.text

            if dashboard_response.status_code == 200:
                print("Successfully received the dashboard page!")

                data_token_key_match = re.search(r'(?<=data-tokenKey=").*?(?=")', dashboard_resp)
                data_token_value_match = re.search(r'(?<=data-token=").*?(?=")', dashboard_resp)

                data_token_key = data_token_key_match.group(0) if data_token_key_match else None
                data_token_value = data_token_value_match.group(0) if data_token_value_match else None

                break
            time.sleep(2)

        return data_token_key, data_token_value, session

    def unix_time_variables():
        current_date = datetime.now()
        
        month = (current_date.month + month_index - 2) % 12 + 1
        year = current_date.year + (current_date.month + month_index - 2) // 12
        fake_current_date = datetime(year, month, 27)

        formatted_fake_current_date = fake_current_date.strftime("%d.%m.%Y")

        end_of_today = datetime(current_date.year, current_date.month, current_date.day, 23, 59, 59)

        unix_time_end_of_today = int((end_of_today - datetime(1970, 1, 1)).total_seconds())
        today_23_59_59 = str(unix_time_end_of_today)

        previous_month_26 = (fake_current_date.replace(day=1) - timedelta(days=1)).replace(day=26)

        if isinstance(previous_month_26, datetime):

            previous_month_26_year_ago = previous_month_26.replace(year=previous_month_26.year - 1)
        else:
            previous_month_26 = datetime.strptime(previous_month_26, "%Y-%m-%d %H:%M:%S")
            previous_month_26_year_ago = previous_month_26.replace(year=previous_month_26.year - 1)

        unix_time_prev_month = int((previous_month_26 - datetime(1970, 1, 1)).total_seconds())
        previous_month_26 = str(unix_time_prev_month)

        unix_time_prev_month_year_ago = int((previous_month_26_year_ago - datetime(1970, 1, 1)).total_seconds())
        one_year_ago_26 = str(unix_time_prev_month_year_ago)

        this_month_25 = datetime(fake_current_date.year, fake_current_date.month, 25, 23, 59, 59)
        unix_time_this_month = int((this_month_25 - datetime(1970, 1, 1)).total_seconds())
        this_month_25_23_59_59 = str(unix_time_this_month)

        this_month_25_year_ago = this_month_25.replace(year=this_month_25.year - 1)
        unix_time_this_month_year_ago = int((this_month_25_year_ago - datetime(1970, 1, 1)).total_seconds())
        one_year_ago_25 = str(unix_time_this_month_year_ago)

        one_year_ago = current_date.replace(year=current_date.year - 1, hour=0, minute=0, second=0, microsecond=0)

        one_year_ago_unix_time = int((one_year_ago - datetime(1970, 1, 1)).total_seconds())
        one_year_ago = str(one_year_ago_unix_time)

        today_23_59_59 = int(today_23_59_59)
        previous_month_26 = int(previous_month_26)
        one_year_ago_26 = int(one_year_ago_26)
        this_month_25_23_59_59 = int(this_month_25_23_59_59)
        one_year_ago_25 = int(one_year_ago_25)
        one_year_ago = int(one_year_ago)

        return today_23_59_59, previous_month_26, one_year_ago_26, this_month_25_23_59_59, one_year_ago_25, one_year_ago

    def unix_update():
        # Current date and time inUTC
        current_time = datetime.now(timezone.utc)

        # 1. The date is exactly three months ago with the time 00:00:00 in UTC
        # three_months_ago = current_time - relativedelta(months=3)
        # three_months_ago_zero = datetime(
        #     three_months_ago.year, three_months_ago.month, three_months_ago.day, 0, 0, 0, tzinfo=timezone.utc
        # )

        # Convert to Unix time format (in seconds)
        # unix_time_three_months_ago = int(three_months_ago_zero.timestamp())

        # We save the value into a variable
        # three_months_ago_value = unix_time_three_months_ago
        # three_months_ago = unix_time_three_months_ago

        # 2. Current time in Unix format (seconds and milliseconds)
        unix_now_secs = int(current_time.timestamp())
        unix_now_milliseconds = int(current_time.timestamp() * 1000)

        # We save values into variables
        # unix_now_secs_value = unix_now_secs
        # unix_now_milliseconds_value = unix_now_milliseconds

        # print("three_months_ago:", three_months_ago_value)  # Value three months ago (in seconds)
        # print("unix_now_secs:", unix_now_secs_value)  # Current time in seconds
        # print("unix_now_milliseconds:", unix_now_milliseconds_value)  # Current time in milliseconds

        return unix_now_secs, unix_now_milliseconds
    
    def entries():
        entries_resp = None

        print("\nStart Entries")
        entries_get_attempts = 10
        for _ in range(entries_get_attempts):

            url = "https://app.sellerboard.com/en/dashboard/entries/"
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

            unix_now_secs, unix_now_milliseconds = unix_update()

            payload = {
                '_': unix_now_milliseconds,
                'dashboardSessionId': dashboard_session_id,
                "viewType": "table",
                "entryType": "product",
                "periodicity": "period",
                "sortField": "units",
                "sortDirection": "desc",
                "page": pages_counter,
                "groupByAsin": "",
                "groupBy": "",
                "trendsParameter": "sales",
                "sortingPeriodStart": "",
                "sortingPeriodEnd": "",
                "mode": "dashboard",
                "periodStart": previous_month_26,  # 1724630400,
                "periodEnd": this_month_25_23_59_59,  # 1727308799
                "market": "",
                "mapDataFilter": "sales",
                "compare": "false",
                "collapseEntriesByAsin": "false",

            }

            response_post = session.post(url, data=payload, timeout=(30, 120))
            if response_post.status_code == 200:

                try:
                    entries_resp = response_post.json()
                    print(entries_resp["state"])
                    break

                except requests.exceptions.JSONDecodeError:
                    print("\nERROR JSON")
                    time.sleep(5)

            else:
                print(f"Error status {response_post.status_code}")
                time.sleep(5)

        return session, entries_resp

    def periods():
        print("Start periods")
        
        update_periods_attempts = 10
        for _ in range(update_periods_attempts):

            url = "https://app.sellerboard.com/en/dashboard/periods"
            unix_now_secs, unix_now_milliseconds = unix_update()
            payload = {
                "_": unix_now_milliseconds,
                "dashboardSessionId": dashboard_session_id,
                "startedAt": f"{unix_now_secs}:ad47b09a28a558d6c949c3c2f83797e3",
                "viewType": "panels",
                "entryType": "product",
                "periodicity": "period",
                "sortField": "units",
                "sortDirection": "desc",
                "page": 1,
                "groupByAsin": "",
                "groupBy": "",
                "trendsParameter": "sales",
                "sortingPeriodStart": "",
                "sortingPeriodEnd": "",
                "mode": "dashboard",
                "mapDataFilter": "sales",
                "loadBy": "periods",
                "market": "",
                "periodsPreset": f"custom-range:c|{previous_month_26}-{this_month_25_23_59_59}*{one_year_ago_26}-{one_year_ago_25}",
                "periods[0][start]": previous_month_26,  # 1724630400,
                "periods[0][end]": this_month_25_23_59_59,  # 1727308799
                "periods[0][compare_to_start]": one_year_ago_26,  # 1693008000,
                "periods[0][compare_to_end]": one_year_ago_25,  # 1695686399,
                "periods[0][forecast]": 0,
                "periods[0][standard]": 0,
                "periods[0][key]": "custom",
                "compare": "false",
                "collapseEntriesByAsin": "false",

            }

            response = session.post(url, data=payload, timeout=(30, 120))
            if response.status_code == 200:
                print("PERIODS CHANGED")
                entries_resp = response.json()
                # print(entries_resp)

                time.sleep(5)
                break

            else:
                print(f"Error status {response.status_code}")
                time.sleep(5)

        return session

    def entries_export():
        print("Start entries_export")
        url = "https://app.sellerboard.com/en/dashboard/entriesExport"

        payload = {
            f'{data_token_key}': data_token_value,
            'dashboardSessionId': dashboard_session_id,
            "viewType": "panels",
            "entryType": "product",
            "periodicity": "period",
            "sortField": "units",
            "sortDirection": "desc",
            "page": pages_counter,
            "groupByAsin": "",
            "groupBy": "",
            "trendsParameter": "sales",
            "sortingPeriodStart": "",
            "sortingPeriodEnd": "",
            "mode": "dashboard",
            "periodStart": previous_month_26,
            "periodEnd": this_month_25_23_59_59,
            "market": "",
            "mapDataFilter": "sales",
            "compare": "",
            "format": "xls",
        }

        entries_resp = None
        entries_export_attempts = 5
        for _ in range(entries_export_attempts):

            response = session.post(url, data=payload)

            if response.status_code == 200:

                entries_resp = response.json()

                if entries_resp['status'] != 'error':
                    break

            else:
                print(f"{response.status_code} in entriesExport")

                time.sleep(5)

        return session, entries_resp

    def export_status():
        print("Start export status")

        for attempt in range(1, 31):
            print("attempt", attempt)

            url = "https://app.sellerboard.com/en/dashboard/entriesExportStatus"

            payload = {
                f'{data_token_key}': data_token_value,
                'dashboardSessionId': dashboard_session_id,
                'task_id': task_id,
                'format': 'xls',
                'entryType': 'product'
            }

            response_post = session.post(url, data=payload, timeout=(60, 120))

            if response_post.status_code == 200:

                try:
                    response_data = response_post.json()
                    # print(response_data)

                    if response_data["report_status"] == "finished":
                        # print(response_data["report_status"])
                        print()
                        return True, session

                    else:
                        print(response_data["report_status"])
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

    def export_from_download():
        def read_workbook():
            current_month_rows = []
            if 'application/vnd.ms-excel' in content_type:
                print("\nvnd.ms-excel FOUND\n")

                workbook = xlrd.open_workbook(file_contents=response_get.content)
                sheet = workbook.sheet_by_index(0)
                current_month_rows = [sheet.row_values(row) for row in range(sheet.nrows)]

            elif 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' in content_type:
                print("\nXLSX FILE FOUND\n")
                workbook = openpyxl.load_workbook(BytesIO(response_get.content))
                sheet = workbook.active  # Get first sheet
                current_month_rows = [list(row) for row in sheet.iter_rows(values_only=True)]
            return current_month_rows

        def fill_sheet():
            new_sheet_name = months_objs[month_index]
            sheet = t.try_table_operation(lambda: t.spreadsheet.add_worksheet(title=new_sheet_name, rows=100, cols=10))

            if month_index == 0:
                t.update_range(sheet, current_month_rows, start_row=1)
            else:
                def adding_last_sheet_computations():
                    def calculate_net_profit_difference(u_value, ab_value):
                        if not u_value and not ab_value:
                            return 0

                        if not u_value:

                            if isinstance(ab_value, str):
                                ab_value_numeric = float('-' + ab_value)
                                return ab_value_numeric
                            else:
                                return abs(ab_value)

                        else:
                            u_value_numeric = float(u_value)

                        if not ab_value:
                            return u_value_numeric
                        else:
                            ab_value_numeric = float(ab_value)

                        return round(u_value_numeric - ab_value_numeric, 2)

                    def calculate_profit_points(value):
                        if not ad_spend or value < 100:
                            return 0
                        elif 100 <= value <= 250:
                            return 1
                        elif 250 < value <= 500:
                            return 2
                        elif 500 < value <= 750:
                            return 3
                        elif 750 < value <= 1000:
                            return 4
                        elif 1000 < value <= 1500:
                            return 5
                        elif 1500 < value <= 2000:
                            return 10
                        elif 2000 < value:
                            return 20

                    print("Preparing data for this month")
                    asin_col = 1
                    sku_col = 2
                    ad_spend_col = 7
                    net_profit_col = 20
                    for second_month_row in current_month_rows[1:]:
                        asin = second_month_row[asin_col]
                        sku = second_month_row[sku_col]

                        for first_month_row in first_month_rows:

                            if first_month_row[asin_col] == asin and first_month_row[sku_col] == sku:
                                ad_spend = second_month_row[ad_spend_col]
                                net_profit_sec_mon = second_month_row[net_profit_col]

                                first_mon_net_profit = first_month_row[net_profit_col]
                                diff_net_profit_m_m = calculate_net_profit_difference(net_profit_sec_mon,
                                                                                      first_mon_net_profit)

                                profit_point = calculate_profit_points(diff_net_profit_m_m)
                                list_to_end = [first_mon_net_profit, diff_net_profit_m_m, profit_point]

                                second_month_row.extend(list_to_end)
                                break
                        else:
                            print(f"Match for ({asin}, {sku}) not found in all_rows_first_mon.")

                    current_month_rows[0].extend(
                        ['Net profit(previous month)', 'Разница Net Profit m / m', 'Баллы за профит'])

                    return current_month_rows

                updated_second_month_rows = adding_last_sheet_computations()

                t.update_range(sheet, updated_second_month_rows, start_row=1)
                print("Data successfully add for this month")

        for i in range(5):

            url = json_res['redirect']
            response_get = session.get(url)

            if response_get.status_code == 200:
                content_type = response_get.headers.get('Content-Type')
                print("content_type ", content_type)
                print('OK DOWNLOAD')
                current_month_rows = read_workbook()

                if not current_month_rows:
                    continue

                fill_sheet()
                return current_month_rows

            else:
                print("Status code", response_get.status_code)
                time.sleep(3)

    acc_count_ready = 0

    delete_sheets_from_table()

    cookies_dict = start_log_in("cookies.json")

    # accounts = read_accounts(os.path.join(base_dir, "sellerboard-accounts.txt"))
    accounts = read_accounts()

    for acc_name, details in accounts.items():
        table_sb_acc_id = details["table_sb_acc_id"]
        print(f"\nAccount Name: {acc_name}")
        print(f"Table SB Account ID: {details}")
        print("=" * 50)
        print(f"Get account {acc_name}")

        months_objs = sheet_names()

        session = requests.Session()
        session.cookies.update(cookies_dict)

        data_token_key, data_token_value, session = dashboard()

        switch_profile_attempts = 6
        for _ in range(switch_profile_attempts):

            switchAccount_resp = switch_account(table_sb_acc_id, data_token_key, data_token_value, session)

            dashboard_session_id, sb_acc_id, sb_user_id, data_token_key, data_token_value, change_flag = (
                prepare_variables(switchAccount_resp, table_sb_acc_id))

            if change_flag:
                break
            else:
                time.sleep(10)
        
        print(f"SWITCH ACCOUNT {acc_name}")

        first_month_rows = []
        for month_index in range(2):
            print(acc_name, month_index)
            attempts = 26

            for attempt in range(attempts):
                print("Attempt :", attempt + 1)

                try:
                    (today_23_59_59, previous_month_26, one_year_ago_26, this_month_25_23_59_59, one_year_ago_25, 
                     one_year_ago) = unix_time_variables()

                    session, entries_resp = entries()

                    if entries_resp and entries_resp["state"] == "ready":
                        print('\nEntries are ready')

                        session, entries_resp = entries_export()

                        if entries_resp["status"] == "success":
                            print("ENTRIES SUCCESS")

                            task_id = entries_resp["id"]
                            flag_status, session = export_status()
                            
                            if not flag_status:
                                print("flag_status false, trying again ")
                                continue
                            
                            print("FLAG_STATUS")

                            json_res = download_entries_export(data_token_key, data_token_value,
                                                               dashboard_session_id, task_id, session)
                            if json_res:
                                print("JSON_RES")
                                first_month_rows = export_from_download()
                                break                                

                        time.sleep(5)

                    session = periods()

                except requests.exceptions.RequestException as e:
                    print(f"Attempt {attempt + 1} of {attempts} failed: ConnectionError - {e}")
                    time.sleep(5)  # Delay between attempts for network stability
                    continue

    print("Work account count", acc_count_ready)


if __name__ == "__main__":
    big_error = Exception
    try:
        start_time = time.time()

        main()

        end_time = time.time()

        execution_time = end_time - start_time
        print(f"Program execution time: {execution_time:.2f} seconds")

    except Exception as e:
        big_error = e
        print("big_error", str(big_error))
        time.sleep(5)

    if big_error is not Exception:
        raise big_error
