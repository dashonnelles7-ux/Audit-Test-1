try:
    import sys
    import os

    sys.path.append("C:\\Program Files\\ZennoLab\\RU\\ZennoPoster Pro V7\\7.7.0.0\\Progs\\Projects\\AMZ Professional\\SellerBoard Python Project\\utilities")
    sys.path.append(
        "D:\\PyCharm Projects\\SellerBoard Python Project\\utilities")
    import traceback
    import time
    import requests
    from datetime import datetime, timedelta, timezone
    from dateutil.relativedelta import relativedelta
    import re
    from tables import Table
    import json
    import random
    import calendar
    from sellerboard_abc import main as transfer_to_total
    from dotenv import load_dotenv
    from log_in import start_log_in
    from sellerboard_interaction import (read_accounts, user_agent)

    load_dotenv()
except Exception as e:
    traceback.print_exc()
    time.sleep(5)
    raise e


errors = []


def dashboard(count_switch, session, table_sb_acc_id):

    while count_switch != 2:

        dashboard_url = "https://app.sellerboard.com/en/dashboard/"

        headers = {'Referer': 'https://app.sellerboard.com/en/dashboard/'}

        dashboard_response = session.get(dashboard_url, headers=headers)

        dashboard_resp = dashboard_response.text

        if dashboard_response.status_code == 200:

            print("Successfully received the dashboard page!")

            data_token_key_match = re.search(r'(?<=data-tokenKey=").*?(?=")', dashboard_resp)
            data_token_value_match = re.search(r'(?<=data-token=").*?(?=")', dashboard_resp)

            data_token_key = data_token_key_match.group(0) if data_token_key_match else None
            data_token_value = data_token_value_match.group(0) if data_token_value_match else None

            print("data_token_key:", data_token_key)
            print("data_token_value:", data_token_value)
            print()

            # table_sb_acc_id = "f4c3b8ba-07cb-11eb-85dc-e4434b2f1e00"  # account ID
            post_data = {
                "account": table_sb_acc_id,
                data_token_key: data_token_value
            }

            switch_account_url = "https://app.sellerboard.com/en/setup/switchAccount"


            switch_account_response = session.post(switch_account_url,headers=headers, data=post_data)

            switchAccount_resp = switch_account_response.text

            if switch_account_response.status_code == 200:

                print("Successfully switched to account!")

                break
                # print("(switchAccount_resp):", switchAccount_resp)
            else:
                print()
                print("Error switching to account!")
                print("Status code:", switch_account_response.status_code)
                time.sleep(10)

                # response = session.post(url, data=payload)
                # print(response.status_code)

                count_switch += 1
                #print("", switchAccount_resp)
        else:
            print("Error getting dashboard page!")
            print("Status code:", dashboard_response.status_code)
            time.sleep(10)
            switchAccount_resp = None
            count_switch+=1


    return switchAccount_resp, session


def prepare_variables(switchAccount_resp, table_sb_acc_id):

    dashboard_session_id_match = re.search(r'data-dashboardSessionId\s*=\s*"(.*?)"', switchAccount_resp)
    sb_acc_id_match = re.search(r'(?<=ACCOUNT_ID":").*?(?=")', switchAccount_resp)
    sb_user_id_match = re.search(r'(?<=USER_ID":").*?(?=")', switchAccount_resp)
    data_token_key_match = re.search(r'(?<=data-tokenKey=").*?(?=")', switchAccount_resp)
    data_token_value_match = re.search(r'(?<=data-token=").*?(?=")', switchAccount_resp)

    # data_token_key_match = re.search(r'(?<=data-tokenKey=").*?(?=")', dashboard_resp)
    # data_token_value_match = re.search(r'(?<=data-token=").*?(?=")', dashboard_resp)

    dashboard_session_id = dashboard_session_id_match.group(1) if dashboard_session_id_match else None
    sb_acc_id = sb_acc_id_match.group(0) if sb_acc_id_match else None
    sb_user_id = sb_user_id_match.group(0) if sb_user_id_match else None
    data_token_key = data_token_key_match.group(0) if data_token_key_match else None
    data_token_value = data_token_value_match.group(0) if data_token_value_match else None

    print("dashboard_session_id:", dashboard_session_id)
    print("sb_acc_id:", sb_acc_id)
    print("sb_user_id:", sb_user_id)
    # print("data_token_key:", data_token_key)
    # print("data_token_value:", data_token_value)

    if table_sb_acc_id != sb_acc_id:
        print("Could not change an account!")

    else:
        print("Account changed successfully!")

    entries_counter = 0
    pages_counter = 1

    return entries_counter, pages_counter, dashboard_session_id, sb_acc_id, sb_user_id


def unix_update():

    # Current date and time inUTC
    current_time = datetime.now(timezone.utc)

    # Current time in Unix format (seconds and milliseconds)
    unix_now_secs = int(current_time.timestamp())
    unix_now_milliseconds = int(current_time.timestamp() * 1000)

    return unix_now_secs, unix_now_milliseconds


def get_saturday_unix_time():

    current_date = datetime.now(timezone.utc)

    days_until_saturday = (5 - current_date.weekday()) % 7
    next_saturday = current_date + timedelta(days=days_until_saturday)

    next_saturday = next_saturday.replace(hour=23, minute=59, second=59, microsecond=0)

    next_saturday_unix = int(next_saturday.timestamp())

    return next_saturday_unix


def get_last_sunday_three_months_ago():

    current_date = datetime.utcnow()

    three_months_ago = current_date - relativedelta(months=3)

    days_to_last_sunday = (three_months_ago.weekday() + 1) % 7
    last_sunday = three_months_ago - timedelta(days=days_to_last_sunday)

    last_sunday = last_sunday.replace(hour=0, minute=0, second=0, microsecond=0)

    unix_time = calendar.timegm(last_sunday.timetuple())

    return unix_time


def periods(session, dashboard_session_id):

    print("Start periods")

    for i in range(10):

        url = "https://app.sellerboard.com/en/dashboard/periods"

        unix_now_secs, unix_now_milliseconds = unix_update()

        this_saturday = get_saturday_unix_time()
        sunday_3_mon_ago = get_last_sunday_three_months_ago()

        payload = {
            "_": unix_now_milliseconds,
            "dashboardSessionId": dashboard_session_id,
            "startedAt": f"{unix_now_secs}:efbfb10c6893425bf261b8466b5c9276",
            "viewType": "table",
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
            "market": "",
            "mapDataFilter": "sales",
            "loadBy": "rangeTable",
            "periods[0][end]": this_saturday,       # 1729382399,
            "periods[0][start]": sunday_3_mon_ago,  # 1720915200,
            "periods[0][forecast]": 0,
            "periods[0][standard]": 0,
            "periods[0][status]": "preparing",
            "periods[0][is_totals]": 1,
            "compare": False,
            "collapseEntriesByAsin": False,
            "periodsPreset": "today_yesterday_forecast"
        }

        response = session.post(url, data=payload, timeout=(30, 120))
        if response.status_code == 200:

            break

        else:
            print(f"Error status {response.status_code}")
            time.sleep(5)

    return session


def entries(sb_acc_id, sb_user_id, session, dashboard_session_id, pages_counter):
    print("Start Entries")


    for i in range(10):

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
            "priority":"u=1, i"
        }

        session.headers.update(headers)

        unix_now_secs, unix_now_milliseconds = unix_update()

        this_saturday = get_saturday_unix_time()
        sunday_3_mon_ago = get_last_sunday_three_months_ago()

        # # Data for POST-response
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
            "periodStart": sunday_3_mon_ago,  #1720915200,
            "periodEnd": this_saturday,       # 1729382399
            "market": "",
            "mapDataFilter": "sales",
            "compare": "false",
            "collapseEntriesByAsin": "false"
        }

        response_post = session.post(url, data=payload, timeout=(30, 120))
        if response_post.status_code == 200:
            try:
                entries_resp = response_post.json()
                break

            except requests.exceptions.JSONDecodeError:
                print()
                print("ERROR JSON")
                time.sleep(5)
                entries_resp = None

        else:
            print(f"Error status {response_post.status_code}")
            time.sleep(5)


    return session, entries_resp


def check_next_page(pages_counter, entries_counter, total_count):

    if (pages_counter - 1) * 50 + entries_counter < total_count:
        entries_counter = 0
        pages_counter += 1
        return entries_counter, pages_counter

    return None, None


def accumulate_table_data(entries_resp, entries_counter, data_table):

    print("totalCount: ", entries_resp["totalCount"])
    total_count = entries_resp["totalCount"]

    for entry in entries_resp['entries']:
        # print(entry)
        # print(f"{entry['asin']}")
        # print(f"{entry['sku']}")

        data_table.append(
            [entry['asin'], entry['sku'], '', entry['units'], '', '', '', entry['refunds'], entry['sales'], '', '', '',
             entry['promotionValue'], entry['advertising'], '', entry['refundCostsTotal'], entry['amazonFeesTotal'],
             entry['productCosts'], entry['netProfit'], '', '', '', entry['margin'], entry['roi']])
        entries_counter += 1


    print("entries_counter", entries_counter)
    #print("data_table ", data_table)
    return data_table, entries_counter, total_count


def fill_google_spreadsheet(prepared_data_for_table):

    t = Table(os.getenv("ABC_DATA_SPREADSHEET_ID"))
    sheet_name = os.getenv("ABC_SHEET_NAME")

    sheet = t.read_range(sheet_name)
    t.try_table_operation(lambda: sheet.resize(2)) # Keep only the first two rows by resizing the sheet.
    values = t.try_table_operation(lambda: sheet.get_all_values())

    #print(prepared_data_for_table)
    rows_count = len(prepared_data_for_table) + 1
    update_rows = [{"range": f"{rows_count}:2", "values": prepared_data_for_table}]

    result = update_rows

    print(result[0]['values'])
    result = result[0]['values']

    result.extend(t.empty_last_rows(sheet, result))

    total_rows_in_sheet = sheet.row_count

    empty_rows = total_rows_in_sheet - len(values)

    add_empty_rows = max(0, len(result) - empty_rows)

    t.try_table_operation(lambda: sheet.add_rows(add_empty_rows+1))

    t.update_range(sheet, result)

def prepare_data(data_table):

    def calculate_totals(data_table):

        total_units = 0
        total_sales = 0.0
        total_profit = 0.0

        for row in data_table:
            # Check for data in the required columns and cast types
            try:
                units = int(row[3]) if row[3] not in ['', None] else 0  # Number of units in column 4 (index 3)
                sales = float(row[8]) if row[8] not in ['', None] else 0.0  # Sales in column 9 (index 8)
                profit = float(row[18]) if row[18] not in ['', None] else 0.0  # Profit in column 19 (index 18)
            except ValueError:
                # If type conversion causes an error, skip this line
                continue

            total_units += units
            total_sales += sales
            total_profit += profit

        total_sales = round(total_sales, 2)
        total_profit = round(total_profit, 2)

        print(f"Total Units: {total_units}")
        print(f"Total Sales: {total_sales}")
        print(f"Total Profit: {total_profit}")

        return total_units, total_sales, total_profit

    def deltas(total_units, total_sales, total_profit, data_table):
        # List of columns to process
        list_col = [3, 8, 18]

        # Processing data for each column from the list
        for edit_column in list_col:
            # Determine total_num based on the selected column
            if edit_column == 3:
                total_num = total_units
            elif edit_column == 8:
                total_num = total_sales
            elif edit_column == 18:
                total_num = total_profit
            else:
                total_num = 0.0

            # Sort data by the current column in reverse order
            data_table = sorted(data_table, key=lambda x: x[edit_column] if x[edit_column] not in ['', None] else 0.0,
                                reverse=True)

            # Initializing a variable for summing deltas
            sum_delta = 0.0

            for row in data_table:

                value = row[edit_column]

                # Checking and converting a value to float
                try:
                    units = float(value) if value not in ['', None] else 0.0
                except ValueError:
                    units = 0.0

                # Calculating delta units (delta_units)
                delta_units = units / total_num if total_num != 0 else 0.0

                # Accumulation of the sum of deltas
                sum_delta += delta_units

                # Definition of category
                if sum_delta < 0.8:
                    category = "A"
                elif sum_delta < 0.95:
                    category = "B"
                else:
                    category = "C"

                # Write results to a table row depending on the column
                if edit_column == 3:
                    row[4] = delta_units
                    row[5] = sum_delta
                    row[6] = category
                elif edit_column == 8:
                    row[9] = delta_units
                    row[10] = sum_delta
                    row[11] = category
                elif edit_column == 18:
                    row[19] = delta_units
                    row[20] = sum_delta
                    row[21] = category

        # for row in data_table:
        #     print(row)

        return data_table

    def calculate_unit_and_user_profit(data_table):

        def extend_row(row, target_length):
            """Expands a string to the desired length."""
            while len(row) < target_length:
                row.append('')  # Add empty values if the string is shorter than the required length

        for row in data_table:
            # Make sure that the string has the required number of elements (at least 26)
            extend_row(row, 26)

            units = float(row[3])
            refunds = float(row[7])
            sales = float(row[8])  # Col "I" (index 8)
            amazon_fees = float(row[16]) * -1  # Col "Q" (index 16)
            np = float(row[18])  # Col "S" (index 18)

            # Calculating the percentage of returns
            perc_refunds = (refunds / units * 100) if units != 0 else 0.0
            row[14] = perc_refunds  # Insert the return percentage into the column 14 (index 14)

            # Calculation of net profit per unit of goods
            np_unit = (np / units) if units != 0 else 0.0
            row[24] = np_unit  # Insert net profit per unit into the column 24 (index 24)

            # Calculating Gross Profit Per Unit
            gross_np_unit = ((sales - amazon_fees) / units) if units != 0 else 0.0
            row[25] = gross_np_unit  # Insert the gross profit per unit into the column 25 (index 25)

        # for row in data_table:
        #     print(row)

        return data_table

    def replace_dots_with_commas(data_table):
        new_data_table = []
        for row in data_table:
            new_row = []
            for item in row:
                if isinstance(item, float):
                    new_row.append(str(item).replace('.', ','))
                else:
                    new_row.append(item)
            new_data_table.append(new_row)
        return new_data_table


    total_units, total_sales, total_profit = calculate_totals(data_table)

    data_table = deltas(total_units, total_sales, total_profit, data_table)

    data_table = calculate_unit_and_user_profit(data_table)

    # add date to column C
    today = datetime.today().strftime("%d.%m.%Y")
    for row in data_table:
        row[2] = today

    data_table = replace_dots_with_commas(data_table)

    print()

    total_units = str(total_units).replace('.', ',')
    total_sales = str(total_sales).replace('.', ',')
    total_profit = str(total_profit).replace('.', ',')

    total_row = ["Total", "", "", total_units, "", "", "", "", total_sales, "", "", "", "", "", "", "", "", "",
                 total_profit]
    data_table.append(total_row)

    return data_table


def main():

    try:
        acc_count_ready = 0
        data_table = []
        miss_accounts = {}

        base_dir = os.path.normpath(os.path.join(os.path.dirname(__file__)))

        cookies_dict = start_log_in(os.path.join(base_dir, "abc_cookies.json"))
        accounts = read_accounts(os.path.join(base_dir, "sellerboard-accounts.txt"))

        for acc_name, details in accounts.items():
            table_sb_acc_id = details["table_sb_acc_id"]
            print(f"Account Name: {acc_name}")
            print(f"Table SB Account ID: {table_sb_acc_id}")
            print(f"URL: {details['url']}")
            print("=" * 50)

            print()

            count_switch = 0

            session = requests.Session()
            session.cookies.update(cookies_dict)

            switch_account_resp, session = dashboard(count_switch, session, table_sb_acc_id)
            print()
            entries_counter, pages_counter, dashboard_session_id, sb_acc_id, sb_user_id = prepare_variables(switch_account_resp, table_sb_acc_id)

            print("entries_counter:", entries_counter)
            print("pages_counter:", pages_counter)
            print(f"SWITCH ACCOUNT: {acc_name}")

            pause_sec = 10
            attempts = 26

            for attempt in range(1, attempts):
                print()
                print("Attempt :", attempt)

                session, entries_resp = entries(sb_acc_id, sb_user_id, session, dashboard_session_id, pages_counter)

                if entries_resp["state"] == "ready":
                    print()
                    print('Entries are ready')

                    data_table, entries_counter, total_count = accumulate_table_data(entries_resp, entries_counter, data_table)

                    entries_counter, pages_counter = check_next_page(pages_counter, entries_counter, total_count)
                    print("entries_counter ", entries_counter)

                    if entries_counter is None and pages_counter is None:
                        acc_count_ready += 1
                        print("work account count", acc_count_ready)
                        break
                    else:
                        continue
                else:

                    print('Entries are', entries_resp["state"])
                    delay = random.uniform(1, 4)
                    #delay = 5
                    time.sleep(delay)

                    session = periods(session, dashboard_session_id)

                if attempt == 25:

                    miss_accounts[acc_name] = details
                    time.sleep(pause_sec)


        print("Work account count", acc_count_ready)

        print(f"Accounts with missing data: {miss_accounts}")

        prepared_data_for_table = prepare_data(data_table)

        fill_google_spreadsheet(prepared_data_for_table)

        transfer_to_total()


    except Exception as e:

        print(f'Error: {e}')
        traceback.print_exc()
        time.sleep(3)

    for en, err in enumerate(errors):
        print("ERROR", en, err["acc_name"], err["error_text"])


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