try:
    import sys
    import os
    import traceback
    sys.path.append("C:\\Program Files\\ZennoLab\\RU\\ZennoPoster Pro V7\\7.7.0.0\\Progs\\Projects\\AMZ Professional\\SellerBoard Python Project\\utilities")
    sys.path.append(
        "D:\\PyCharm Projects\\SellerBoard Python Project\\utilities")
    import time
    import requests
    import re
    import json
    from datetime import datetime, timedelta, timezone

    import xlrd   # intstall
    from tables import Table
    from log_in import start_log_in
    from sellerboard_interaction import main_frame

    from dotenv import load_dotenv
    load_dotenv()
except Exception as e:
    traceback.print_exc()
    time.sleep(5)
    raise e


accounts = {}

tracker = []
pages_counter = 1
count_fill_data = 0


n_days = 60

def last_n_days_period():
    global n_days
    current_date = datetime.now()

    print(n_days)
    start_date = current_date - timedelta(days=n_days)
    start_date = datetime(start_date.year, start_date.month, start_date.day, 0, 0, 0)

    end_date = current_date - timedelta(days=1)
    end_date = datetime(end_date.year, end_date.month, end_date.day, 23, 59, 59)

    start_unix_timestamp = int(start_date.replace(tzinfo=timezone.utc).timestamp())
    end_unix_timestamp = int(end_date.replace(tzinfo=timezone.utc).timestamp())

    print("Period Start:", start_unix_timestamp)
    print("Period End:", end_unix_timestamp)

    period_start = str(start_unix_timestamp)
    period_end = str(end_unix_timestamp)

    return period_start, period_end


def delete_rows_n_days_ago(spreadsheet_id):
    global n_days

    sheet_name = "SellerBoard"  # for working

    #spreadsheet_id = '14pgfr20kfwhCY_h2Xocn1uVks8LTmE0YLbsibQ8INzI'  # for testing
    t = Table(spreadsheet_id)   # for working

    sheet = t.read_range(sheet_name, spreadsheet_id)

    # all_rows = data.get_all_values()
    all_rows = t.try_table_operation(lambda: sheet.get_all_values())

    now = datetime.now()
    today = f"{now.month}/{now.day}/{now.year}"
    dates_to_delete = {today}

    for d in range(1, n_days + 1):
        n_days_ago = (now - timedelta(days=d)).strftime("%m/%d/%Y").lstrip('0').replace('/0', '/')
        dates_to_delete.add(n_days_ago)

    rows_to_delete = [
        index + 1 for index, row in enumerate(all_rows) if row[2] in dates_to_delete
    ]

    #print("Rows for delete:", rows_to_delete)

    batch_size = 500
    print()
    print("START DELETE TODAY'S ROWS")

    for i in range(0, len(rows_to_delete), batch_size):
        batch = rows_to_delete[i:i + batch_size]

        batch = [row_index - i for row_index in batch]

        t.try_table_operation(lambda: sheet.delete_rows(batch[0], batch[-1]))

    print("Lines with today's date have been removed.")


def main():

    base_dir = "../sb_dashb_to_prod_n_days"
    cookies_file = "n_days_cookies.json"

    main_frame(base_dir=base_dir, cookies_file=cookies_file,
               delete_rows_func=delete_rows_n_days_ago, export_period_func=last_n_days_period,
    )


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