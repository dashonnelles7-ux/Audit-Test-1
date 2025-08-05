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

    file_path = "sellerboard-accounts.txt"
    import xlrd   # intstall
    from tables import Table
    from dotenv import load_dotenv
    from log_in import start_log_in
    from sellerboard_interaction import (main_frame, update_local_file)

    load_dotenv()
except Exception as e:
    traceback.print_exc()
    time.sleep(5)
    raise e
accounts = {}

tracker = []
pages_counter = 1
count_fill_data = 0


def today_period():

    current_date_utc = datetime.now(timezone.utc)

    start_date = current_date_utc.replace(hour=0, minute=0, second=0, microsecond=0)
    end_date = start_date + timedelta(seconds=86399)

    start_unix_timestamp = int(start_date.timestamp())
    end_unix_timestamp = int(end_date.timestamp())

    print("Current Date (UTC):", current_date_utc.strftime("%Y-%m-%d %H:%M:%S"))
    print("Start of the day (UNIX):", start_unix_timestamp)
    print("End of the day (UNIX):", end_unix_timestamp)
    return start_unix_timestamp, end_unix_timestamp


def delete_rows_today(spreadsheet_id):

    sheet_name = "SellerBoard"
    t = Table(spreadsheet_id)   # for working
    sheet = t.read_range(sheet_name, spreadsheet_id)

    # all_rows = data.get_all_values()
    all_rows = t.try_table_operation(lambda: sheet.get_all_values())

    now = datetime.now()
    current_date = f"{now.month}/{now.day}/{now.year}"

    rows_to_delete = [index + 1 for index, row in enumerate(all_rows) if row[2] == current_date]

    batch_size = 50
    print()
    print("START DELETE TODAY'S ROWS")

    for i in range(0, len(rows_to_delete), batch_size):
        batch = rows_to_delete[i:i + batch_size]

        batch = [row_index - i for row_index in batch]

        t.try_table_operation(lambda: sheet.delete_rows(batch[0], batch[-1]))

    print("Lines with today's date have been removed.")


def main():

    base_dir = "../sb_dashb_to_prod_today"
    cookies_file = "today_cookies.json"

    main_frame(base_dir=base_dir, cookies_file=cookies_file,
               delete_rows_func=delete_rows_today, export_period_func=today_period,
               time_difference_hours=24
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