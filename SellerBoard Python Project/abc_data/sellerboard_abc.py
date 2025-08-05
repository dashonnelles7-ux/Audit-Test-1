
import os
import time
import multiprocessing
from sys import platform
import sys
sys.path.append("C:\\Program Files\\ZennoLab\\RU\\ZennoPoster Pro V7\\7.7.0.0\\Progs\\Projects\\AMZ Professional\\SellerBoard Python Project\\utilities")
sys.path.append(
    "D:\\PyCharm Projects\\SellerBoard Python Project\\utilities")
from tables import Table
from sellerboard_interaction import update_local_file
from datetime import datetime as dt
import traceback

from dotenv import load_dotenv
load_dotenv()

errors = []


def abc():
    # session, headers = prepare_session(profile)

    # path = 'en/dashboard/'
    # headers.update({
    #     'Authority': "app.sellerboard.com",
    #     "Path": '/' + path
    # })
    #
    # resp_get = session.get(url=sb_url + path, headers=headers)
    # resp_text = resp_get.text
    # dash_board_seesion_id = re.search(r'(?<=data-dashboardSessionId=").*?(?=")', resp_text).group()
    # sb_ac_id = re.search(r'(?<=data-account=").*?(?=")', resp_text).group()
    # sb_user_id = re.search(r'(?<=data-userId=").*?(?=")', resp_text).group()
    #
    # path = "en/dashboard/entries"

    now = dt.now()
    today = now.strftime('%d.%m.%Y')
    # this_monday = now - datetime.timedelta(days=now.weekday())
    # this_sunday = now + datetime.timedelta(days=6 - now.weekday())
    #
    # this_monday_int = int(dt(this_monday.year, this_monday.month, this_monday.day, 0, 0, 0,
    #                                tzinfo=datetime.timezone.utc).timestamp())
    # this_sunday_int = int(dt(this_sunday.year, this_sunday.month, this_sunday.day, 23, 59, 59,
    #                                tzinfo=datetime.timezone.utc).timestamp())
    # today_last_min_int = int(dt(now.year, now.month, now.day, 23, 59, 59,
    #                                   tzinfo=datetime.timezone.utc).timestamp())
    # year_ago_start = int(dt(now.year - 1, now.month, now.day, 0, 0, 0,
    #                               tzinfo=datetime.timezone.utc).timestamp())
    #
    # headers.update({
    #     "Path": '/' + path,
    #     "Method": "POST",
    #     "Accept-Language": "en-US,en;q=0.9,es-US;q=0.8,es;q=0.7",
    #     "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
    #     "Origin": "https://app.sellerboard.com",
    #     "Referer": "https://app.sellerboard.com/en/dashboard/?viewType=table&periodsTableRange=Last_3_months_by_week&tablePeriod%5Bstart%5D=" + str(this_monday_int) + " &tablePeriod%5Bend%5D=" + str(this_sunday_int) + "&tablePeriod%5Bforecast%5D=false",
    #     "Sec-Ch-Ua": "\"Chromium\";v=\"118\", \"Google Chrome\";v=\"118\", \"Not=A?Brand\";v=\"99\"",
    #     "Sec-Ch-Ua-Platform": "\"Linux\"",
    #     "Sec-Fetch-Dest": "empty",
    #     "Sec-Fetch-Mode": "cors",
    #     "Sec-Fetch-Site": "same-origin",
    #     "Sellerboard-Account-Id": sb_ac_id,
    #     "Sellerboard-User-Id": sb_user_id,
    #     "X-Requested-With": "XMLHttpRequest",
    # })
    #
    #
    # n = 1
    # rows = []
    #
    # data = {
    #     "_": 0,  # GMT: Wednesday, 20 December 2023 р., 11:53:43.822
    #     "dashboardSessionId": dash_board_seesion_id,
    #     "viewType": "table",
    #     "entryType": "product",
    #     "periodStart": this_monday_int,  # GMT: Monday, 18 December 2023 р., 00:00:00
    #     "periodEnd": this_sunday_int,  # GMT: Sunday, 24 December 2023 р., 23:59:59
    #     "periodicity": "period",
    #     "sortField": "units",
    #     "sortDirection": "desc",
    #     "page": n,  # n
    #     "groupByAsin": "",
    #     "groupBy": "",
    #     "rangeStart": year_ago_start,  # GMT: Tuesday, 20 December 2022 р., 00:00:00
    #     "rangeEnd": today_last_min_int,  # GMT: Wednesday, 20 December 2023 р., 23:59:59
    #     "rangePeriodicity": "month",
    #     "trendsParameter": "sales",
    #     "collapseEntriesByAsin": "false"
    # }
    #     "compare": "false",
    #
    # while True:
    #     now_int = int(dt.now().timestamp())
    #
    #     data.update({
    #         "_": now_int,
    #         "page": n,
    #     })
    #
    #
    #     resp = session.post(sb_url + path, headers=headers, json=data, timeout=300)
    #     s = resp.content
    #     resp_json = resp.json()
    #     new_rows = resp_json['entries']
    #
    #     data_str = ''
    #     for key in list(data.keys()):
    #         data_str += key + "=" + str(data[key])
    #         if key is not list(data.keys())[-1]:
    #             data_str += "&"
    #     resp = session.post(sb_url + path, headers=headers, data=data_str, timeout=300)
    #     resp_json = resp.json()
    #
    #     new_rows = resp['entries']
    #     rows.extend(new_rows)
    #
    #     if len(new_rows) == 0:
    #         break
    abc_data_spreadsheet_id = os.getenv("ABC_DATA_SPREADSHEET_ID")
    t = Table(abc_data_spreadsheet_id)
    sheet_name = os.getenv("ABC_SHEET_NAME")

    sheet = t.read_range(sheet_name)
    data_rows = t.try_table_operation(lambda: sheet.get_all_values())
    data_rows.pop(0)
    data_rows.pop()
    total_sheet_name = os.getenv("ABC_TOTAL_SHEET_NAME")

    total_sheet = t.read_range(total_sheet_name)

    values = t.try_table_operation(lambda: total_sheet.get_all_values())
    values.pop(0)

    new_data = create_dimensioned_table(values, data_rows, 2, 2, today)
    t.update_range(total_sheet, new_data, 2)
    t.group_rows(total_sheet, start_col=2)

    update_local_file(abc_data_spreadsheet_id)


def create_dimensioned_table(old_values: list, new_values: list, start_col_to_dimension: int = 1, checker_col: int = 1,
                             checker_data: str = "") -> list:
    start_list = [""] * start_col_to_dimension

    for o, old_row in enumerate(old_values):
        if old_row[start_col_to_dimension] != "":
            continue
        if old_row[0] == "" and old_row[start_col_to_dimension] == "":
            break
        if old_row[0] != "":
            group_name = old_row[0]
            for new_row in new_values:
                if group_name == new_row[0]:
                    is_that_row_already_in_table = False
                    for sub_old_index in range(o+1, len(old_values)):
                        old_row_check = old_values[sub_old_index]
                        if old_row_check[checker_col] == checker_data:
                            is_that_row_already_in_table = True
                            break
                        if old_row_check[0] != "" or \
                                (old_row_check[0] == "" and old_row_check[start_col_to_dimension] == ""):
                            break
                    if is_that_row_already_in_table:
                        new_values.remove(new_row)
                        break
                    old_values.insert(o + 1, start_list + new_row[start_col_to_dimension:])
                    new_values.remove(new_row)
                    break

    edited_values = old_values.copy()
    for new_row in new_values:
        edited_values.append(new_row[:start_col_to_dimension])
        edited_values.append(start_list + new_row[start_col_to_dimension:])
    return edited_values


def main():
    try:
        abc()
    except Exception:
        error_text = traceback.format_exc()
        errors.append({
            "error_text": error_text
        })

    for en, err in enumerate(errors):
        print("ERROR", en, err["error_text"])

    if platform == "win32":
        # os.system('taskkill /im chrome.exe /f')
        os.system('taskkill /im chromedriver.exe /f')
    else:
        os.system('killall -9 chrome')
        os.system('killall -9 chromedriver')


if __name__ == '__main__':
    big_error = Exception
    try:
        multiprocessing.freeze_support()
        main()

    except Exception as e:
        big_error = e
        print("big_error", str(big_error))
        time.sleep(5)
    if big_error is not Exception:
        raise big_error
    time.sleep(10)