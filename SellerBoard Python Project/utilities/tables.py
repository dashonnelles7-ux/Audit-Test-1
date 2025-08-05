import re
import time
import random
import json
from itertools import cycle

import gspread
from gspread.worksheet import Worksheet
from gspread.exceptions import WorksheetNotFound
from requests.exceptions import ConnectionError
import datetime
from google.auth.exceptions import TransportError
import os
from dotenv import load_dotenv

load_dotenv()
import sys
sys.path.append(
    "C:\\Program Files\\ZennoLab\\RU\\ZennoPoster Pro V7\\7.7.0.0\\Progs\\Projects\\AMZ Professional\\SellerBoard Python Project\\utilities")
current_dir = os.path.dirname(__file__)


class Table(object):
    def __init__(self, spreadsheet_id: str = ""):
        google_service_accounts_path = os.path.join(current_dir, 'google_service_accounts.json')
        with open(google_service_accounts_path) as gsa_file:
            self.credentials_list = json.load(gsa_file)["files"]
        random.shuffle(self.credentials_list)
        self.credentials = self.credentials_list[0]
        credentials_file_path = os.path.join(current_dir, self.credentials)
        self.client = gspread.service_account(filename=credentials_file_path)
        self.max_attempts = int(os.getenv("MAX_TABLE_OPERATION_ATTEMPTS", 3))
        if spreadsheet_id != "":
            self.spreadsheet_id = spreadsheet_id
            self.spreadsheet = self.try_table_operation(lambda: self.client.open_by_key(self.spreadsheet_id))

    def read_range(self, sheet_name, source_key: str = os.getenv("TEMPLATE_SPREADSHEET_ID")) -> Worksheet:
        spreadsheet = self.spreadsheet

        try:
            sheet = self.try_table_operation(lambda: spreadsheet.worksheet(sheet_name))
        except WorksheetNotFound:
            template_spreadsheet = self.try_table_operation(lambda: self.client.open_by_key(source_key))
            old_sheet_dict = self.try_table_operation(lambda: template_spreadsheet.worksheet(sheet_name))
            new_sheet_dict = self.try_table_operation(lambda: old_sheet_dict.copy_to(self.spreadsheet_id))
            sheet = self.try_table_operation(lambda: spreadsheet.worksheet(new_sheet_dict['title']))
            self.try_table_operation(lambda: sheet.update_title(sheet_name))

        return sheet

    def update_range(self, sheet: Worksheet,  values: list, start_row: int = 2, start_col: str = "A"):
        new_err = None

        range_str = start_col + str(start_row) + ":" + str(start_row + len(values) - 1)
        try:

            result_updating = self.try_table_operation(lambda: sheet.batch_update([
                {
                    'range': range_str, #start_col + str(start_row) + ":" + chr(ord(start_col) + len(values[0]) - 1).upper(),
                    'values': values,
                }
            ], value_input_option='USER_ENTERED'))
            print(result_updating)

        except Exception as e:
            if "Invalid value at" in str(e):
                try:
                    self.try_table_operation(lambda: sheet.update(range_str, values))
                except Exception as e1:
                    new_err = e1
            if "Please try again in 30 seconds" in str(e) or \
                    (new_err is not None and "Please try again in 30 seconds" in str(new_err)):
                try:
                    time.sleep(30)
                    self.try_table_operation(lambda: sheet.update(range_str, values))
                    return
                except Exception as e2:
                    new_err = e2
                    print("update_range Помилка при оновленні даних після 30 секунд:", str(e))
            if "cells in the workbook above the limit of 10000000 cells" in str(e) or \
                    (new_err is not None and "cells in the workbook above the limit of 10000000 cells" in str(new_err)):
                try:
                    new_len = round(len(values) / 2)
                    self.update_range(sheet, values[:new_len])
                    self.update_range(sheet, values[new_len:], start_row=new_len + 2)

                    return
                except Exception as e3:
                    new_err = e3
                    print("update_range Помилка при оновленні даних після 30 секунд:", str(e))
            if "exceeds grid limits" in str(e) or \
                    (new_err is not None and "exceeds grid limits" in str(new_err)):
                try:
                    self.append_rows(sheet, values)

                    return
                except Exception as e4:

                    print("update_range Помилка при оновленні даних після 30 секунд:", str(e))
            print("update_range Помилка при оновленні даних:", str(e))

    def update_ranges(self, values: list, start_row: int = 2, start_col: str = "A") -> dict:
        range_str = start_col + str(start_row) + ":" + str(start_row + len(values) - 1)
        return {
                'range': range_str,
                'values': values,
            }

    def append_rows(self, sheet: Worksheet, values: list):
        try:
            self.try_table_operation(lambda: sheet.append_rows(values))
        except Exception as e:
            if "cells in the workbook above the limit of 10000000 cells" in str(e):
                new_len = round(len(values) / 2)
                self.append_rows(sheet, values[:new_len])
                self.append_rows(sheet, values[new_len:])
            else:
                raise Exception("Error in appending rows")

    def group_rows(self, sheet: Worksheet, start_row: int = -1, last_row: int = -1, update_type: str = "BY_FIRST_COLUMN",
                   start_col: int = 1, dimension_columns: list = [], undimensioned_columns: list = []):

        values = self.try_table_operation(lambda: sheet.get_all_values())
        row_count = len(values)
        grouped_rows = None
        applied_dimensions = []
        sheet_id = self.try_table_operation(lambda: sheet.id)
        req = {"requests": []}

        if update_type == "BY_FIRST_COLUMN":
            for r, row in enumerate(values):
                is_row_only_name = row[0] != '' and len("".join(row[start_col:-1])) == 0
                is_sub_row = row[0] == '' and row[start_col] != ''
                is_last_row = r == row_count-1

                if start_row == -1 and is_row_only_name:
                    start_row = r
                    last_row = r + 1

                elif start_row > 0 and (is_row_only_name or is_last_row):
                    if is_last_row and not is_row_only_name:
                        last_row += 1
                    if start_row != last_row - 1:
                        applied_dimensions.append([start_row + 1, last_row])

                        if grouped_rows is None:
                            grouped_rows = self.try_table_operation(lambda: sheet.list_dimension_group_rows())
                        exist_dimension = False

                        exist_dimension = any('range' in dim and 'endIndex' in dim['range'] and 'startIndex' in dim['range'] and
                                              dim['range']['endIndex'] == last_row and dim['range']['startIndex'] == start_row + 1
                                              for dim in grouped_rows)
                        if not exist_dimension:
                            req['requests'].append({
                                "addDimensionGroup": {
                                    "range": {
                                        "sheetId": sheet_id,
                                        "startIndex": start_row + 1,
                                        "endIndex": last_row,
                                        "dimension": "ROWS"
                                    }
                                }
                            })
                    start_row = r
                    last_row = r+1

                elif start_row > 0 and is_sub_row:
                    last_row += 1
            try:
                res = self.try_table_operation(lambda: self.spreadsheet.batch_update(req))
            except Exception as e:
                print(str(e))
                raise e
            if grouped_rows is None:
                grouped_rows = self.try_table_operation(lambda: sheet.list_dimension_group_rows())

            for table_dim in grouped_rows:
                start_t_dim = table_dim['range']['startIndex']
                end_t_dim = table_dim['range']['endIndex']
                is_in_applied = any(end_t_dim == a_dim[1] and start_t_dim == a_dim[0]
                                    for a_dim in applied_dimensions)
                if not is_in_applied:
                    self.try_table_operation(lambda: sheet.delete_dimension_group_rows(start_t_dim, end_t_dim))

            return

        if update_type == "BY_NUMEROUS_COLUMNS":

            # delete from
            grouped_rows = self.try_table_operation(lambda: sheet.list_dimension_group_rows())

            for table_dim in grouped_rows:
                start_t_dim = table_dim['range']['startIndex']
                end_t_dim = table_dim['range']['endIndex']
                req['requests'].append({
                    "deleteDimensionGroup": {
                        "range": {
                            "sheetId": sheet_id,
                            "startIndex": start_t_dim,
                            "endIndex": end_t_dim,
                            "dimension": "ROWS"
                        }
                    }
                })
            try:
                if req["requests"]:
                    res = self.try_table_operation(lambda: self.spreadsheet.batch_update(req))
            except Exception as e:
                raise e

            req = {"requests": []}

            # #delete to

            start_row_default = start_row
            last_row_default = -1
            for d_c_index, d_c in enumerate(dimension_columns[:-1]):
                prev_dim_data = None
                start_row = start_row_default
                last_row = last_row_default
                for r, row in enumerate(values):
                    if r < start_row:
                        continue
                    is_dim_data = row[d_c] != '' and (d_c not in undimensioned_columns or\
                                                    row[d_c] != prev_dim_data) # and row[dimension_columns[-1]] == ''
                    # if all cells after last undimensioned == ""
                    # try:
                    is_sub_row = ((undimensioned_columns == []) and all(row[sub_d_c] == "" for sub_d_c in dimension_columns[:d_c_index+1]))\
                                 or (undimensioned_columns != [] and d_c_index > undimensioned_columns[-1] and all(row[sub_d_c] == "" for sub_d_c in dimension_columns[len(undimensioned_columns):d_c_index+1]))\
                                 or (undimensioned_columns != [] and d_c in undimensioned_columns and row[d_c] == prev_dim_data)
                    # except Exception as e:
                    #     raise e
                    is_last_row = r == row_count - 1

                    prev_dim_data = row[d_c]

                    if start_row == -1 and is_dim_data:
                        start_row = r
                        # last_row = r + 1
                    if last_row == -1:
                        last_row = start_row + 1

                    elif start_row > 0 and (is_dim_data or is_last_row):
                        if is_last_row and not is_dim_data:
                            last_row += 1
                        if start_row != last_row - 1:
                            applied_dimensions.append([start_row + 1, last_row])

                            if grouped_rows is None:
                                grouped_rows = self.try_table_operation(lambda: sheet.list_dimension_group_rows())

                            try:
                                exist_dimension = any('range' in dim and 'endIndex' in dim['range'] and 'startIndex' in dim['range'] and
                                                      dim['range']['endIndex'] == last_row and dim['range']['startIndex'] == start_row + 1
                                                      for dim in grouped_rows)
                            except Exception as e1231:
                                raise e1231

                            if not exist_dimension:
                                req['requests'].append({
                                    "addDimensionGroup": {
                                        "range": {
                                            "sheetId": sheet_id,
                                            "startIndex": start_row + 1,
                                            "endIndex": last_row,
                                            "dimension": "ROWS"
                                        }
                                    }
                                })
                                # response = self.try_table_operation(lambda: sheet.add_dimension_group_rows(start_row + 1, last_row))
                        start_row = r
                        last_row = r+1

                    elif start_row > 0 and is_sub_row:
                        last_row += 1

            print('req ', req)
            if req["requests"]:
                res = self.try_table_operation(lambda: self.spreadsheet.batch_update(req))
            else:
                print("No requests for batch_update")

            # res = self.try_table_operation(lambda: self.spreadsheet.batch_update(req))
            req = {"requests": []}

            print(res)
            if grouped_rows is not None:
                for table_dim in grouped_rows:
                    start_t_dim = table_dim['range']['startIndex']
                    end_t_dim = table_dim['range']['endIndex']
                    is_in_applied = any(end_t_dim == a_dim[1] and start_t_dim == a_dim[0]
                                        for a_dim in applied_dimensions)
                    if not is_in_applied:
                        req['requests'].append({
                            "deleteDimensionGroup": {
                                "range": {
                                    "sheetId": sheet_id,
                                    "startIndex": start_row + 1,
                                    "endIndex": last_row,
                                    "dimension": "ROWS"
                                }
                            }
                        })
                        # self.try_table_operation(lambda: sheet.delete_dimension_group_rows(start_t_dim, end_t_dim))
            if req["requests"]:
                res = self.try_table_operation(lambda: self.spreadsheet.batch_update(req))

            return

        if start_row == -1:
            start_row = 0
        if last_row == -1:
            last_row = len(values) - 1

        grouped_rows = self.try_table_operation(lambda: sheet.list_dimension_group_rows())

        exist_dimension = any('range' in dim and 'endIndex' in dim['range'] and 'startIndex' in dim['range'] and
                              dim['range']['endIndex'] == last_row and dim['range']['startIndex'] == start_row + 2
                              for dim in grouped_rows)

        if not exist_dimension:
            sheet_id = self.try_table_operation(lambda: sheet.id)
            grouped_rows = sheet.list_dimension_group_rows()

            self.try_table_operation(lambda: sheet.batch_update([
                {
                    "addDimensionGroup": {
                        "range": {
                            "dimension": "ROWS",
                            "sheetId": sheet_id,
                            "startIndex": start_row + 2,
                            "endIndex": last_row
                        }
                    }
                }
            ]))

    def check_is_data_updated(self, sheet_name, penultimate_row: list, last_row: list):
        print("check_is_data_updated start", self.spreadsheet_id, sheet_name)

        sheet = self.read_range(sheet_name)
        while True:
            try:
                t_last_row = self.try_table_operation(lambda: sheet.get_all_values())[-1]
                if t_last_row == penultimate_row:
                    self.try_table_operation(lambda: sheet.append_row(last_row))
                    raise Exception("Did not add any new row to sheet " + sheet_name)
                print("check_is_data_updated finish", self.spreadsheet_id, sheet_name)
                break
            except Exception as e:
                if "429" in str(e):
                    print("error code: 429")
                    time.sleep(random.randint(10, 20))
                    continue
                raise Exception(self.spreadsheet_id, "check_is_data_updated Помилка при отриманні даних:", e)

    def check_periodically(self, sheet_name, period, start_col: int = 0, last_col: int = 0, row_check: int = -1,
                           days_delta: int = 0):
        """
                Checks if the data was filled in during this time period.

                Args:
                    sheet_name (str): Name of the sheet in the table.
                    period (Period): Regularity of checking the data.
                    start_col (int): Starting column for checking the data.
                    last_col (int): Last column for checking the data.
                    row_check (int): Index of the row for checking the data.
                    days_delta (int): Days offset from the current date.

                Returns:
                    bool, str: A boolean indicating the success of the check, followed by a message
                               about the result of the check
        """
        print("check_periodically", self.spreadsheet_id, sheet_name, period, start_col, last_col)
        now = datetime.datetime.now() - datetime.timedelta(days=days_delta)
        time_pattern = r"\d+:\d+:\d+"

        try:
            last_row_values = self.try_table_operation(lambda: self.read_range(sheet_name).get_all_values())
        except Exception as e:
            return False, f"Failed to retrieve data: {str(e)}"

        if row_check >= len(last_row_values):
            return False, f"{sheet_name} Current timedate is not included"

        last_row = last_row_values[row_check][start_col:last_col or None]
        is_date_included = False

        for i, cell in enumerate(last_row):
            if re.match(time_pattern, cell):
                prev_cell = last_row[i - 1] if i > 0 else now.strftime("%d.%m.%Y")
                cell_date = datetime.datetime.strptime(f"{prev_cell} {cell}", "%d.%m.%Y %H:%M:%S")
                minutes_diff = (now - cell_date).total_seconds() / 60
                periods = {
                    Period.HALF_DAILY: 780,
                    Period.TWO_HOURS: 121,
                    Period.HOURLY: 61,
                    Period.HALF_HOURLY: 31,
                    Period.DAILY: 1440
                }
                if minutes_diff < periods.get(period, 1440) or (
                        period == Period.DAILY and now.strftime("%d.%m.%Y") == cell_date.strftime("%d.%m.%Y")):
                    is_date_included = True
                    break

        #  If the data was entered during this period, then True; if not, then False.
        if is_date_included:
            return True, f"{sheet_name} Data is added successfully"
        else:
            return False, f"{sheet_name} Current timedate is not included"

    def clear_check_last_rows(self, sheet_name, source_key: str = os.getenv("TEMPLATE_SPREADSHEET_ID")):
        print("clear_check_last_rows start", self.spreadsheet_id, sheet_name)
        before_last_row, last_row = [], []

        sheet = self.read_range(sheet_name, source_key)

        sheet_values = self.try_table_operation(lambda: sheet.get_all_values())
        # while True:
        #     try:
        #         sheet_values = sheet.get_all_values()
        #         break
        #     except Exception as e:
        #         if "429" in str(e):
        #             print("error code: 429")
        #             time.sleep(10)
        #             continue
        #         raise Exception(self.spreadsheet_id, "clear_check_last_rows", e)

        row_count = len(sheet_values)
        if row_count > 1:
            before_last_row = sheet_values[-2]
            last_row = sheet_values[-1]

            # Delete last row
            self.try_table_operation(lambda: sheet.delete_rows(row_count))
            self.try_table_operation(lambda: sheet.append_rows([[]]))

        elif row_count == 1:
            before_last_row = sheet_values[0]
        elif row_count == 0:
            print("There is no first row with names of columns. "
                  "\nAdd some information to your table")

        print("clear_check_last_rows finish", self.spreadsheet_id, sheet_name)

        return before_last_row, last_row

    def empty_last_rows(self, sheet: Worksheet, rows: list[list]):

        values = self.try_table_operation(lambda: sheet.get_all_values())
        empty_row_count = len(values) + 2 - len(rows)
        if empty_row_count > 0:
            empty_rows = [[""] * len(values[0]) for _ in range(empty_row_count)]
        else:
            empty_rows = []
        return empty_rows

    def try_table_operation(self, operation, sleep_time: int = random.randint(3, 6)):
        """
        Attempts to execute an operation on the table, handling possible exceptions.
        Automatically retries the operation upon encountering certain types of errors.

        Args:
            operation (function): The function representing the operation to be executed.
            sleep_time (int): The delay time between attempts to execute the operation in case of errors.

        Returns:
            The result of the operation (function) or None in case of unforeseen errors.

        Description:
            The method automatically retries the operation when encountering specific types of errors,
            such as exceeding request limits or temporary service unavailability (e.g., 429, 503, 500, 409, 404).
            For errors that do not envisage repetition (e.g., network error or access error), an exception is raised.
            This could lead to returning None if the exception is not handled.
        """

        k = 0  # Attempt counter
        while True:
            if k == self.max_attempts:
                raise Exception('Max attempt try!')

            try:
                return operation()  # Attempt to execute the operation

            except TransportError as e0:
                print(f"error code: {str(e0)},{str(operation.__code__)}")
                k += 1
                time.sleep(sleep_time + 1)
            except Exception as e:
                if (any(code in str(e) for code in ["429", "503", "500", "409", "404"])
                        or "Please try again in 30 seconds" in str(e)):
                    operation_error = ''
                    try:
                        operation_error = e.args[0]['details'][0]['metadata']['quota_limit']
                    except:
                        pass
                    error_code = 0
                    if "429" in str(e):
                        error_code = 429
                    elif "503" in str(e):
                        error_code = 503
                    elif "500" in str(e):
                        error_code = 500
                    elif "409" in str(e):
                        error_code = 409
                    elif "404" in str(e):
                        error_code = 404
                    print(f"error code: {error_code}, {operation_error},{str(operation.__code__)}")
                    time.sleep(sleep_time + 1)  # Delay before the next attempt
                    k += 1  # Increase attempt counter
                else:
                    # Re-throw the exception to the calling code if it does not envisage repetition
                    raise e


class Period:
    HALF_HOURLY = "half_hourly"
    HOURLY = "hourly"
    HALF_DAILY = "half_daily"
    DAILY = "daily"
    TWO_HOURS = "every two hours"
    WEEKLY = "weekly"


def get_col_letter_by_num(column_int, start_index: int = 0):
    # start_index = 0  # it can start either at 0 or at 1
    col_letter = ''

    while column_int > 25 + start_index:
        col_letter += chr(65 + int((column_int - start_index) / 26) - 1)
        column_int = column_int - (int((column_int - start_index) / 26)) * 26
    col_letter += chr(65 - start_index + (int(column_int)))
    return col_letter
