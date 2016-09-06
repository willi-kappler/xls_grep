#!/usr/bin/env python3

# xls_grep, written by Willi Kappler
# Licensed under the MIT License
# https://github.com/willi-kappler/xls_grep

import argparse
import os
import logging
import datetime
import openpyxl
import xlrd

VERSION = "2016.09.06"

def process_folder(folder, expression):
    all_items = os.walk(folder)

    for root, sub_folders, files in all_items:
        for file_name in files:
            if file_name.endswith(".xls"):
                full_path = os.path.join(root, file_name)
                try:
                    wb = xlrd.open_workbook(full_path)
                    for sheet in wb.sheets():
                        for row in range(0, sheet.nrows):
                            for col in range(0, sheet.ncols):
                                data = sheet.cell(row, col)
                                # TODO: use regular expression
                                if (data.ctype == 1) and (expression in data.value):
                                    logging.info("examine file '{}'".format(full_path))
                                    logging.info("Found match: '{}'".format(data))
                except Exception as e:
                    logging.info("Error while reading the file: {}".format(e))

            elif file_name.endswith(".xlsx"):
                full_path = os.path.join(root, file_name)
                try:
                    wb = openpyxl.load_workbook(full_path)
                    for sheet in wb.worksheets:
                        for row in sheet.rows:
                            for cell in row:
                                # TODO: use regular expression
                                if cell.data_type == "s" and expression in cell.value:
                                    logging.info("examine file '{}'".format(full_path))
                                    logging.info("Found match: '{}'".format(data))

                except Exception as e:
                    logging.error("Error while reading the file: {}".format(e))


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Grep for excel files")
    parser.add_argument("--folder", default=".", help="The search folder, default '.'")
    parser.add_argument("--expression", required=True, help="The expression to search for, mandatory")

    args = parser.parse_args()

    esd_now = datetime.datetime.now().strftime("%Y_%m_%d___%H_%M_%S")
    log_file_name = "grep_result_{0:s}.log".format(esd_now)
    logging.basicConfig(format="%(asctime)s %(message)s", filename=log_file_name, level=logging.INFO)
    logging.info("XLS Grep - version: {0:s}".format(VERSION))

    process_folder(args.folder, args.expression)
