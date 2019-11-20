import csv
import logging
import os

import openpyxl as xl


def convert_csv_to_excel(csv_path):
    (file_path, file_extension) = os.path.splitext(csv_path)

    # check file extension if valid
    wb = xl.Workbook()
    ws = wb.active
    logging.info("converting file to xlsx: '{}'".format(csv_path))
    with open(csv_path, newline='') as csv_file:
        rd = csv.reader(csv_file, delimiter=",", quotechar='"')
        for row in rd:
            ws.append(row)

    output_path = os.path.join(file_path + '.xlsx')
    logging.info("saving to file: '{}'".format(output_path))
    wb.save(output_path)
    return output_path