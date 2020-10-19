import xlsxwriter
import random
import pandas as pd
import csv
import pdb


def highlight_csv_column(csv_file_path):
	xlsx_file_path = csv_file_path.replace(".csv", ".xlsx")
	workbook = xlsxwriter.Workbook(xlsx_file_path)
	worksheet = workbook.add_worksheet()

	headers = []
	column_number = 0
	highlighted_abstract_column_name = "highlighted_abstract2"

	for index, row in enumerate(csv.DictReader(open(csv_file_path))):
		# write headers
		if index == 0:
			headers = list(row.keys()) + [highlighted_abstract_column_name]
			worksheet.write_row(0, 0, headers)

			# get the column number to highlight
			# use it later while writting
			column_number = next((__index for __index, _r in enumerate(headers) if _r == highlighted_abstract_column_name), None)

		worksheet.write_row(index + 1, 0, list(row.values()) + [row["abstract2"]])

		relative_offset = 0
		cell = row["abstract2"]
		formatted_collection = [cell]

		highlight_code = eval(row["code_highlighted_abstract2"])
		for h_point in highlight_code:

			position = h_point['position'] - relative_offset
			cell = formatted_collection.pop()

			formatted_collection += [
				cell[:position],
				workbook.add_format({'color': h_point['color']}),
				cell[position:position + h_point['offset']],
				cell[position + h_point['offset']:]
			]

			relative_offset = h_point['position'] + h_point['offset']

		formatted_collection = list(filter(None, formatted_collection))
		worksheet.write_rich_string("{0}{1}".format(chr(column_number + 65), index + 2), *formatted_collection)

	workbook.close()
	return xlsx_file_path



csv_file_path = "simple.csv"
csv_file_path = "test_highlight.csv"
highlight_csv_column(csv_file_path)
