import xlsxwriter
import random
import pandas as pd
import csv
import pdb



class HighPoint:

	column_name = None
	row = None
	position = None
	offset = None
	color = None

	def __init__(self, column_name=None, row=None, position=None, offset=None, color=None):
		self.column_name = column_name
		self.row = row
		self.position = position
		self.offset = offset
		self.color = color


csv_file_path = "example.csv"

xlsx_file_path = csv_file_path.replace(".csv", ".xlsx")
workbook = xlsxwriter.Workbook(xlsx_file_path)
worksheet = workbook.add_worksheet()

data = []
headers = []

for index, row in enumerate(csv.DictReader(open(csv_file_path))):
	# write headers
	if index == 0:
		headers = list(row.keys())
		worksheet.write_row(0, 0, headers)

	worksheet.write_row(index + 1, 0, row.values())
	data.append(row)


# creating dummy data
highlight_points = []
for index in range(1, 999):

	colors = ["red", "yellow", "green"]

	column_name = "E Mail"
	column_number = next((__index for __index, _r in enumerate(headers) if _r == column_name), None)

	# h_point = HighPoint(column_number, index, random.choice(range(1, 10)), random.choice(range(1, 10)), random.choice(colors))
	h_point = HighPoint(column_number, index, 1, 2, "green")
	highlight_points.append(h_point)

	_color = workbook.add_format({'color': h_point.color})

	try:
		worksheet.write_rich_string(
			"{0}{1}".format(chr(column_number + 65), index),
			data[index][column_name][0:h_point.position],
			_color, data[index][column_name][h_point.position:h_point.position + h_point.offset],
			data[index][column_name][h_point.offset:]
		)
	except Exception as e:
		pass

workbook.close()

