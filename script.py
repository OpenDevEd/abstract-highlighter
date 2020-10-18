import xlsxwriter
import random
import pandas as pd
import csv
import pdb


class HighPointAxis:

	def __init__(self, position=None, offset=None, color=None):
		self.color = color		
		self.offset = offset
		self.position = position


class HighPoint:

	def __init__(self, row=None, column=None, column_name=None):
		self.row = row
		self.column = column
		self.column_name = column_name

		self.axis_collection = []

	def add_axis(self, position=None, offset=None, color=None):
		axis = HighPointAxis(position, offset, color)
		self.axis_collection.append(axis)



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
for index in range(0, len(data)):

	colors = ["red", "yellow", "green"]

	column_name = "E Mail"
	column_number = next((__index for __index, _r in enumerate(headers) if _r == column_name), None)

	h_point = HighPoint(index, column_number, column_name)
	h_point.add_axis(0, 2, "green")
	h_point.add_axis(3, 2, "yellow")
	h_point.add_axis(6, 2, "red")

	highlight_points.append(h_point)

# real working
for h_point in highlight_points:
	print("Row: {0}".format(h_point.row))

	relative_offset = 0
	cell = data[h_point.row][h_point.column_name]
	formatted_collection = [cell]

	for axis in h_point.axis_collection:
		position = axis.position - relative_offset
		cell = formatted_collection.pop()

		formatted_collection += [
			cell[:position],
			workbook.add_format({'color': axis.color}),
			cell[position:position + axis.offset],
			cell[position + axis.offset:]
		]

		relative_offset = axis.position + axis.offset

	formatted_collection = list(filter(None, formatted_collection))
	worksheet.write_rich_string("{0}{1}".format(chr(h_point.column + 65), h_point.row + 2), *formatted_collection)

workbook.close()

