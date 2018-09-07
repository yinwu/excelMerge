#coding=utf-8

# This functionality is to merge the excel file under input_data folder to one single file output_data.
import sys
import os
import xlrd
import xlwt

INPUT_FOLDER = "./input_data/"
OUTPUT_FOLDER = "./output_data/"

output_length = 0

def is_sheet_header(row_value):
	if row_value[0] == u"项目1":
		print "this is header"
	return True if u"项目1" in row_value else False 

def handle_work_book(sheet_data, output):
	global output_length
	for i in range(sheet_data.nrows):
		row_data = sheet_data.row_values(i)
		# the row 0 is sheet name
		if is_sheet_header(row_data):
			continue

		output_length = output_length + 1
		for j in range(sheet_data.ncols):
			#print "write data on row :" + str(output_length)
			output.write(output_length, j, sheet_data.cell(i, j).value)

def write_header(input, output):
	file_with_path = INPUT_FOLDER + input[0]
	sheet_data = xlrd.open_workbook(file_with_path).sheet_by_index(0)
	for i in range(sheet_data.ncols):
		output.write(0, i, sheet_data.cell(0, i).value)
			
def merge_file(input, output):

	write_header(input, output)
	for file in input:
		file_with_path = INPUT_FOLDER + file
		sheet_data = xlrd.open_workbook(file_with_path).sheet_by_index(0)
		handle_work_book(sheet_data, output)


def read_input_file():
	try:
		file_list = os.listdir(INPUT_FOLDER)
	except e:
		print "Error while reading input file " + e
		file_list = []
	
	return file_list

def open_output_file():
	try:
		output_sheet_data = xlwt.Workbook(encoding='utf-8', style_compression=0)
		return output_sheet_data
	except e:
		print "Failed to read output file"
		return None

def main():
	output_file = open_output_file()

	output_sheet = output_file.add_sheet('result', cell_overwrite_ok=True)
	
	input_file_list = read_input_file()

	if output_file is not None and len(input_file_list) > 0:
		merge_file(input_file_list, output_sheet)
	else:
		print "Something wrong..."

	output_file.save(u"output_result.xls")

if __name__ == '__main__':
	sys.exit(main())