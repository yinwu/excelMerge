# This functionality is to merge the excel file under input_data folder to one single file output_data.
import sys
import os
import xlrd
import xlwt

INPUT_FOLDER = "./input_data/"
OUTPUT_FOLDER = "./output_data/"

output_length = 0

def is_sheet_header():
	return False

def handle_work_book(sheet_data, output):
	for i in range(sheet_data.nrows):
		# the row 0 is sheet name
		if i == 0:
			continue
		global output_length
		output_line = output_length
		for j in range(sheet_data.ncols):
			print "write data on row :" + str(output_line)
			output.write(output_line, j, sheet_data.cell(i, j).value)
		output_length = output_line + 1	


def merge_file(input, output):
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