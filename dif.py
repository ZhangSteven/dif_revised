# coding=utf-8
# 
# Read holdings from China Life trustee's DIF excel file. It is actually
# a rewritten of the old DIF package, with a more clear structure. Structure
# and code are similar to clamc_trustee package.
# 

from xlrd import open_workbook
from functools import reduce
from itertools import chain
from datetime import datetime
import csv, re

import logging
logger = logging.getLogger(__name__)



def worksheetToLines(ws):
	"""
	wb: a worksheet object (from xlrd.open_workbook)

	output: [list] a list of lines in the worksheet. A line is a list of
		content in the columns.
	"""
	lines = []
	row = 0
	while row < ws.nrows:
		thisRow = []
		column = 0
		while column < ws.ncols:
			cellValue = ws.cell_value(row, column)
			if isinstance(cellValue, str):
				cellValue = cellValue.strip()
			thisRow.append(cellValue)
			column = column + 1

		lines.append(thisRow)
		row = row + 1

	return lines



def linesToSections(lines):
	"""
	lines: [iterable] a list of lines from a 

	output: [list] a list of sections, each section being a list 
		of lines in that section.
	"""
	def notEmptyLine(line):
		for i in range(len(line) if len(line) < 20 else 20):
			if not isinstance(line[i], str) or line[i] != '':
				return True

		return False

	def startOfSection(line):
		"""
		Tell whether the line represents the start of a section.

		A section starts if the first element of the line starts like
		this:

		I. Cash - CNY xxx
		IV. Debt Securities xxx
		VIII. Accruals xxx
		"""
		if isinstance((line[0]), str) and re.match('[IVX]+\.{0,1}\s+', line[0]):
			return True
		else:
			return False
	# end of startOfSection()

	sections = []
	tempSection = []
	for line in filter(notEmptyLine, lines):
		if not startOfSection(line):
			tempSection.append(line)
		else:
			sections.append(tempSection)
			tempSection = [line]

	return sections



def writeCsv(fileName, rows):
	with open(fileName, 'w', newline='') as csvfile:
		file_writer = csv.writer(csvfile)
		for row in rows:
			file_writer.writerow(row)




if __name__ == '__main__':
	from dif_revised.utility import get_current_path
	from os.path import join
	import logging.config
	logging.config.fileConfig('logging.config', disable_existing_loggers=False)

	def writeSection():
		file = join(get_current_path(), 'samples', 
						'CL Franklin DIF 2018-05-28(2nd Revised).xls')
		ws = open_workbook(filename=file).sheet_by_name('Portfolio Val.')
		sections = linesToSections(worksheetToLines(ws))
		writeCsv('htm bond.csv', sections[8])
	# end of writeSection()
	writeSection()





