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



def toSubSections(section):
	"""
	section: [list] a list of lines in a section.

	output: [list] a list of sub sections in this section.

	The structure of a section is like below:

	sub section 0: header lines (up to 'Description')
	sub section 1: holdings (in between '(i) held to maturity' and 'total')
	sub section 2: holdings (in between '(ii) held to maturity' and 'total')
	etc.

	In the above, 'held to maturiy' can be replaced by 'trading' or other
	accounting treatment.
	"""
	def findHeaderLines():
		for i in range(len(section)):
			if section[i][0].startswith('Description'):
				return i
		raise ValueError('toSubSections(): header line not found')

	i = findHeaderLines()
	subSections = [[section[i-2], section[i-1], section[i]]]

	def startOfHolding(text):
		"""
		Tell whether the text string indicates start of a holding sub section
		"""
		return bool(re.match('\([ivx]+\)\s', text))

	def endOfHolding(text):
		"""
		Tell whether the text string indicates end of a holding sub section
		"""
		return text.startswith('Total (總額)')
	
	tempSub = []
	for line in section[i+1:]:
		if tempSub == [] and not startOfHolding(line[0]):
			continue
		elif tempSub == [] and startOfHolding(line[0]):
			tempSub.append(line)
		elif tempSub != [] and endOfHolding(line[0]):
			subSections.append(tempSub)
			tempSub = []
		else:
			tempSub.append(line)

	return subSections



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

	def htmSection():
		file = join(get_current_path(), 'samples', 
						'CL Franklin DIF 2018-05-28(2nd Revised).xls')
		ws = open_workbook(filename=file).sheet_by_name('Portfolio Val.')
		sections = linesToSections(worksheetToLines(ws))
		return sections[8]
	# end of htmSection()
	# writeCsv('htm section.csv', htmSection())

	def htmSubSectionHeader():
		subSections = toSubSections(htmSection())
		return subSections[0]

	def htmSubSectionHolding():
		subSections = toSubSections(htmSection())
		return subSections[1]

	# writeCsv('htm subsection header.csv', htmSubSectionHeader())
	# writeCsv('htm subsection holding.csv', htmSubSectionHolding())



