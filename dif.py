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



class InvalidAccoutingInfo(Exception):
	pass

class ValuationDateNotFound(Exception):
	pass



def readHolding(file):
	ws = open_workbook(filename=file).sheet_by_name('Portfolio Val.')
	sections = linesToSections(worksheetToLines(ws))
	valuationDate = getValuationDate(sections[0])
	records = []
	for section in sections[1:]:
		records = chain(records, sectionToRecords(section))

	def addValuationDate(record):
		record['valuation_date'] = valuationDate
		return record

	return map(addValuationDate, records)



def getValuationDate(lines):
	"""
	lines: [list] a list of lines in the first section, that contains
		fund name, valuation date etc.

	output: a string in "yyyy-mm-dd" representing the date
	"""
	def getDateFromLine(line):
		i = 0
		for item in line:
			if isinstance(item, float):
				i = i + 1
			if i == 2:
				return dateToString(ordinalToDate(item))

		raise ValuationDateNotFound()

	for line in lines:
		if line[0].startswith('Valuation Period'):
			return getDateFromLine(line)


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



def sectionToRecords(lines):
	"""
	lines: [list] a list of lines of a section.

	output: [iterable] a list of records (dictionary object) in the
		section.
	"""
	sectionType = getSectionType(lines[0])
	headerLines, holdingLines = divideSection(lines)
	records = linesToRecords(sectionHeader(headerLines), holdingLines)

	def addSecurityInfo(record):
		record['type'] = sectionType
		return record

	def nonEmptyPosition(record):
		if sectionType in ('bond', 'equity', 'futures') and record['quantity'] in (0, ''):
			return False
		elif record['book_cost'] in (0, ''):
			return False
		else:
			return True

	def toDateString(record):
		if record['type'] == 'futures':
			# FIXME: futures' maturity date is different, don't know what to do
			return record

		for key in ('coupon_start_date', 'maturity_date', 'last_trade_date', 'trade_date'):
			if key in record:
				record[key] = dateToString(ordinalToDate(record[key]))
		return record

	return map(toDateString, map(addSecurityInfo, filter(nonEmptyPosition, records)))



def getSectionType(line):
	"""
	line: the first line of a section

	output: a string representing the type of the section, i.e., cash,
		bond, equity, futures
	"""
	if re.search('\sCash\s', line[0]):
		return 'cash'
	elif re.search('\sBroker Account\s', line[0]):
		return 'broker account cash'
	elif re.search('\sDebt Securities\s', line[0]):
		return 'bond'
	elif re.search('\sEquities\s', line[0]):
		return 'equity'
	elif re.search('\sFutures\s', line[0]):
		return 'futures'
	elif re.search('\sForwards\s', line[0]):
		return 'forwards'
	elif re.search('\sFixed Deposit\s', line[0]):
		return 'fixed deposit cash'
	else:
		logger.error('getSectionType(): invalid type {0}'.format(line[0]))
		raise ValueError()



def linesToRecords(headers, lines):
	"""
	lines: [list] a list of lines in the sub section, the first line being
		the accounting treatment (like (i) held to maturity), the rest are
		holdings

	output: [iterable] a list of records in the sub section, with empty
		positions filtered out.
	"""
	try:
		accounting = getAccountingTreatment(lines[0])
		startingLine = 1
	except InvalidAccoutingInfo:
		accounting = ''
		startingLine = 0

	def lineToRecord(line):
		headerValuePairs = filter(lambda x: x[0] != '', zip(headers, line))
		return {key: value for (key, value) in headerValuePairs}

	def addAccoutingInfo(record):
		record['accounting'] = accounting
		return record

	return map(addAccoutingInfo, map(lineToRecord, lines[startingLine:]))



def getAccountingTreatment(line):
	"""
	line: the first line of a sub section

	output: a string for the sub section's accouting treatment, i..e, htm,
		afs, trading. Or raise an exception if not found.
	"""
	if line[0].startswith('(i) Trading'):
		return 'trading'
	elif line[0].startswith('(i) Held to Maturity'):
		return 'htm'
	else:
		raise InvalidAccoutingInfo()



def divideSection(lines):
	"""
	lines: [list] a list of lines in a section.

	output: [list] a list of sub sections in this section.

	A section can be divided into 2 sub sections:

	sub section 0: header lines (up to 'Description')
	sub section 1: entries (the rest, up to 'total')
	"""
	def findHeaderLines():
		for i in range(len(lines)):
			if lines[i][0].startswith('Description'):
				return i
		raise ValueError('divideSection(): header line not found')

	i = findHeaderLines()
	headerLines = [lines[i-1], lines[i]]	# 2 lines for headers

	def endOfHolding(text):
		return text.startswith('Total (總額)')
	
	holdingLines = []
	for line in lines[i+1:]:
		if endOfHolding(line[0]):
			break
		else:
			holdingLines.append(line)

	return headerLines, holdingLines



def sectionHeader(lines):
	"""
	lines: [list] a list of lines (2 lines) reprenting the headers

	output: [list] a list of header as string
	"""
	headerMap = {
		('', ''): '',
		('項目', 'Description'): 'description',

		# Bond fields
		('票面值', 'Par Amt'):'quantity',
		('幣值', 'CCY'):'currency',
		('上市 (是/否)', 'Listed (Y/N)'):'is_listed',
		('Primary', 'Exchange'):'listed_location',
		('(AVG) FX', 'for TXN'):'fx_on_trade_day',
		('Int.', 'Rate (%)'):'coupon_rate',
		('Int.', 'Start Day'):'coupon_start_date',
		('到期日', 'Maturity'):'maturity_date',
		('Cost', '(%)'):'average_cost',
		('Price', '(%)'):'price',
		('(Amortized)', '(%)'):'amortized_cost',
		('成本價', 'Book Cost'):'book_cost',
		('Int.', 'Bought'):'interest_bought',
		('市價', 'M. Value'):'market_value',
		('Adjusted Value', '(Amortized)'):'amortized_value',
		('應收利息', 'Accr. Int.'):'accrued_interest',
		('Year-End', 'Amortization'):'amortized_gain_loss',
		('Gain/(Loss)', 'M. Value'):'market_gain_loss',
		('FX', 'HKD Equiv.'):'fx_gain_loss_hkd',
		('%', '(Fund)'): 'percentage_of_fund',

		# for trustee Macau fund
		('', 'Listed (Y/N)'):'is_listed',
		('Location', 'of Listed'):'listed_location',
		('FX', 'MOP Equiv.'):'fx_gain_loss_mop',


		# Equity fields
		('股數', 'Share'):'quantity',
		('Location', 'of Listed'):'listed_location',
		('最後交易日', 'Latest V.D.'):'last_trade_date',
		('Avg.', 'Price'):'average_cost',
		('Market', 'Price'):'price',

		# for trustee Macau fund
		('上市 (是/否)', 'Listed (Y/N)'):'is_listed',


		# Cash fields
		('戶口號碼', 'Account No.'): 'account_number',
		('FX', 'for TXN'):'fx_on_trade_day',
		('FX', 'at TXN'):'fx_on_trade_day',
		('市值', 'M. Value'): 'market_value',

		# Futures fields
		('合約數量', 'No. of Contracts'): 'quantity',
		('', 'Long/ Short'): 'long_short',
		('', 'Trade Date'): 'trade_date',

		# Fixed Deposit fields
		('FX', 'at V.D.'): 'fx_on_trade_day',
		('交易日', 'V.D.'): 'trade_date',
		('Int.', 'Rate(%)'): 'interest_rate',


		# headers to ignore (after header column % of fund)
		(2004.0, '購入'): '',
		('Yield', '%'): '',
		(37986.0, 'Market Price'): ''
	}

	try:
		return [headerMap[item] for item in zip(*lines)]
	except KeyError:
		logger.exception('sectionHeader(): header not found')
		raise



def recordsToRows(records):
	"""
	records: a list of position records with the same set of headers, 
		such as HTM bonds, or AFS bonds, equitys, cash entries.

	headers: the headers of the records
	
	output: a list of rows ready to be written to csv, with the first
		row being headers, the rest being values from each record.
		headers.
	"""
	if not records:
		return []

	headers = list(records[0].keys())
	def toValueList(record):
		return [record[header] for header in headers]

	return [headers] + [toValueList(record) for record in records]



def ordinalToDate(ordinal):
	# from: https://stackoverflow.com/a/31359287
	return datetime.fromordinal(datetime(1900, 1, 1).toordinal() + 
									int(ordinal) - 2)



def dateToString(dt):
	return str(dt.year) + '-' + str(dt.month) + '-' + str(dt.day)



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

	def cashSection():
		file = join(get_current_path(), 'samples', 
						'CL Franklin DIF 2018-05-28(2nd Revised).xls')
		ws = open_workbook(filename=file).sheet_by_name('Portfolio Val.')
		sections = linesToSections(worksheetToLines(ws))
		return sections[1]

	def htmSection():
		file = join(get_current_path(), 'samples', 
						'CL Franklin DIF 2018-05-28(2nd Revised).xls')
		ws = open_workbook(filename=file).sheet_by_name('Portfolio Val.')
		sections = linesToSections(worksheetToLines(ws))
		return sections[8]
	# end of htmSection()
	# writeCsv('htm section.csv', htmSection())

	def equitySection():
		file = join(get_current_path(), 'samples', 
						'CL Franklin DIF 2018-05-28(2nd Revised).xls')
		ws = open_workbook(filename=file).sheet_by_name('Portfolio Val.')
		sections = linesToSections(worksheetToLines(ws))
		return sections[14]

	# writeCsv('equity section.csv', equitySection())

	def forwardsSection():
		file = join(get_current_path(), 'samples', 
						'CL Franklin DIF 2018-05-28(2nd Revised).xls')
		ws = open_workbook(filename=file).sheet_by_name('Portfolio Val.')
		sections = linesToSections(worksheetToLines(ws))
		return sections[16]

	def fixedDepositSection():
		file = join(get_current_path(), 'samples', 
						'CL Franklin DIF 2018-05-28(2nd Revised).xls')
		ws = open_workbook(filename=file).sheet_by_name('Portfolio Val.')
		sections = linesToSections(worksheetToLines(ws))
		return sections[18]

	def futuresSection():
		file = join(get_current_path(), 'samples', 
						'CL Franklin DIF 2018-05-28(2nd Revised).xls')
		ws = open_workbook(filename=file).sheet_by_name('Portfolio Val.')
		sections = linesToSections(worksheetToLines(ws))
		return sections[19]

	# writeCsv('equity section.csv', equitySection())


	def htmHeaderLines():
		headerLines, holdingLines = divideSection(htmSection())
		return headerLines

	def htmHoldingLines():
		headerLines, holdingLines = divideSection(htmSection())
		return holdingLines

	# writeCsv('htm subsection header.csv', htmHeaderLines())
	# writeCsv('htm subsection holding.csv', htmHoldingLines())

	# print(sectionHeader(htmHeaderLines()))

	def htmRecords():
		return sectionToRecords(htmSection())

	def equityRecords():
		return sectionToRecords(equitySection())

	# writeCsv('htm records.csv', recordsToRows(list(htmRecords())))
	# writeCsv('equity records.csv', recordsToRows(list(equityRecords())))
	# writeCsv('forwards records.csv', recordsToRows(list(sectionToRecords(forwardsSection()))))
	# writeCsv('fixed deposit records.csv', recordsToRows(list(sectionToRecords(fixedDepositSection()))))
	# writeCsv('futures records.csv', recordsToRows(list(sectionToRecords(futuresSection()))))
	# writeCsv('cash records.csv', recordsToRows(list(sectionToRecords(cashSection()))))

	def allRecords():
		file = join(get_current_path(), 'samples', 
						'CL Franklin DIF 2018-05-28(2nd Revised).xls')
		return readHolding(file)

	def tradingBond(record):
		if record['type'] == 'bond' and record['accounting'] == 'trading':
			return True
		return False

	writeCsv('all trading bond records.csv', recordsToRows(list(filter(tradingBond, allRecords()))))