# coding=utf-8
# 
# Read holdings from China Life trustee's DIF excel file. It is actually
# a rewritten of the old DIF package, with a more clear structure. Structure
# and code are similar to clamc_trustee package.
# 
# known bug: futures' position's maturity date is of different format and
# 	not converted. search for 'FIXME'

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

class ExchangeRateNotFound(Exception):
	pass

class InconsistentRecordSum(Exception):
	pass

class RecordTypeNotSupported(Exception):
	pass

class UnderlyingTypeNotFound(Exception):
	pass



def readFile(file):
	"""
	ws: the full path to the China Life trustee's DIF file.

	output: 
	"""
	wb = open_workbook(filename=file)
	records = readHolding(wb.sheet_by_name('Portfolio Val.'))
	summary = readSummary(wb.sheet_by_name('Portfolio Sum.'))
	validate(records, summary)
	return records



def readSummary(ws):
	"""
	ws: the excel worksheet for DIF holdings.

	output: [dictionary] a summary containing the portfolio's total values,
		such as cash, bond, equity, futures, fixed deposit
	"""
	lines = worksheetToLines(ws)
	for i in range(0, len(lines)):	# find where summary starts
		if lines[i][0] == 'Current Portfolio':
			break

	summary = {}
	typeMap = {
		'Cash (現金)': 'cash',
		'Debt Securities (債務票據)': 'bond',
		'Debt Amortization (債務攤銷)': 'bond amortization',
		'Equities (股票)': 'equity',
		'Fixed Deposit (定期存款)': 'fixed deposit',
		'Futures (期貨合約)': 'futures'
	}

	def readValue(line):
		i = 0
		for item in line:
			if isinstance(item, float):
				i = i + 1
			if i == 2:
				return item

	for line in lines[i+1:i+14]:
		try:
			summary[typeMap[line[0]]] = readValue(line)
		except KeyError:
			pass

	summary['bond'] = summary['bond'] + summary.pop('bond amortization')
	return summary



def validate(records, summary):
	def recordValue(record):
		if record['type'] in ('cash', 'broker account cash', 'fixed deposit'):
			return record['book_cost']
		
		elif record['type'] == 'bond':
			if record['accounting'] == 'htm':
				return record['quantity'] / 100 * record['amortized_cost'] + record['accrued_interest']
			else:
				return record['quantity'] / 100 * record['price'] + record['accrued_interest']

		elif record['type'] == 'futures':
			return record['market_gain_loss']

		elif record['type'] == 'equity':
			# if record['underlying'] == 'bond': # bond treated as equity
			# 	return record['quantity'] * record['price'] / 100
			# else:
			# 	return record['quantity'] * record['price']
			return record['market_value']

		else:
			raise RecordTypeNotSupported('{0}'.format(record))

	def sumUp(total, record):
		try:
			return total + record['exchange_rate'] * recordValue(record)
		except:
			print(record)
			import sys
			sys.exit(1)

	for recordType in summary:
		if recordType == 'cash':
			tempRecords = filter(lambda r: r['type'] in ('cash', 'broker account cash'), records)
		else:
			tempRecords = filter(lambda r: r['type'] == recordType, records)

		diff = summary[recordType] - reduce(sumUp, tempRecords, 0)
		if abs(diff) > 0.2:
			raise InconsistentRecordSum('validate(): diff {0} for {1}'.format(diff, recordType))



def readHolding(ws):
	"""
	ws: the excel worksheet for DIF holdings.

	output: [list] a list of records in DIF portfolio, including cash,
		bond, equity, forwards, futures, fixed deposit etc.
	"""
	sections = linesToSections(worksheetToLines(ws))
	valuationDate, portfolio, custodian = getPortfolioInfo(sections[0])
	records = []
	for section in sections[1:]:
		records = chain(records, sectionToRecords(section))

	def addPortfolioInfo(record):
		record['valuation_date'] = valuationDate
		record['portfolio'] = portfolio
		record['custodian'] = custodian
		return record

	return list(map(addPortfolioInfo, records))



def getPortfolioInfo(lines):
	"""
	lines: [list] a list of lines in the first section, that contains
		fund name, valuation date etc.

	output: 3 values for the portfolio,
		valuationDate: a string in "yyyy-mm-dd" for the valuation date
		portfolio: a string for the portfolio id
		custodian: a string for the portfolio's custodian bank
	"""
	def getDateFromLine(line):
		i = 0
		for item in line:
			if isinstance(item, float):
				i = i + 1
			if i == 2:
				return dateToString(ordinalToDate(item))

		raise ValuationDateNotFound()

	def getPortfolioId(text):
		"""
		text: a string containing the portfolio's name
		"""
		if 'CHINA LIFE MACAU BRANCH BALANCED ' in text:
			return '30004'
		elif 'CHINA LIFE MACAU BRANCH GUARANTEE ' in text:
			return '30003'
		elif 'CHINA LIFE MACAU BRANCH GROWTH ' in text:
			return '30005'
		elif 'Diversified Income Fund' in text:
			return '19437'
		else:
			raise ValueError('getPortfolioId(): unsupported portfolio name {0}'.format(text))

	for line in lines:
		if line[0].startswith('Valuation Period'):
			valuationDate = getDateFromLine(line)
		elif line[0].startswith('Fund Name'):
			portfolio = getPortfolioId(line[0])

	custodianMap = {
		'30003':'ICBCMACAU',
		'30004':'ICBCMACAU',
		'30005':'ICBCMACAU',
		'19437':'BOCHK'
	}

	try:
		return valuationDate, portfolio, custodianMap[portfolio]
	except:
		logger.exception('getPortfolioInfo()')
		raise



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
	sectionType, sectionCurrency = getSectionInfo(lines[0])
	headerLines, holdingLines, trailLines = divideSection(lines)
	records = linesToRecords(sectionHeader(headerLines), holdingLines)
	exchangeRate = getExchangeRate(trailLines)

	def extractId(text):
		m = re.match('\(([A-Z0-9]{5,12})\)', text)
		if m:
			return m.group(1)
		else:
			logger.error('extractId(): find id failed.')
			raise ValueError('text=\'{0}\''.format(text))

	def convertTicker(text):
		"""
		in DIF, the following is used to identify an equity (H0939), we
		convert them to a ticker format more widely used.

		H0939: 939 HK
		H1186: 1186 HK
		N0011: 11 HK
		N2388: 2388 HK
		"""
		m = re.match('[HN]([0-9]{4})', text)
		if m:
			return str(int(m.group(1))) + ' HK'	# remove leading zeros
		else:
			logger.warning('convertTicker(): {0} is not converted'.format(text))
			return text

	def addSecurityInfo(record):
		record['type'] = sectionType
		if sectionCurrency and not 'currency' in record:
			record['currency'] = sectionCurrency
		if sectionType in ('bond', 'equity'):
			securityId = extractId(record['description'])
			if sectionType == 'bond' or (sectionType == 'equity' and len(securityId) == 12):
				idType = 'isin'
			else:
				idType = 'ticker'
				securityId = convertTicker(securityId)

			record[idType] = securityId
		if exchangeRate:
			record['exchange_rate'] = exchangeRate
		return record

	def nonEmptyPosition(record):
		if not 'quantity' in record and not 'book_cost' in record:
			return False

		if 'quantity' in record and record['quantity'] in (0, '') or \
			'book_cost' and record['book_cost'] in (0, ''):
			return False
		
		return True 	# either quantity or book_cost is non-trival

	def toDateString(record):
		if record['type'] == 'futures':
			# FIXME: futures' maturity date is of different format, 
			# cannot use the below ordinalToDate() function. So skip for now.
			return record

		for key in ('coupon_start_date', 'maturity_date', 'last_trade_date', 'trade_date'):
			if key in record:
				"""
				In most cases, the date from Excel is read in as a float
				number. However, in rare cases, it can be a string. So we
				handle them separately.
				"""
				if isinstance(record[key], float):
					record[key] = dateToString(ordinalToDate(record[key]))
				else:
					record[key] = convertStringDate(record[key])
		return record

	return map(toDateString, map(addSecurityInfo, filter(nonEmptyPosition, records)))



def getSectionInfo(line):
	"""
	line: the first line of a section

	output: two strings, one for the type of the section and the other
		for the currency of the section.

		type of the section: cash, bond, equity, futures, etc.
		currency of the section: currency of the section, if not found
			then return an empty string.
	"""
	def getSectionType(line):
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
			raise ValueError('getSectionType(): invalid type {0}'.format(line[0]))

	def getSectionCurrency(line):
		m = re.search('[IVX]+[A-Za-z\s\.]+- ([A-Za-z$]{3})', line[0])
		if m:
			return m.group(1).upper().replace('$', 'D')	# HK$ mapped to HKD
		else:
			logger.warning('getSectionCurrency(): cannot get currency from {0}'.format(line[0]))
			return ''

	return getSectionType(line), getSectionCurrency(line)



def getExchangeRate(lines):
	"""
	lines: lines in a section that may contain exchange rate info.

	output: (float) exchange rate
	"""
	for line in lines:
		if line[0].startswith('Exchange Rate'):
			break

	for item in line[1:]:
		if isinstance(item, float) and item > 0:
			return item

	logger.warning('getExchangeRate(): FX not found in line {0}'.format(line))
	return ''



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
	text = line[0].lower()
	if 'trading' in text:
		return 'trading'
	elif 'held to maturity' in text or 'amortized cost' in text:
		return 'htm'
	elif 'available for sales' in text or 'market value' in text:
		return 'afs'
	else:
		raise InvalidAccoutingInfo()



def divideSection(lines):
	"""
	lines: [list] a list of lines in a section.

	output: 3 sub lists divided from lines:
		header lines: containing headers (2 lines)
		holding lines: containing positions
		remaining lines: containing total, exchange rate (if any),
			etc.

	A section can be divided into 2 sub sections:

	sub section 0: header lines (up to 'Description')
	sub section 1: entries (the rest, up to 'total')
	"""
	def findHeaderLines():
		for i in range(len(lines)):
			if lines[i][0].startswith('Description'):
				return i
		raise ValueError('divideSection(): header line not found')

	hIndex = findHeaderLines()

	def endOfHolding(text):
		return text.startswith('Total (總額)')
	
	for i in range(hIndex+1, len(lines)):
		if endOfHolding(lines[i][0]):
			break

	return lines[hIndex-1:hIndex+1], lines[hIndex+1:i], lines[i:]



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
		('幣值', 'CCY'):'currency',
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
		(37986.0, 'Market Price'): '',
		('', 'checking'): '',
		(0.0, 0.0): '',
	}

	try:
		return [headerMap[item] for item in zip(*lines)]
	except KeyError:
		logger.exception('sectionHeader(): header not found')
		raise



def recordsToRows(records, headers=None):
	"""
	records: a list of position records with the same set of headers, 
		such as HTM bonds, or AFS bonds, equitys, cash entries.

	headers: the headers of the records, if provided.
	
	output: a list of rows ready to be written to csv, with the first
		row being headers, the rest being values from each record.
		headers.
	"""
	if not records:
		return []
	if not headers:
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



def convertStringDate(dtString):
	"""
	For trustee Excel files, based on experience, if the date is read in
	as a string, then it is of 'dd/mm/yyyy' format. We just conver it
	to a format as 'yyyy-mm-dd'. 
	"""
	m = re.match('(\d{1,2})/(\d{1,2})/(\d{4})', dtString)
	if m:
		return m.group(3) + '-' + m.group(2) + '-' + m.group(1)
	else:
		raise ValueError('convertStringDate(): {0} cannot be converted'.format(dtString))



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