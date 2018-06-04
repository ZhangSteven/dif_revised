# coding=utf-8
# 
# Use holding records from dif.py, to generate csv files for Geneva
# reconciliation purpose. 
# 

from dif_revised.dif import readFile, recordsToRows, writeCsv
from functools import reduce



def open_dif(inputFile, portValues, outputDir, prefix):
	"""
	Read an input file, write 3 output csv files, namely,

		1. HTM positions
		2. AFS positions
		3. cash

	the output file should be written to outputDir, and output file
	names should follow the below convention:

	<prefix>_yyyy-mm-dd_cash.csv
	<prefix>_yyyy-mm-dd_afs_positions.csv
	<prefix>_yyyy-mm-dd_htm_positions.csv

	return 3 string, for the full file path to the 3 output csv files.

	The interface is exactly the same as the old DIF package's
	open_dif.open_dif() function, to replace it.
	"""
	from os.path import join
	records, valuationSummary = readFile(inputFile)
	portfolioId = records[0]['portfolio']
	if portfolioId == '19437':
		prefix = 'DIF_'


	valuationDate = records[0]['valuation_date']
	cashCsvFile = join(outputDir, prefix + valuationDate + '_cash.csv')
	afsCsvFile = join(outputDir, prefix + valuationDate + '_afs_positions.csv')
	htmCsvFile = join(outputDir, prefix + valuationDate + '_htm_positions.csv')

	writeCashCsv(cashCsvFile, records)
	writeAfsCsv(afsCsvFile, records)
	writeHtmCsv(htmCsvFile, records)

	portValues['valuation_date'] = valuationDate
	portValues['portfolio'] = portfolioId
	for (k, v) in valuationSummary.items():
		portValues[k] = v

	return [cashCsvFile, afsCsvFile, htmCsvFile]



def writeCashCsv(file, records, delimiter='|'):
	"""
	records: the holding records of the portfolio, including cash, bond,
		equity, futures, etc.

	file: the output csv file

	output: no return value, the function writes cash records to
		the output csv file with headers needed by Geneva reconciliation.
		The cash records include bank cash and futures broker account cash.
	"""
	cashHeaders = ['portfolio', 'custodian', 'date', 'account_type',
					'account_num', 'currency', 'balance', 'fx_rate',
					'local_currency_equivalent']

	bankMap = {	# map bank name to custodian name
		'Citibank': 'CITI',
		'ICBC (Macau) Ltd': 'ICBCMACAU',
		'JPMorgan Chase Bank, N.A.': 'JPM',
		'Bank of China Ltd. (Macau Branch)': 'BOCMACAU',
		'Luso International Banking Ltd.': 'LUSO',
		'China Guangfa Bank Co., Ltd Macau Branch': 'GUANGFA_MACAU',
		'Bank of China (HK)': 'BOCHK'
	}

	def cash(record):
		if record['type'] in ('cash', 'broker account cash'):
			return True
		return False

	def toCashRecords(record):
		r = {}
		r['date'] = record['valuation_date']
		if record['type'] == 'cash':
			try:
				r['custodian'] = bankMap[record['bank']]
			except KeyError:
				raise KeyError('toCashRecords(): {0} map custodian failed'.format(record))
		elif record['type'] == 'broker account cash':
			r['custodian'] = record['bank']

		r['account_num'] = record['account_number']
		r['balance'] = record['book_cost']
		r['fx_rate'] = record['exchange_rate']
		r['local_currency_equivalent'] = record['exchange_rate'] * record['book_cost']

		for header in ['portfolio', 'account_type', 'currency']:
			r[header] = record[header]

		return r

	def consolidateCash(cashList, cash):
		"""
		Merge cash entries of the same bank, of the same currency to one
		entry.
		"""
		def findMatchingCash(cashList, cash):
			for i in range(len(cashList)):
				c = cashList[i]
				if c['custodian'] == cash['custodian'] and c['currency'] == cash['currency']:
					return i

			return -1

		def mergeCashToList(cashList, index, cash):
			if index == -1:
				cashList.append(cash)
			else:
				cashList[index]['balance'] = cashList[index]['balance'] + cash['balance']
				cashList[index]['local_currency_equivalent'] = cashList[index]['local_currency_equivalent'] + cash['local_currency_equivalent']
			return cashList

		i = findMatchingCash(cashList, cash)
		return mergeCashToList(cashList, i, cash)

	writeCsv(file,
		recordsToRows(reduce(consolidateCash, map(toCashRecords, filter(cash, records)), []), 
						cashHeaders),
		delimiter)



def writeAfsCsv(file, records, delimiter='|'):
	"""
	records: the holding records of the portfolio, including cash, bond,
		equity, futures, etc.

	file: the output csv file

	output: no return value, the function writes all non HTM records to
		the output csv file with headers needed by Geneva reconciliation.
	"""
	afsHeaders = ['portfolio', 'date', 'custodian', 'ticker', 'isin',
				'bloomberg_figi', 'name', 'currency', 'accounting_treatment',
				'quantity', 'average_cost', 'price', 'book_cost', 'market_value',
				'market_gain_loss', 'fx_gain_loss']

	def afsPosition(record):
		if record['type'] == 'equity' or \
			(record['type'] == 'bond' and record['accounting'] != 'htm'):
			return True
		return False

	def toAfsRecords(record):
		"""
		map an avaible for sale (AFS) or Trading position (can be either equity
		or bond) to the record ready to be written to the csv.
		"""
		r = {}
		r['date'] = record['valuation_date']
		r['geneva_investment_id'] = ''
		r['bloomberg_figi'] = ''
		r['accounting_treatment'] = record['accounting'].upper()
		r['name'] = record['description']
		try:
			r['fx_gain_loss'] = record['fx_gain_loss_hkd']
		except KeyError:
			r['fx_gain_loss'] = record['fx_gain_loss_mop']

		for header in afsHeaders:
			if header in ('date', 'geneva_investment_id', 'bloomberg_figi',
							'accounting_treatment', 'name', 'fx_gain_loss'):
				pass
			elif header in ('ticker', 'isin'):
				try:
					r[header] = record[header]
				except KeyError:
					r[header] = ''
			else:
				r[header] = record[header]

		return r
	# end of toAfsRecords()
	writeCsv(file,
		recordsToRows(list(map(toAfsRecords, filter(afsPosition, records))), afsHeaders),
		delimiter)



def writeHtmCsv(file, records, delimiter='|'):
	"""
	records: the holding records of the portfolio, including cash, bond,
		equity, futures, etc.

	file: the output csv file

	output: no return value, the function writes the HTM bond records to
		the output csv file with headers needed by Geneva reconciliation.
	"""
	htmHeaders = ['portfolio', 'date', 'custodian', 'geneva_investment_id', 'isin',
				'bloomberg_figi', 'name', 'currency', 'accounting_treatment',
				'par_amount', 'is_listed', 'listed_location', 'fx_on_trade_day',
				'coupon_rate', 'coupon_start_date', 'maturity_date', 'average_cost',
				'amortized_cost', 'book_cost', 'interest_bought', 'amortized_value',
				'accrued_interest', 'amortized_gain_loss', 'fx_gain_loss']

	def htmPosition(record):
		if record['type'] == 'bond' and record['accounting'] == 'htm':
			return True
		return False

	def toHtmRecords(record):
		"""
		map a htm bond record to the record ready to be written to the csv.
		"""
		r = {}
		r['date'] = record['valuation_date']
		r['geneva_investment_id'] = record['isin'] + ' HTM'
		r['bloomberg_figi'] = ''
		r['name'] = record['description']
		r['accounting_treatment'] = record['accounting'].upper()
		r['par_amount'] = record['quantity']
		try:
			r['fx_gain_loss'] = record['fx_gain_loss_hkd']
		except KeyError:
			r['fx_gain_loss'] = record['fx_gain_loss_mop']

		for header in htmHeaders:
			if header in ('date', 'geneva_investment_id', 'bloomberg_figi',
							'name', 'accounting_treatment', 'par_amount', 
							'fx_gain_loss'):
				continue

			r[header] = record[header]

		return r
	# end of toHtmRecords()
	writeCsv(file, 
		recordsToRows(list(map(toHtmRecords, filter(htmPosition, records))), htmHeaders),
		delimiter)



if __name__ == '__main__':
	from dif_revised.utility import get_current_path
	from os.path import join
	import logging.config
	logging.config.fileConfig('logging.config', disable_existing_loggers=False)

	def getRecords():
		file = join(get_current_path(), 'samples', 
						'CL Franklin DIF 2018-05-28(2nd Revised).xls')
		return readFile(file)

	def getRecordsBal():
		file = join(get_current_path(), 'samples', 
						'CLM BAL 2017-07-27.xls')
		return readFile(file)

	def getRecordsGnt():
		file = join(get_current_path(), 'samples', 
						'CLM GNT 2017-10-25.xls')
		return readFile(file)

	# records, valuationSummary = getRecords()
	# writeHtmCsv('htm holding.csv', records, '|')
	# writeAfsCsv('afs holding.csv', records, '|')
	# writeCashCsv('cash.csv', records, '|')

	path = join('C:\\temp\\Reconciliation\\')
	portValues = {}
	output = open_dif(join(path, 'DIF', 'CL Franklin DIF 2018-05-28(2nd Revised).xls'),
						portValues,
						join(path, 'result'), 'dif')
	print(output)
	print(portValues)