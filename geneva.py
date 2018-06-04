# coding=utf-8
# 
# Use holding records from dif.py, then generate csv files for Geneva
# reconciliation purpose. 
# 

from dif_revised.dif import readFile, recordsToRows, writeCsv



def writeCashCsv(file, records):
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

	def cash(record):
		if record['type'] in ('cash', 'broker account cash'):
			return True
		return False

	def toCashRecords(record):
		


def writeAfsCsv(file, records):
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
		recordsToRows(list(map(toAfsRecords, filter(afsPosition, records))), afsHeaders))



def writeHtmCsv(file, records):
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
		recordsToRows(list(map(toHtmRecords, filter(htmPosition, records))), htmHeaders))



if __name__ == '__main__':
	from dif_revised.utility import get_current_path
	from os.path import join
	import logging.config
	logging.config.fileConfig('logging.config', disable_existing_loggers=False)

	def getRecords():
		file = join(get_current_path(), 'samples', 
						'CL Franklin DIF 2018-05-28(2nd Revised).xls')
		return readFile(file)

	difRecords = getRecords()
	writeHtmCsv('dif htm.csv', difRecords)
	writeAfsCsv('dif afs.csv', difRecords)