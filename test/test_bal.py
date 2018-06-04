# coding=utf-8
# 

import unittest2
from os.path import join
from xlrd import open_workbook
from dif_revised.utility import get_current_path
from dif_revised.dif import readHolding, readSummary, validate



def equity(record):
	if record['type'] == 'equity':
		return True
	return False

def bondAfs(record):
	if record['type'] == 'bond' and record['accounting'] == 'afs':
		return True
	return False

def cash(record):
	if record['type'] in ('cash', 'broker account cash'):
		return True
	return False



class TestBal(unittest2.TestCase):
	def __init__(self, *args, **kwargs):
		super(TestBal, self).__init__(*args, **kwargs)

	@classmethod
	def setUpClass(TestBal):
		"""
		Called only once before all tests
		"""
		file = join(get_current_path(), 'samples', 
						'CLM BAL 2017-07-27.xls')
		wb = open_workbook(filename=file)
		TestBal.balrecords = readHolding(wb.sheet_by_name('Portfolio Val.'))
		TestBal.balsummary = readSummary(wb.sheet_by_name('Portfolio Sum.'))

		file = join(get_current_path(), 'samples', 
						'CLM GNT 2017-10-25.xls')
		wb = open_workbook(filename=file)
		TestBal.gntrecords = readHolding(wb.sheet_by_name('Portfolio Val.'))
		TestBal.gntsummary = readSummary(wb.sheet_by_name('Portfolio Sum.'))

	@classmethod
	def tearDownClass(TestBal):
		pass



	def testEquityBal(self):
		records = list(filter(equity, TestBal.balrecords))
		self.assertEqual(len(records), 14)
		record = records[0]
		self.assertEqual(record['portfolio'], '30004')
		self.assertEqual(record['valuation_date'], '2017-7-27')
		self.assertEqual(record['ticker'], '522 HK')
		self.assertFalse('isin' in record)
		self.assertAlmostEqual(record['exchange_rate'], 1.03, 4)
		self.assertEqual(record['quantity'], 4100)
		self.assertEqual(record['currency'], 'HKD')
		self.assertEqual(record['last_trade_date'], '2017-7-27')
		self.assertAlmostEqual(record['average_cost'], 121.8225, 4)
		self.assertAlmostEqual(record['price'], 101.3, 6)
		self.assertAlmostEqual(record['percentage_of_fund'], 0.76, 6)



	def testCashBal(self):
		records = list(filter(cash, TestBal.balrecords))
		self.assertEqual(len(records), 7)
		record = records[0]
		self.assertEqual(record['bank'], 'Citibank')
		self.assertEqual(record['account_type'], 'Saving Account')
		self.assertEqual(record['currency'], 'HKD')
		self.assertAlmostEqual(record['book_cost'], 7278.97)
		self.assertAlmostEqual(record['exchange_rate'], 1.030046455)
		self.assertFalse('custodian' in record)
		self.assertEqual(record['portfolio'], '30004')


	def testValidateBal(self):
		try:
			validate(TestBal.balrecords, TestBal.balsummary)
		except:
			self.fail('validate() failed')



	def testBondAfsGnt(self):
		records = list(filter(bondAfs, TestBal.gntrecords))
		self.assertEqual(len(records), 11)
		record = records[0]
		self.assertEqual(record['portfolio'], '30003')
		self.assertEqual(record['valuation_date'], '2017-10-25')
		self.assertEqual(record['isin'], 'XS1389124774')
		self.assertEqual(record['quantity'], 22000000)
		self.assertAlmostEqual(record['coupon_rate'], 0.0605)
		self.assertEqual(record['maturity_date'], '2056-2-15')
		self.assertAlmostEqual(record['price'], 108.069)
		self.assertAlmostEqual(record['accrued_interest'], 928002.78)



	def testCashGnt(self):
		records = list(filter(cash, TestBal.gntrecords))
		self.assertEqual(len(records), 8)
		record = records[7]
		self.assertEqual(record['bank'], 'Luso International Banking Ltd.')
		self.assertEqual(record['account_type'], '')
		self.assertEqual(record['currency'], 'USD')
		self.assertAlmostEqual(record['book_cost'], 36.67)
		self.assertAlmostEqual(record['exchange_rate'], 8.0366209)
		self.assertFalse('custodian' in record)
		self.assertEqual(record['portfolio'], '30003')


	def testValidateGnt(self):
		try:
			validate(TestBal.gntrecords, TestBal.gntsummary)
		except:
			self.fail('validate() failed')
