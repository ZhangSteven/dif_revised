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
		TestBal.records = readHolding(wb.sheet_by_name('Portfolio Val.'))
		TestBal.summary = readSummary(wb.sheet_by_name('Portfolio Sum.'))

	@classmethod
	def tearDownClass(TestBal):
		pass



	def testEquity(self):
		records = list(filter(equity, TestBal.records))
		self.assertEqual(len(records), 14)
		record = records[0]
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



	def testValidate(self):
		try:
			validate(TestBal.records, TestBal.summary)
		except:
			self.fail('validate() failed')
