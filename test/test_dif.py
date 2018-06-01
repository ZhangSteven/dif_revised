# coding=utf-8
# 

import unittest2
from os.path import join
from dif_revised.utility import get_current_path
from dif_revised.dif import readHolding



def htmBond(record):
	if record['type'] == 'bond' and record['accounting'] == 'htm':
		return True
	return False

def tradingBond(record):
	if record['type'] == 'bond' and record['accounting'] == 'trading':
		return True
	return False

def equity(record):
	if record['type'] == 'equity':
		return True
	return False

def cash(record):
	if 'cash' in record['type']:#cash, broker account cash, fixed deposit cash
		return True
	return False



class TestDif(unittest2.TestCase):
	def __init__(self, *args, **kwargs):
		super(TestDif, self).__init__(*args, **kwargs)

	@classmethod
	def setUpClass(TestDif):
		"""
		Called only once before all tests
		"""
		file = join(get_current_path(), 'samples', 
						'CL Franklin DIF 2018-05-28(2nd Revised).xls')
		TestDif.records = readHolding(file)

	@classmethod
	def tearDownClass(TestDif):
		pass



	def testHtm(self):
		records = list(filter(htmBond, TestDif.records))
		self.assertEqual(len(records), 4)
		self.verifyHtmBond(records[0])



	def testTradingBond(self):
		records = list(filter(tradingBond, TestDif.records))
		self.assertEqual(len(records), 65)
		self.verifyTradingBond(records[0])



	def testEquity(self):
		records = list(filter(equity, TestDif.records))
		self.assertEqual(len(records), 14)
		self.verifyEquity(records[11])



	def testCash(self):
		records = list(filter(cash, TestDif.records))
		self.assertEqual(len(records), 4)
		self.verifyCash(records[3])



	def verifyHtmBond(self, record):
		self.assertEqual(record['valuation_date'], '2018-5-28')
		self.assertEqual(record['description'], '(USY9896RAB79) Zoomlion HK SPV Co Ltd 6.125%')
		self.assertEqual(record['isin'], 'USY9896RAB79')
		self.assertAlmostEqual(record['exchange_rate'], 7.8452, 6)
		self.assertEqual(record['quantity'], 13700000)
		self.assertEqual(record['coupon_rate'], 0.06125)
		self.assertEqual(record['coupon_start_date'], '2017-12-20')
		self.assertEqual(record['maturity_date'], '2022-12-20')
		self.assertAlmostEqual(record['average_cost'], 96.4166058)
		self.assertAlmostEqual(record['amortized_cost'], 97.2761909)



	def verifyTradingBond(self, record):
		self.assertEqual(record['portfolio'], '19437')
		self.assertEqual(record['isin'], 'XS1376566714')
		self.assertAlmostEqual(record['exchange_rate'], 7.8452, 6)
		self.assertEqual(record['quantity'], 5000000)
		self.assertAlmostEqual(record['coupon_rate'], 0.0555, 6)
		self.assertEqual(record['maturity_date'], '2021-4-14')
		self.assertAlmostEqual(record['average_cost'], 96.618, 6)
		self.assertEqual(record['market_value'], 1433350)
		self.assertEqual(record['market_gain_loss'], -3397550)



	def verifyEquity(self, record):
		"""
		It's a bond treated as equity
		"""
		self.assertEqual(record['valuation_date'], '2018-5-28')
		self.assertEqual(record['ticker'], 'XS1328130197')
		self.assertAlmostEqual(record['exchange_rate'], 7.8452, 6)
		self.assertEqual(record['quantity'], 3924000)
		self.assertEqual(record['currency'], 'USD')
		self.assertEqual(record['last_trade_date'], '2018-1-3')
		self.assertAlmostEqual(record['average_cost'], 104.332436, 6)
		self.assertAlmostEqual(record['price'], 99.268, 6)
		self.assertAlmostEqual(record['percentage_of_fund'], 0.71, 6)



	def verifyCash(self, record):
		"""
		It's the broker account cash
		"""
		self.assertEqual(record['description'], 'Morgan Stanley - Broker Account')
		self.assertEqual(record['currency'], 'USD')
		self.assertEqual(record['account_number'], '045621UE7')
		self.assertAlmostEqual(record['book_cost'], 3938502.58, 6)
		self.assertAlmostEqual(record['exchange_rate'], 7.8452, 6)