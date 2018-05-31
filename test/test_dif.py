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



class TestDif(unittest2.TestCase):
	def __init__(self, *args, **kwargs):
		super(TestDif, self).__init__(*args, **kwargs)

	def testHtm(self):
		file = join(get_current_path(), 'samples', 
						'CL Franklin DIF 2018-05-28(2nd Revised).xls')
		records = list(filter(htmBond, readHolding(file)))
		self.assertEqual(len(records), 4)
		self.verifyHtmBond(records[0])



	def verifyHtmBond(self, record):
		self.assertEqual(record['valuation_date'], '2018-5-28')
		self.assertEqual(record['description'], '(USY9896RAB79) Zoomlion HK SPV Co Ltd 6.125%')
		self.assertEqual(record['quantity'], 13700000)
		self.assertEqual(record['coupon_rate'], 0.06125)
		self.assertEqual(record['coupon_start_date'], '2017-12-20')
		self.assertEqual(record['maturity_date'], '2022-12-20')
		self.assertAlmostEqual(record['average_cost'], 96.4166058)
		self.assertAlmostEqual(record['amortized_cost'], 97.2761909)
