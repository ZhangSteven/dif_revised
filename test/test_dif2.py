# coding=utf-8
# 

import unittest2
from os.path import join
from xlrd import open_workbook
from dif_revised.utility import get_current_path
from dif_revised.dif import readHolding, readSummary, validate



class TestDif2(unittest2.TestCase):
	def __init__(self, *args, **kwargs):
		super(TestDif2, self).__init__(*args, **kwargs)



	def testAll(self):
		file = join(get_current_path(), 'samples', 
						'CL Franklin DIF 2018-07-24.xls')
		wb = open_workbook(filename=file)
		records = readHolding(wb.sheet_by_name('Portfolio Val.'))
		summary = readSummary(wb.sheet_by_name('Portfolio Sum.'))
		try:
			validate(records, summary)
		except e:
			self.fail('validation failed ' + e)