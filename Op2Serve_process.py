from TableView import tableview
from XlFileMgr import *
from itertools import groupby
import os, sys, re
from math import ceil


class Op2Serve(tableview):
	"""docstring for Op2Serve"""
	def __init__(self, table):
		super(Op2Serve, self).__init__(table)
		self.result_field = ['원처방일자','처방번호[묶음]','환자번호','환자명','약품명','총량','규격단위','처방의명','병동','GET_DEPT_NM']
		self.esstional_field = {'원처방일자', '병동', '처방번호[묶음]','환자번호', '환자명','약품명','총량','GET_DEPT_NM','처방의명','규격단위'}

		if not self.esstional_field <= set(self.field):
			print('Esstional Field is missing')
			sys.exit()

	def __call__(self):
	
		dstc = self.get_distinct('원처방일자','환자번호','약품명','처방번호[묶음]',elemination=True)
		self.update_curView(dstc)
		self.order_by(('병동',1),('GET_DEPT_NM',1),('원처방일자',-1))
		sel = self.get_select(True, *self.result_field, **{'약품명':'[^.]'})
		return sel

class Op2ServeSaveTo(xlWriter):
	def __init__(self, _2dlist, sheet_names=['sheet1','sheet2','sheet3']):
		super(Op2ServeSaveTo, self).__init__(sheet_names)
		self.table = _2dlist

	def __call__(self):
		header = self.table[0]

		dates = list(filter(None,[rec[0] for rec in self.table[1:]]))
		min_date, max_date = min(dates), max(dates)

		self.register_style("title","font:height 500;align: horiz center, vertical center")
		self.register_style("normal","borders: top thin, left thin, right thin, bottom thin",num_format_str="#")
		
		
		self.set_col_width('sheet1',[130*20,50*20,150*20,100*20,400*20,80*20,80*20,80*20])
		self.shts['sheet1'].row(0).height_mismatch = True
		self.shts['sheet1'].height = 50*20
		self.merge_area('sheet1',right=len(self.table[0]), value="{}~{}마약류 미불출 현황".format(min_date, max_date), style=self.styles['title'])
		self.shts['sheet1'].row(0).height =50*20

		for r, row in enumerate(self.table,1):
			for c, data in enumerate(row):
				self.shts['sheet1'].write(r,c,data, self.styles['normal'])

		fname = '{}~{}마약류 미불출 현황.xls'.format(min_date, max_date)
		fn, ext = os.path.splitext(fname)
		i=1
		while True:
			try:
				self.wb.save(fname)
			except:
				fname = '{}({}){}'.format(fn,i,ext)
				i+=1
				continue
			else:
				os.startfile(fname)
				break