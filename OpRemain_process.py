from TableView import tableview
from XlFileMgr import *
from itertools import groupby
import os, sys, re
from math import ceil


codeset = {
"7MZLX" : "미졸람 주 1mg/ml 5ml" ,
"7MZL1X" : "미졸람 주 1mg/ml 5ml" ,
"7PETX" : "염산페치딘 주사 1ml" ,
"7PET1X" : "염산페치딘 주사 1ml" ,
"7FTNX" : "구연산펜타닐 주 100mcg/2ml" ,
"7FTN2X" : "구연산펜타닐 주 100mcg/2ml" ,
"7FTN3X" : "구연산펜타닐 주 500mcg/10ml" ,
"7FTN10X" : "구연산펜타닐 주 500mcg/10ml",
"7MORPX" : "비씨모르핀황산염수화물주사 5mg/5ml",
"7MORX" : "비씨모르핀황산염수화물주사 10mg/1ml",
"7KTM1X" : "휴온스케타민염산염주사 250mg/5ml" 
}

unit_ptn = re.compile('\d+(?=ml)|\d+(?=mg)|\d+\.?\d*(?=g)|\d+\.?\d*(?=G)|\d+(?=mL)|\d+(?=ML)|\d+(?=MG)')
unit_ptn2 = re.compile('ml|mg|g|ML|mL|MG|G')





class OpRemain(tableview):
	"""docstring for OpRemain"""
	def __init__(self, table):
		super(OpRemain, self).__init__(table)
		self.result_field = ['불출일자','병동','환자번호','환자명','약품명','처방량(규격단위)','잔량(규격단위)','규격단위','잔량(함량단위)']
		self.esstional_field = {'불출일자', '병동', '환자번호', '환자명','약품명','처방량(규격단위)','반납량(폐기량)','규격단위'}

		if not self.esstional_field <= set(self.field):
			print('Esstional Field is missing')
			sys.exit()

	def __call__(self):
		rm_unit = '잔량(규격단위)'
		rm_vol = '잔량(함량단위)'
		def calc_rm(rec, *ref_cols):
			col, *_ = ref_cols
			try:
				amp = float(rec[col])
				return round(ceil(amp) - amp,2)
			except:
				return None
		
		def calc_vol(rec, *ref_cols):
			unit, amount = ref_cols
			try:
				ml = float(unit_ptn.findall(rec[unit])[-1])
				amp = float(rec[amount])
				return round(amp * ml,2)
			except :
				pass
		
		self.add_column(rm_unit, rm_vol)
		self.update_records(rm_unit, calc_rm, '처방량(규격단위)')
		self.update_records(rm_vol, calc_vol, '약품명','잔량(규격단위)')
		self.update_records('약품명', None, '약품코드', **codeset)
		self.order_by(('약품명',1), ('불출일자',1))
		sel = self.get_select(True, *self.result_field, **{'처방량(규격단위)':'\d+\.\d+','약품코드':'^7'})
		self.update_curView(sel)
		subtotal = self.get_group('약품명', True,**{rm_unit:len,rm_vol:sum})
		self.update_curView(subtotal)


		return self.get_curView()

class OpRemainSaveTo(xlWriter):
	def __init__(self, _2dlist, sheet_names=['sheet1','sheet2','sheet3']):
		super(OpRemainSaveTo, self).__init__(sheet_names)
		self.table = _2dlist

	def __call__(self):
		header = self.table[0]

		dates = list(filter(None,[rec[0] for rec in self.table[1:]]))
		min_date, max_date = min(dates), max(dates)

		self.register_style("title","font:height 500;align: horiz center, vertical center")
		self.register_style("normal","borders: top thin, left thin, right thin, bottom thin")
		self.register_style("ml",'borders: top thin, left thin, right thin, bottom thin', num_format_str='0.0 "ml"')
		self.register_style("mg",'borders: top thin, left thin, right thin, bottom thin', num_format_str='0.0 "mg"')
		self.register_style("g",'borders: top thin, left thin, right thin, bottom thin', num_format_str='0.0 "g"')
		
		self.set_col_width('sheet1',[130*20,50*20,150*20,100*20,400*20,80*20,80*20,80*20])
		self.shts['sheet1'].row(0).height_mismatch = True
		self.shts['sheet1'].height = 50*20
		self.merge_area('sheet1',right=len(self.table[0]), value="{}~{}마약류 잔여량 현황".format(min_date, max_date), style=self.styles['title'])
		self.shts['sheet1'].row(0).height =50*20

		for r, row in enumerate(self.table,1):
			for c, data in enumerate(row):
				if data == '':
					continue
				if c == header.index('잔량(함량단위)') and r >1:
					try:
						style = self.styles[unit_ptn2.findall(row[header.index('약품명')])[-1]]
						self.shts['sheet1'].write(r,c,data, style)
					except:
						pass
				else:
					self.shts['sheet1'].write(r,c,data, self.styles['normal'])

		grp = []
		for g, l in groupby(self.table[1:], key=lambda row:row[4]):
			l = list(l)
			grp.append(l[-1])

		for r, row in enumerate(grp,len(self.table)+3):	
			for c, data in enumerate(row):
				if data == "":
					continue
				if c == header.index('잔량(함량단위)'):
					try:
						style = self.styles[unit_ptn2.findall(row[header.index('약품명')])[-1]]
						self.shts['sheet1'].write(r,c,data, style)
					except:
						pass
				else:
					self.shts['sheet1'].write(r,c,data, self.styles['normal'])					


		fname = '{}~{}마약류 잔여량 현황.xls'.format(min_date, max_date)
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


