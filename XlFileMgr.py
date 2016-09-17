import os, sys, csv
import xlrd, xlwt


class xlReader(object):
	def __init__(self, *file):
		self.tables = {}
		for xl in self._file_filter(*file):
			wb=xlrd.open_workbook(xl)
			for ns in range(wb.nsheets):
				sht = wb.sheet_by_index(ns)
				self.tables.setdefault(xl,[]).append([sht.row_values(r) for r in range(0, sht.nrows)])

	def merge_tables(self):
		field_set = None
		records = []
		for f, tbls in self.tables.items():
			for tbl in tbls:
				header, *tup = tbl
				if field_set:
					if set(field_set) == set(header):
						records += tup
				else:
					field_set = header
					records = tup
		return [field_set] + records 

	def __call__(self):
		return self.merge_tables()

	def _file_filter(self, *file):
		for f in file:
			fn, ext = os.path.splitext(f)
			if ext in ['.xls','xlsx']:
				yield f




class xlWriter(object):
	"""docstring for xlWriter"""
	def __init__(self, sheet_names):
		self.wb = xlwt.Workbook(encoding='utf-8')
		self.shts = {nsht: self.wb.add_sheet(nsht) for nsht in sheet_names}
		self.styles, self.data_format = {}, {}

	def set_col_width(self, sht_name, list_col_width):
		sht = self.shts[sht_name]
		for c, width in enumerate(list_col_width):
			sht.col(c).width = width

	def merge_area(self, sht_name, top=0, left=0, bottom=0, right=0, value='', style=''):
		sht = self.shts[sht_name]
		sht.write_merge(top, left, bottom, right-1, value, style)

	def register_style(self, style_name, style_sz):
		self.styles[style_name] = xlwt.easyxf(style_sz)

	def register_number_format(self, style_name, num_format):
		style = xlwt.XFStyle()
		style.num_format_str = num_format
		self.data_format[style_name] = style
		

def tocsv(_2dlist, retFile, isrun = True, always_new=True):
		if os.path.exists(retFile) and not always_new:
			fn, ext = os.path.splitext(retFile)
			mat = re.match('(.*)\((\d+)\)', fn)
			if mat:
				f, n = mat.groups()
				fn = '{}({})'.format(f, int(n)+1)
			else:
				fn = '{}({})'.format(fn, 1)
			fname = fn + ext
			tocsv(fname, isrun)
		else:
			with open(retFile,'wt', newline ='') as fp:
				wtr = csv.writer(fp)
				for row in _2dlist:
					wtr.writerow(row)
			if isrun:
				os.startfile(retFile)