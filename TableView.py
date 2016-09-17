from operator import itemgetter
from queue import Queue
from itertools import groupby, chain
from fnmatch import fnmatch
from copy import deepcopy
import re


class tableview(object):
	"""table work for type of 2dlist"""
	def __init__(self, table):
		self.field, *self.records = table # backup for original table data
	
	def _get_record_indexs(self, *names):
		if set(names) <= set(self.field):
			return [self.field.index(name) for name in names]

	def get_curView(self):
		return [self.field] + self.records

	def update_curView(self, new_table):
		self.field, *self.records = new_table 	

	def update_records(self, tgt_column, func, *ref_columns, **lookup_set):
		arg_idx = self._get_record_indexs(*ref_columns)
		tgt_idx, *_ = self._get_record_indexs(tgt_column)

		if func:
			for rec in self.records:
				rec[tgt_idx]=func(rec, *arg_idx)
		elif lookup_set:
			for rec in self.records:
				for k, v in lookup_set.items():
					if fnmatch(rec[arg_idx[0]],k):
						rec[tgt_idx] = v

	def add_column(self, *new_col_name): 
		self.field +=list(new_col_name)
		for rec in self.records:
			rec += [None]*len(new_col_name)

	def order_by(self,  column_bDesc):
		rngQ = Queue()
		rngQ.put(slice(0, len(self.records)))
		while True:
			colname, bDesc = column_bDesc.pop(0)
			col, *_ = self._get_record_indexs(colname)
			while not rngQ.empty():
				slc = rngQ.get()
				self.records[slc]=list(sorted(self.records[slc], key=itemgetter(col), reverse=bDesc))
			if not column_bDesc:
				break
			idx = 0
			for g, l in groupby(self.records, key=itemgetter(col)):
				cnt = len(list(l))
				rngQ.put(slice(idx, idx+cnt))
				idx += cnt
		
	def get_group(self, column,  reduc_subtotal, **kwargs):
		col, *_ = self._get_record_indexs(column) 
		ret = []
		sub_total = []
		for g, l in groupby(sorted(self.records, key=itemgetter(col)),key=itemgetter(col)):
			l = list(l)
			pad = [''] * len(self.field)
			for cname, func in kwargs.items():
				[c] = self._get_record_indexs(cname)
				if func == sum:
 					pad[c] = func(list(filter(None,(e[c] for e in l))))
				else:
 					pad[c] = func([e[c] for e in l])
			pad[col] = g
			ret.append(pad)
			if reduc_subtotal:
				l.append(pad)
				sub_total+=l
		if reduc_subtotal:
			return [self.field]+sub_total
		else:
			return [self.field]+ret

	def get_distinct(self, *columns, elemination=False):
		ret = [self.field]
		cols = self._get_record_indexs(*columns)
		for g, l in groupby(sorted(self.records, key=itemgetter(*cols)), key=itemgetter(*cols)):
			head, *body = l
			if elemination and body:
				continue
			else:
				ret.append(head)
		return ret

	def get_select(self,  And=True, *columns, **where):
		column, *_ = columns
		p = re.compile('(<)(.+)|(>)(.+)|(==)(.+)|(>=)(.+)|(<=)(.+)|(!=)(.+)|')
		if column in ['*']:
			cols = range(len(self.field))
		else:	
			cols = self._get_record_indexs(*columns)
		ret = [list(itemgetter(*cols)(self.field))]
		
		row_set = {}
		while where:
			cname, keys = where.popitem()
			if type(keys) in [list]:
				key = keys.pop()
				if keys:
					where[cname]= keys
			else:
				key=keys

			keyPat = re.compile(key)
			for i, rec in enumerate(self.records):
				col, *_ = self._get_record_indexs(cname)
				if isinstance(key, type(self.get_select)) and 1 in match_case :
					if func(rec[col]):
						row_set.setdefault(key,set()).add(i)
				elif keyPat.search(rec[col]):
					row_set.setdefault(key,set()).add(i)
				elif p.search(key):
					try:
						op, val = filter(None,p.search(key).groups())
						exp = '"{}" {} "{}"'.format(rec[col], op, val) \
						if isinstance(rec[col], str) else '{} {} {}'.format(rec[col], op, val)
						if eval(exp):
							row_set.setdefault(key,set()).add(i)		
					except:
						pass
		row_set = list(row_set.values())
		if not where:
			s0=row_set.append(set(range(len(self.records))))
		s0 = row_set.pop()
		while row_set:
			if And:
				s0&=row_set.pop()
			else:
				s0|=row_set.pop()

		s0 = sorted(s0)
		if len(s0) >1 :
			return ret + [list(itemgetter(*cols)(rec)) for rec in itemgetter(*s0)(self.records)]
		else:
			return ret + [list(itemgetter(*cols)(rec)) for rec in [itemgetter(*s0)(self.records)]]

	def test(self, **test):
		print(test)	



tbl = [['불출일자', '병동', '환자번호', '환자명', '약품명', '처방량(규격단위)', '잔량(규격단위)', '잔량(함량단위)'], ['2016-06-02', '73', '0016002796', '이승호', '구연산펜타닐 주 100mcg/2ml', '0.5', 0.5, 1.0], ['2016-06-03', '72', '0016002796', '이승호', '구연산펜타닐 주 100mcg/2ml', '0.5', 0.5, 1.0], ['2016-06-02', '71', '0016002796', '이승호', '구연산펜타닐 주 100mcg/2ml', '0.5', 0.5, 1.0]]


tv = tableview(tbl)

def f(rec, *ref_cols):
	병동, *_ = ref_cols
	if itemgetter(*ref_cols)(rec) == '73':
		return int(rec[병동])*5

tv.add_column('몸무게')
tv.update_records('몸무게',f,'병동')

				

