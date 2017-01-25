#!/usr/bin/python
# -*- coding: utf-8 -*-
#
# Name:     Spreadsheet.py
# Version:  1.0.0
# Author:   Glenn Abastillas
# Date:     August 28, 2016

from Row import Row
from Column import Column

class Spreadsheet(object):

	def __init__(self, **kwargs):

		attributes = {'cols', 'rows', 'save', 'file', 'sep'}
		self.as_cols = True

		self.attributes = dict([(attribute, kwargs[attribute]) if attribute in kwargs else (attribute, None) for attribute in attributes])
		self.spreadsheet = list()

		if self.attributes['sep'] is None:
			self.attributes['sep']="\t"

		if self.attributes['file'] is not None:
			self.open(self.attributes['file'], self.attributes['sep'])

		if (self.attributes['rows']>0 and self.attributes['cols'] is None) or (self.attributes['cols']>0 and self.attributes['rows'] is None):
			self.attributes['cols']=2

		print self.attributes

	def __str__(self):
		self.to_rows()
		return "\n".join(str(c) for c in self.spreadsheet)

	def add_col(self, name="blank_column", fill=""):

		prior_as_cols_state = self.as_cols
		self.to_cols()

		self.spreadsheet.append(Column(name,self.attributes['rows'],fill))

		self.attributes['cols'] += 1

		if prior_as_cols_state != self.as_cols:
			self.to_rows()

	def add_row(self, fill=""):

		prior_as_cols_state = self.as_cols
		self.to_rows()
		
		self.spreadsheet.append(Row(self.attributes['cols'],fill,sep=self.attributes['sep']))

		self.attributes['rows'] += 1

		if prior_as_cols_state != self.as_cols:
			self.to_cols()

	def add_to_col(self, index, value):
		prior_as_cols_state = self.as_cols

		self.to_cols()
		self.spreadsheet[index].add(value)
		self.refresh()

		if prior_as_cols_state != self.as_cols:
			self.to_rows()
	
	def add_to_row(self, index, value):

		prior_as_cols_state = self.as_cols

		self.to_rows()
		self.spreadsheet[index].add(value)

		max_len = len(max(self.spreadsheet, key=len))

		for row in self.spreadsheet:

			if len(row) < max_len:
				difference = max_len - len(row)

				for extra_cell in xrange(difference):
					row.add("")

		self.attributes['cols'] += 1

		if prior_as_cols_state != self.as_cols:
			self.to_cols()

	def open(self, file_path, sep="\t"):
		with open(file_path, 'r') as file_in:
			text = [line.split(sep) for line in file_in.read().splitlines() if len(line)>0]

		self.attributes['cols'] = len(max(text, key=len))
		self.attributes['rows'] = len(text)

		self.spreadsheet = [Column(name=text[0][c], size=self.attributes['rows']) for c in xrange(self.attributes['cols'])]

		for i, row in enumerate(text):
			for j, attribute in enumerate(row):
				self.spreadsheet[j].set(i-1, attribute)

	def refresh(self):
		max_len = len(max(self.spreadsheet, key=len))

		if self.as_cols:
			for col in self.spreadsheet:

				if len(col) < max_len:
					difference = max_len - len(col)

					for extra_cell in xrange(difference):
						col.add("")

			self.attributes['rows'] += 1


	def to_cols(self):
		if not self.as_cols:
			transposed_spreadsheet = [Column(name=str(self.spreadsheet[0][c]), size=self.attributes['rows']) for c in xrange(self.attributes['cols'])]

			for i,r in enumerate(self.spreadsheet[1:]):

				for j,c in enumerate(r):

					transposed_spreadsheet[j].set(i,c)

			self.spreadsheet = transposed_spreadsheet
			self.as_cols = not self.as_cols

	def to_rows(self):
		if self.as_cols:
			transposed_spreadsheet = [Row(self.attributes['cols']) for r in xrange(self.attributes['rows']+1)]

			for i,c in enumerate(self.spreadsheet):

				for j,r in enumerate([c.get_name()]+c[:]):

					transposed_spreadsheet[j].set(i,r)

			self.spreadsheet = transposed_spreadsheet
			self.as_cols = not self.as_cols

	def save(self, name="Spreadsheet_object", extension="txt", path="."):
		if self.attributes['save'] is not None:
			path = self.attributes['save']
		
		with open("{}/{}.{}".format(path,name,extension), 'w') as output:
			output.write(str(self))


if __name__ == '__main__':
	s = Spreadsheet(rows=1, file="../data/spreadsheet_test.txt")
	s.to_rows()
	s.add_col("TEST")
	#s.refresh()

	s.add_row("FSDF")

	#s.refresh()

	s.add_to_row(3,"BOAST")
	#s.refresh()
	s.add_to_col(2,"POOPS")
	s.refresh()
	s.to_cols()
	#s.refresh()

	print s.attributes['cols'], s.attributes['rows']

	s.save("test")
