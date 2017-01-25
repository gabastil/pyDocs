#!/usr/bin/python
# -*- coding: utf-8 -*-
#
# Name:     Column.py
# Version:  1.0.0
# Author:   Glenn Abastillas
# Date:     August 25, 2016

from Cell import Cell

class Column(object):
	
	def __init__(self, name="column", size=1, default="", column_index=0, sep="\n"):
		self.index = 0
		self.sep = sep
		self.column_name = name

		self.column_index = column_index
		self.column = [Cell(default) for i in xrange(size)]

	def __getitem__(self, i):
		""" return item at specified index """
		return self.column[i]

	def __getslice__(self, i, j):
		""" return entries at specified span """
		return self.column[i:j]

	def __iter__(self):
		"""	allow for iteration over this object
			@return self
		"""
		return self

	def __len__(self):
		""" return row length """
		return len(self.column)

	def __setitem__(self, i, value):
		""" set value of specified cell """
		self.column[i] = Cell(value)

	def __setslice__(self, i, j, value):
		""" set value of specified cells """
		self.column[i:j] = [Cell(entry) for entry in value]

	def __str__(self):
		""" return string representation of row """
		return self.sep.join([self.column_name]+[str(c) for c in self.column])

	def next(self):
		""" returns the next object when object is iterated against
			@return	next item in self.matrix
		"""
		try:
			self.index += 1
			return self.column[self.index-1]
		except(IndexError, KeyError):
			self.index = 0
			raise StopIteration

	def add(self, value):
		""" append a new value to this row 
			@param	value: value to add to row
		"""
		self.column.append(Cell(value))

	def clear(self):
		""" reset specified cell 
			@param	index: index of specified cell
		"""
		for cell in self.column:
			cell.set("")

	def clear_cell(self, index, default=""):
		""" reset specified cell 
			@param	index: index of specified cell
		"""
		self.column[index] = Cell(default)

	def delete(self, index):
		""" remove the specified cell from self.column
			@param	index: idnex of specified cell
		"""
		del self.column[index]

	def fill(self, value):
		""" set all cells to fill value
			@param	value: value to set cells to
		"""
		for cell in self.column:
			cell.set(value)

	def get(self, index):
		""" return cell object at index 
			@param	index: position of cell to return
		"""
		return self.column[index]

	def get_cells(self, indices):
		""" return a list of cells indicated by indices
			@param	indices: indices in the column spreadsheet of cells to return
		"""
		return [self.get(i) for i in indices]

	def get_column_index(self):
		""" return column_index attribute """
		return self.column_index

	def get_name(self):
		""" return column_name attribute """
		return self.column_name

	def set(self, index, value):
		""" set cell value at specified index """
		self.column[index].set(value)

	def set_column(self, row):
		""" set self.row attribute to new list of cells 
			@param	row: list of Cell objects
		"""
		all_cell_objects = sum(type(entry)!=type(Cell()) for entry in row)
		#print row

		if all_cell_objects == 0:
			self.row = row
		else:

			new_column = list()
			append = new_column.append

			for entry in row:

				if type(entry)!=type(Cell()):
					append(Cell(entry))
				else:
					append(entry)

			self.row = new_column

			#print new_column

	def set_column_index(self, index):
		""" set column_index attribute 
			@param	index: integer to set row to 
		"""
		self.column_index = index

	def set_name(self, name="column"):
		""" set column_name attribute 
			@param	name: new name for this column
		"""
		self.column_name = name

	def set_size(self, size):
		""" add cells to the column if size is greater than current size 
			@param	size: new size to set column to
		"""
		number_of_cells_to_add = size-len(self)
		append = self.column.append

		for cell in range(number_of_cells_to_add):
			append(Cell())

	def toString(self, sep="\n"):
		""" return string representation of row """
		self.sep = sep
		return str(self)

if __name__ == '__main__':
	r = Column("test", 4, 3)
	p = Column("test", 4, "boo")
	r.set(2,"say")

	r.fill("test")

	r[1] = 3

	print str(r), len(r), r[0]
	r.set_name("Column Name")
	print str(p), len(p), p[0], r.get_name()

	r.add(2.4)
	r.clear_cell(0)

	print r
	r.clear()
	print r

	q = Column()

	for i in range(len(q)):
		q.set(i, "BC")

	print "BEFORE", q

	q.set_size(4)

	q.fill("AC")

	print "AFTER ", q, q.get_cells([1,3])
