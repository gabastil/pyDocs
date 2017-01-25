#!/usr/bin/python
# -*- coding: utf-8 -*-
#
# Name:     Row.py
# Version:  1.0.0
# Author:   Glenn Abastillas
# Date:     August 25, 2016

from Cell import Cell

class Row(object):
	
	def __init__(self, size=1, default="", row_index=0, sep="\t"):
		self.index = 0
		self.sep = sep

		self.row_index = row_index
		self.row = [Cell(default) for i in xrange(size)]

	def __getitem__(self, i):
		""" return item at specified index """
		#print "GETTING", i, self.row[i]
		return self.row[i]

	def __getslice__(self, i, j):
		""" return entries at specified span """
		return self.row[i:j]

	def __iter__(self):
		"""	allow for iteration over this object
			@return self
		"""
		return self

	def __len__(self):
		""" return row size """
		return len(self.row)

	def __setitem__(self, i, value):
		""" set value of specified cell """
		self.row[i] = Cell(value)

	def __setslice__(self, i, j, value):
		""" set value of specified cells """
		self.row[i:j] = [Cell(entry) for entry in value]

	def __str__(self):
		""" return string representation of row """
		return self.sep.join((str(c) for c in self.row))

	def next(self):
		""" returns the next object when object is iterated against
			@return	next item in self.matrix
		"""
		try:
			self.index += 1
			return self.row[self.index-1]
		except(IndexError, KeyError):
			self.index = 0
			raise StopIteration

	def add(self, value):
		""" append a new value to this row 
			@param	value: value to add to row
		"""
		self.row.append(Cell(value))

	def clear(self, index, default=""):
		""" reset specified cell 
			@param	index: index of specified cell
		"""
		self.row[index] = Cell(default)

	def delete(self, index):
		""" remove the specified cell from self.row
			@param	index: idnex of specified cell
		"""
		del self.row[index]

	def reset(self, default=""):
		""" reset all cells 
			@param default: default value to reset cells with
		"""
		for cell in xrange(len(self)):
			self.row[cell] = Cell(default)

	def get(self, index):
		""" return cell object at index 
			@param	index: position of cell to return
		"""
		return self.row[index]

	def set(self, index, value):
		""" set cell value at specified index """
		self.row[index].set(value)

	def set_row(self, row):
		""" set self.row attribute to new list of cells 
			@param	row: list of Cell objects
		"""
		all_cell_objects = sum(type(entry)!=type(Cell()) for entry in row)
		#print row

		if all_cell_objects == 0:
			self.row = row
		else:

			new_row = list()
			append = new_row.append

			for entry in row:

				if type(entry)!=type(Cell()):
					append(Cell(entry))
				else:
					append(entry)

			self.row = new_row

			#print new_row

	def get_row_index(self):
		""" return row_index attribute """
		return self.row_index

	def set_row_index(self, index):
		""" set row_index attribute 
			@param	index: integer to set row to 
		"""
		self.row_index = index

	def toString(self, sep="\t"):
		""" return string representation of row """
		self.sep = sep
		return str(self)

if __name__ == '__main__':
	r = Row(4, 3)
	p = Row(4, "boo")
	r.set(2,"say")

	for c in r:
		c.set("test")

	r[1] = 3

	print str(r), len(r), r[0]
	print str(p), len(p), p[0]

	r.add(2.4)
	r.clear(0)
	print str(r), r[-1].type()
	p.set(2, "surprise")
	print p.get(2), p[2], p[1]
	print p.toString()
	p.reset()

	print str(p)
	p.set_row([3, Cell(2), "a"])
	p.set_row([Cell('a'), Cell(2), Cell(2.3)])