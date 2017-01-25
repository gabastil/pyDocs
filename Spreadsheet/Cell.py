#!/usr/bin/python
# -*- coding: utf-8 -*-
#
# Name:     Cell.py
# Version:  1.0.0
# Author:   Glenn Abastillas
# Date:     August 25, 2016

class Cell(object):

	def __init__(self, value=None):
		self.value = None
		self.value_type = None

		if value is not None:
			self.set(value)

	def __add__(self, value):
		""" return sum of addition for numbers, or concatenated string for strings """
		return self.value + value

	def __div__(self, divisor):
		if self.value_type == type("a"):
			new_value = [character for character in self.value if character not in divisor]
			return "".join(new_value)
		return self.value / divisor

	def __eq__(self, value):
		""" return boolean value of comparison """
		return self.value == value

	def __len__(self):
		""" return length of string or value of number """
		if self.value_type == type(0):
			return self.value
		return len(self.value)

	def __mul__(self, multiplier):
		""" return product of multiplication, or repeated strings by multiplier """
		return self.value * multiplier

	def __sub__(self, value):
		""" return difference from subtraction, or strings minus letters in value for strings """
		if self.value_type == type("a"):
			return self.value.replace(value, "")
		return self.value - value

	def __str__(self):
		""" return string representation of value """
		return str(self.value)

	def get(self):
		""" return the value of this cell """
		return self.value

	def set(self, value=None):
		""" set the cells value and type 
			@param	value: value to set Cell to
		"""
		self.value = value
		self.value_type = type(value)

	def type(self):
		""" return the type of this cell's value """
		return self.value_type

if __name__=="__main__":
	c = Cell()
	c.set("toasty things are delicious")
	print c.type()==type("test"), c=="Abast", c+"o"#, c*3

	c.set(3)
	print str(c), type(c)