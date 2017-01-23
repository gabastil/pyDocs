#!/usr/bin/python
# -*- coding: utf-8 -*-
#
# Name:     InfoText.py
# Version:  1.0.0
# Author:   Glenn Abastillas
# Date:     October 15, 2015
#
# Purpose: Allows the user to:
#           1.) Print headers, instructions, footers, and other text to the screen.
#
# This class does not have scripting code in place.
#
# This class is directly inherited by the following classes:
# 	- Analyze.py
#
# Updates:
# 1. [2016/02/29] - changed wording of notes in line 14 from '... class is used in the following ...' to '... class is directly inherited by the following ...'.
# - - - - - - - - - - - - -
import os

class InfoText(object):

	"""
		InfoText() class allows the user to call a title and descriptive text for the scripts reflected in the method names.
	"""

	def spreadsheetSearch(self, type = "title", version = 1.0):
		titleText    = "      ___           ___           ___           ___           ___      \n\
     /\__\         /\__\         /\  \         /\  \         /\  \     \n\
    /:/  /        /:/ _/_       /::\  \       /::\  \       /::\  \    \n\
   /:/__/        /:/ /\__\     /:/\:\  \     /:/\:\  \     /:/\ \  \   \n\
  /::\__\____   /:/ /:/ _/_   /::\~\:\  \   /:/  \:\__\   _\:\~\ \  \  \n\
 /:/\:::::\__\ /:/_/:/ /\__\ /:/\:\ \:\__\ /:/__/ \:|__| /\ \:\ \ \__\ \n\
 \/_|:|~~|~    \:\/:/ /:/  / \/_|::\/:/  / \:\  \ /:/  / \:\ \:\ \/__/ \n\
    |:|  |      \::/_/:/  /     |:|::/  /   \:\  /:/  /   \:\ \:\__\   \n\
    |:|  |       \:\/:/  /      |:|\/__/     \:\/:/  /     \:\/:/  /   \n\
    |:|  |        \::/  /       |:|  |        \::/__/       \::/  /    \n\
     \|__|         \/__/         \|__|         ~~            \/__/    "

		authorText   = "\n\nSpreadsheetSearch.py\nVersion: " + str(version) + "\nCreated By: Glenn Abastillas\nCreated On: October 14, 2015\n"
		greetingText = "Welcome to the Keyword finder. This script assumes you \nalready have a text file with excerpts extracted from \ndeidentified documents from the client. \n(e.g., documents from .../cln folder).\n"

		return titleText + '\n' + authorText + '\n' + greetingText

	def dbCheck(self, type = "title", version = 1.0):
		titleText = "      ___           ___           ___           ___           ___      \n\
     /\  \         /\  \         /\  \         /\__\         /\__\     \n\
    /::\  \       /::\  \       /::\  \       /:/  /        /:/  /     \n\
   /:/\:\  \     /:/\:\  \     /:/\:\  \     /:/__/        /:/__/      \n\
  /:/  \:\__\   /::\~\:\__\   /:/  \:\  \   /::\  \ ___   /::\__\____  \n\
 /:/__/ \:|__| /:/\:\ \:|__| /:/__/ \:\__\ /:/\:\  /\__\ /:/\:::::\__\ \n\
 \:\  \ /:/  / \:\~\:\/:/  / \:\  \  \/__/ \/__\:\/:/  / \/_|:|~~|~    \n\
  \:\  /:/  /   \:\ \::/  /   \:\  \            \::/  /     |:|  |     \n\
   \:\/:/  /     \:\/:/  /     \:\  \           /:/  /      |:|  |     \n\
    \::/__/       \::/__/       \:\__\         /:/  /       |:|  |     \n\
     ~~            ~~            \/__/         \/__/         \|__|    "

		authorText   = "\n\nSpreadsheetSearch.py\nVersion: " + str(version) + "\nCreated By: Glenn Abastillas\nCreated On: October 14, 2015\n"
		greetingText = "Welcome to the dbChecker. This script assumes you \nalready analyzed the spreadsheet output from Key Word Search \nand need to check if the entry or entries already exist in the database. \n"

		return titleText + '\n' + authorText + '\n' + greetingText

	def scan(self, type = "title", version = 1.0):
		titleText = "      ___           ___           ___           ___           ___      \n\
     /\  \         /\  \         /\  \         /\__\         /\  \     \n\
    /::\  \       /::\  \       /::\  \       /::|  |       /::\  \    \n\
   /:/\ \  \     /:/\:\  \     /:/\:\  \     /:|:|  |      /:/\:\  \   \n\
  _\:\~\ \  \   /:/  \:\  \   /::\~\:\  \   /:/|:|  |__   /::\~\:\  \  \n\
 /\ \:\ \ \__\ /:/__/ \:\__\ /:/\:\ \:\__\ /:/ |:| /\__\ /:/\:\ \:\__\ \n\
 \:\ \:\ \/__/ \:\  \  \/__/ \/__\:\/:/  / \/__|:|/:/  / \/_|::\/:/  / \n\
  \:\ \:\__\    \:\  \            \::/  /      |:/:/  /     |:|::/  /  \n\
   \:\/:/  /     \:\  \           /:/  /       |::/  /      |:|\/__/   \n\
    \::/  /       \:\__\         /:/  /        /:/  /       |:|  |     \n\
     \/__/         \/__/         \/__/         \/__/         \|__|    "

		authorText   = "\n\Scan.py\nVersion: " + str(version) + "\nCreated By: Glenn Abastillas\nCreated On: October 27, 2015\n"
		greetingText = "Welcome to the File Scanner. This script returns the \nnumber of files contained within the folder indicated \nas well as its subfolders."

		return titleText + '\n' + authorText + '\n' + greetingText


	def analyze(self, type = "title", version = 1.0):
		titleText = "      ___           ___           ___       ___           ___      \n\
     /\  \         /\__\         /\__\     /\  \         /\  \     \n\
    /::\  \       /::|  |       /:/  /     \:\  \       /::\  \    \n\
   /:/\:\  \     /:|:|  |      /:/  /       \:\  \     /:/\:\  \   \n\
  /::\~\:\  \   /:/|:|  |__   /:/  /         \:\  \   /::\~\:\  \  \n\
 /:/\:\ \:\__\ /:/ |:| /\__\ /:/__/    _______\:\__\ /:/\:\ \:\__\ \n\
 \/__\:\/:/  / \/__|:|/:/  / \:\  \    \::::::::/__/ \/_|::\/:/  / \n\
      \::/  /      |:/:/  /   \:\  \    \:\~~\~~        |:|::/  /  \n\
      /:/  /       |::/  /     \:\  \    \:\  \         |:|\/__/   \n\
     /:/  /        /:/  /       \:\__\    \:\__\        |:|  |     \n\
     \/__/         \/__/         \/__/     \/__/         \|__|  "

		authorText   = "\n\Analyzer.py\nVersion: " + str(version) + "\nCreated By: Glenn Abastillas\nCreated On: December 3, 2015\n"
		greetingText = "Welcome to the Analyzer. This script analyzes\na collection of cleaned documents using data\nfrom the types.gd file."

		return titleText + '\n' + authorText + '\n' + greetingText

	def separator(self, separatorType = "=*", length = 20):
		return "\n" + separatorType * length + "\n"

	def location(self):
		return "You are here: " + os.getcwd() + "\n"