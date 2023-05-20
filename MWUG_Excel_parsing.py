"""
Excel parsing
For when the title element didn't parse correctly, look where a word has a Capital letter in the middle of it,(
if capital letter TRUE in middle of word, then split from there, that is where the title begins.
"""
import xlwings as xw

wb = xw.Book('mwug_test_file.xlsx')

sheet = wb.sheets['Sheet1']
sheet['A1'].value = 'Foo1'

