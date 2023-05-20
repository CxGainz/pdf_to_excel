"""
Excel parsing
For when the title element didn't parse correctly, look where a word has a Capital letter in the middle of it,(
if capital letter TRUE in middle of word, then split from there, that is where the title begins.

check for when company website is missing
"""
import xlwings as xw

wb = xw.Book('mwug_test_file.xlsx')
sheet = wb.sheets['Sheet1']


def reg_member_excel(reg_members):
    y_count = 0
    for i in reg_members:
        x_count = 0
        for j in i:
            if x_count == 7:
                if 'SYSTEM' not in j:
                    if 'CANADA' not in j:
                        print(reg_members[y_count])
                        # have to deal with this in the pdf parsing file.

            x = chr(ord('A') + x_count)
            y = str(1 + y_count)
            cell = x + y
            #sheet[cell].value = j
            x_count += 1

        y_count += 1
