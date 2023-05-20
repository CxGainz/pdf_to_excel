"""
Excel parsing
For when the title element didn't parse correctly, look where a word has a Capital letter in the middle of it,(
if capital letter TRUE in middle of word, then split from there, that is where the title begins.

check for when company website is missing

Goals:
1) correct any parsing troubles (Canadian entries mostly) here (unless have to be fixed in pdf_parsing.py)
2) insert the data accordingly into preexisting excel worksheet in alphabetical order
3) highlight newly added row in color pertaining to subscription status

"""
import xlwings as xw
import re

wb = xw.Book('mwug_test_file.xlsx')
sheet = wb.sheets['Sheet1']


def reg_member_excel(reg_members):
    y_count = 1
    for i in reg_members:
        x_count = 0
        x = None

        if i[7] == 'CANADA':
            i[4] = i[4] + ' CANADA'
            i.pop(7)

        for j in i:
            y = str(1 + y_count)
            if x_count == 0:
                x = 'A'
            elif x_count == 1:
                x = 'N'
            elif x_count == 2:
                x = 'J'
            elif x_count == 3:
                x = 'I'
            elif x_count == 4:
                # [start,end,step]
                reverse_str = j[::-1]
                space_index = reverse_str.find(" ")
                last_word = j[-space_index-1:]
                city = j[:-space_index-1]
                state = last_word[:3]
                if city[-3] == " ":
                    if city[-2:].isalpha():
                        state = city[-2:]
                        city = city[:-3]

                sheet['O' + y].value = city

                if last_word[1:2].isalpha():
                    sheet['Q' + y].value = last_word[3:]
                else:
                    sheet['Q' + y].value = last_word

                sheet['P' + y].value = state

                x_count += 1
                continue

            elif x_count == 5:
                x_count += 1
                continue
            elif x_count == 6:
                delimited_list = re.split(",| ", j)
                sheet['D' + y].value = delimited_list[1]
                sheet['E' + y].value = delimited_list[2]
                delimited_list.clear()
                x_count += 1
                continue
            elif x_count == 7:
                # ask whether adding a company email column is wanted
                x_count += 1
                continue
            elif x_count == 8:
                x_count += 1
                continue
            elif x_count == 9:
                x = 'AB'
                j = j.replace("QAD Version:","")
            elif x_count == 10:
                x_count += 1
                continue
            elif x_count == 11:
                delimited_list = re.split("Users:|Industry:", j)
                if len(delimited_list) > 1:
                    sheet['T' + y].value = delimited_list[1]
                    sheet['F' + y].value = delimited_list[2]
                delimited_list.clear()
                x_count += 1
                continue

            if x is not None:
                cell = x + y
                sheet[cell].value = j

            x_count += 1

        y_count += 1
