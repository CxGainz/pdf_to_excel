"""
Excel parsing
For when the title element didn't parse correctly, look where a word has a Capital letter in the middle of it,(
if capital letter TRUE in middle of word, then split from there, that is where the title begins.

check for when company website is missing

Goals:
1) correct any parsing troubles (Canadian entries mostly) here (unless have to be fixed in pdf_parsing.py)
2) insert the data accordingly into preexisting excel worksheet in alphabetical order (optional- excel builtins)
3) highlight newly added row in color pertaining to subscription status (done)

"""
import xlwings as xw
import re

wb = xw.Book("mwug_test_file.xlsx")
sheet = wb.sheets['Sheet1']


def reg_member_excel(reg_members):
    y_count = 1
    for i in reg_members:
        canada_flag = False
        x_count = 0
        x = None
        y = str(1 + y_count)

        # maybe can add elif if country not canada, just add it then pop.
        if i[7] == 'CANADA':
            country = 'CANADA'
            i.pop(7)
            canada_flag = True
        else:
            country = 'United States'

        sheet['AI' + y].value = country

        for j in i:

            if x_count == 0:
                x = 'A'
            elif x_count == 1:
                x = 'N'
            elif x_count == 2:
                x = 'J'
            elif x_count == 3:
                x = 'I'
            elif x_count == 4:
                if canada_flag:
                    canada_flag = False
                    sheet['Q' + y].value = j[-7:]
                    sheet['P' + y].value = j[-9:-7]
                    sheet['O' + y].value = j[:-9]
                else:
                    # [start,end,step]
                    reverse_str = j[::-1]
                    space_index = reverse_str.find(" ")
                    last_word = j[-space_index - 1:]
                    city = j[:-space_index - 1]
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
                x = 'AJ'
                j = j.replace("SYSTEM", "")
            elif x_count == 8:
                x_count += 1
                continue
            elif x_count == 9:
                x = 'AB'
                j = j.replace("QAD Version:", "")
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
            lower = 'A' + y
            upper = ':AL' + y
            comb = lower + upper
            sheet.range(comb).color = (173, 216, 230)

        y_count += 1
    return y_count


def associate_member_excel(associate_members, y):
    for i in associate_members:
        canada_flag = False
        y_coord = str(y)
        if i[2] == ')':
            i.pop(2)
        if i[8] == 'CANADA':
            country = 'CANADA'
            i.pop(8)
            canada_flag = True
        else:
            country = 'United States'

        sheet['AI' + y_coord].value = country

        for j in range(len(i)):
            if j == 0:
                x = 'A'
            elif j == 1:
                i[j].replace('(Box', "")
                x = 'N'
            elif j == 2:
                x = 'I'
            elif j == 3:
                if canada_flag:
                    canada_flag = False
                    sheet['Q' + y_coord].value = i[j][-7:]
                    sheet['P' + y_coord].value = i[j][-9:-7]
                    sheet['O' + y_coord].value = i[j][:-9]
                else:
                    # [start,end,step]
                    reverse_str = i[j][::-1]
                    space_index = reverse_str.find(" ")
                    last_word = i[j][-space_index - 1:]
                    city = i[j][:-space_index - 1]
                    if last_word[1].islower():
                        city = city + last_word[1]
                        last_word = last_word[2:]
                        last_word = " " + last_word
                    state = last_word[:3]
                    if city[-3] == " ":
                        if city[-2:].isalpha():
                            state = city[-2:]
                            city = city[:-3]

                    sheet['O' + y_coord].value = city

                    if last_word[1:2].isalpha():
                        sheet['Q' + y_coord].value = last_word[3:]
                    else:
                        sheet['Q' + y_coord].value = last_word

                    sheet['P' + y_coord].value = state
                continue

            elif j == 4 or j == 5:
                continue
            elif j == 6:
                x = 'J'
            elif j == 7:
                temp = i[7].split()
                sheet['D' + y_coord].value = temp[0]
                sheet['E' + y_coord].value = temp[1]
                continue
            elif j == 8:
                x = 'AJ'
            elif j == 9:
                temp = re.split("Title:", i[j])
                x = 'F'
                cell = x + y_coord
                sheet[cell].value = temp
                continue
            else:
                x = None

            if x is not None:
                cell = x + y_coord
                sheet[cell].value = i[j]
        y += 1
