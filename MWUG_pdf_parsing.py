from MWUG_Excel_parsing import *
import PyPDF2
import re


"""
PDF Parsing:

PDF data is not the usual way we receive new customer entries. These scripts are incase we ever receieve them in PDF format again.
The format of the text data from the PDF documents can be very inconsistent for each customer entry. If the customer is from 
Canada or is missing certain criteria such as company email or phone number, we must account for these discrepancies as the format 
of these extracted pdf text entries will vary.

Note:
     The first couple pages of the PDF contain unimportant information. We must use delimiters to only extract the data we want. 

"""


def extract_text_from_pdf(pdf_file: str) -> [str]:
    with open(pdf_file, 'rb') as pdf:
        reader = PyPDF2.PdfReader(pdf, strict=False)
        pdf_text = []

        for page in reader.pages:
            content = page.extract_text()
            pdf_text.append(content)

        return pdf_text


if __name__ == '__main__':
    extracted_text = extract_text_from_pdf('2023_mwug_memberdirectory.pdf')
    # remove commas, whitespaces, and enters, or split. use git to recover changes
    regular_members = []
    associate_members = []
    # used to be the inner list within the list of lists
    temp_list = []
    canada_flag = False
    blank_flag = False
    for page in extracted_text:
        
        delims = "\n|          Connect | Email:| Phone:|Contact:| INFORMATION|      Users |  Midwest User Group   QAD \
                 |Regular Members|MEMBER DIRECTORY|  Midwest User Group   QAD |Products & Services| E-Mail:|Phone:"

        page_list = re.split(delims, page)
        # break once we reach this keyword which is towards the end of the document passed essential user info
        if page_list[0] == 'MWUG Members – BY QAD  VERSION       ':
            break
        # to not include the first couple pages
        if page_list[10] == '-- Table of Contents  -- ' or page_list[4] == '-- Table of Contents  -- ':
            continue

        # if the member is a regular member create their own list to be inputted into regular_member list.(list of list)
        if page_list[2] != 'Associate Members':
            index = 11
            limit = len(page_list) - 12
            # originally did two for loops, but can't change the indices in for-loops (due to canada adding extra elem)
            while index < limit:
                user_count = 0
                last_user_data = 11
                while user_count <= last_user_data:

                    temp_list.append(page_list[index + user_count])

                    if page_list[index + user_count + 1] == 'CANADA':
                        canada_flag = True
                        last_user_data = 12

                    if user_count == 7 and page_list[index + user_count] != 'CANADA':
                        if 'SYSTEM' not in page_list[index + user_count]:
                            last_user_data = 10
                            temp_list[-1] = "blank email"
                            blank_flag = True
                            temp_list.append("")

                    user_count += 1

                regular_members.append(temp_list[:])
                temp_list.clear()

                if canada_flag:
                    index += 13
                    canada_flag = False
                elif blank_flag:
                    index += 11
                    blank_flag = False
                else:
                    index += 12
        # if the member is an associate member
        else:
            # index 9 to the end of page contains the data we want in the pdf for associate members
            index = 9
            limit = len(page_list) - 10
            while index < limit:
                user_count = 0
                last_user_data = 9
                while user_count <= last_user_data:
                    temp_list.append(page_list[index + user_count])
                    if page_list[index + user_count] == 'CANADA':
                        canada_flag = True
                        last_user_data = 10

                    user_count += 1

                    if user_count == last_user_data:
                        while 'Title:' not in page_list[index + user_count]:
                            if 'Title' in page_list[index + user_count + 1]:
                                title_delim = page_list[index + user_count + 1].split(".")
                                temp_list.append(title_delim[len(title_delim) - 1])

                            user_count += 1

                associate_members.append(temp_list[:])
                temp_list.clear()

                if canada_flag:
                    canada_flag = False

                index += user_count + 1
   
    y_count = reg_member_excel(regular_members)
    associate_member_excel(associate_members, y_count+1)
