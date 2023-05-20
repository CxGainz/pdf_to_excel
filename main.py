import PyPDF2
import re

"""
Should refactor for readability/efficiency and modularize. First get it to work.
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
    for page in extracted_text:
        # for the information not preceded by a title, want to split them up somehow, join function?
        delims = "\n|          Connect | Email:| Phone:|Contact:| INFORMATION|      Users |  Midwest User Group   QAD \
                 |Regular Members|MEMBER DIRECTORY|  Midwest User Group   QAD |Products & Services| E-Mail:"

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
            while index < (limit - 12):
                user_count = 0
                last_user_data = 11
                while user_count <= last_user_data:
                    temp_list.append(page_list[index + user_count])
                    if page_list[index + user_count + 1] == 'CANADA':
                        canada_flag = True
                        last_user_data = 12
                    user_count += 1

                regular_members.append(temp_list[:])
                temp_list.clear()

                if canada_flag:
                    index += 13
                    canada_flag = False
                else:
                    index += 12
        # if the member is an associate member
        else:
            # index 9 to the end of page contains the data we want in the pdf
            print(page_list[9:])

    print(associate_members)
    print(regular_members)
