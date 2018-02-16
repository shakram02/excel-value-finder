from argparse import ArgumentParser

from openpyxl import load_workbook
from openpyxl.cell import Cell
from openpyxl.worksheet import Worksheet

optionParser = ArgumentParser()

optionParser.add_argument("--needle", help="Smaller set of items to be searched for in haystack",
                          default="responses.xlsx")

optionParser.add_argument("--column", help="Column name contianing needle values",
                          default="B")


# optionParser.add_argument("haystack", help="Bigger file to search for items")

def extract_valid_applicants(wb_path, sheet_name, column, reject_colors):
    wb = load_workbook(filename=wb_path)

    form_responses: Worksheet = wb[sheet_name]
    name_set = set()
    names = []
    # Don't forget to add FF at the begning of the color
    out_of_scope = 0

    # Skip the first row, it's a column def.
    for item in form_responses[column][1:]:
        cell: Cell = item
        if cell.value is None:
            continue

        if cell.fill.bgColor.rgb in reject_colors:
            out_of_scope += 1
        else:
            # People in this set should've been called, lets see
            applicant_name = cell.value.strip()

            # TODO, handle duplicates, create a triple of name, phone, phone or name, mail, phone or whateva
            if applicant_name in name_set:
                print("Duplicate:", applicant_name)

            name_set.add(applicant_name)
            names.append(applicant_name)

    print("{}/{} didn't pass application phase".format(out_of_scope, form_responses.max_row))

    from pprint import pprint
    pprint(name_set)
    print("\n\n\n\n\n\n\n\n")
    pprint(names)
    print("Unique:", len(name_set), "All:", len(names))
    return names


def get_called_applicants():
    pass


def main():
    # Protons form verification, we want to make sure that we didn't forget to call anyone

    args = optionParser.parse_args()

    reject_colors = ["FF434343", "FF999999"]
    sheet_name = 'Form Responses 1'
    applicants = extract_valid_applicants(args.needle, sheet_name, args.column, reject_colors)
    print(len(applicants))


# Main
if __name__ == "__main__":
    main()
