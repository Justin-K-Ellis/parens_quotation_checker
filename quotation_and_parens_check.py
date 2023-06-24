"""A simple script that looks inside individual cells in a spreadsheet to check if
1. the opening and closing parentheses match, and
2. if the number of quotations marks match.
If there is a mismatch, the sheet and cell will be printed to the terminal."""

import sys
from openpyxl import load_workbook

sys_workbook = sys.argv[1]

def main():
    wb = load_workbook(sys_workbook)
    parens_checker(wb)
    quote_checker(wb)


def parens_checker(workbook):
    """Checks the number of left and right parens in each cell,
    prints "Parens mismatch at <sheet, cell>" if the numbers are not equal."""
    for sheet in workbook.worksheets:
        for tuplecell in sheet:
            for cell in tuplecell:
                count_l = 0
                count_r = 0
                try:
                    for char in cell.value:
                        if char == "(":
                            count_l = count_l + 1
                        if char == ")":
                            count_r = count_r + 1
                    if count_l != count_r:
                        print("Parens mismatch at: " + str(cell))
                    else:
                        continue  # Moves on from cells with no mismatches.
                except TypeError:
                    continue  # Skips over blank cells.
    print("Parens check complete.")


def quote_checker(workbook):
    """Checks if the number of quotation marks if even or odd,
    prints "Quotation mismatch at <sheet, cell>" if the number is odd."""
    for sheet in workbook.worksheets:
        for tuplecell in sheet:
            for cell in tuplecell:
                q_count = 0
                try:
                    for char in cell.value:
                        if char == "\"":
                            q_count = q_count + 1
                    if q_count % 2 != 0:
                        print("Quotation mismatch at: " + str(cell))
                    else:
                        continue  # Moves on from cells with no mismatches.
                except TypeError:
                    continue  # Skips over blank cells.
    print("Quotation check complete.")


if __name__ == "__main__":
    main()
