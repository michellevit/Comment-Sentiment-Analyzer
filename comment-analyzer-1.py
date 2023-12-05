from openpyxl import Workbook
from openpyxl.styles import PatternFill


def main():
    create_workbook()


# Create an excel file
## Intake: Podcast number + words to filter out
def create_workbook():
    wb = Workbook()
    ws = wb.active
    ws.title = "Set Up"
    ws.column_dimensions['A'].width = 25
    ws.column_dimensions['B'].width = 50
    ws['A1'] = "Podcast Episode Number:"
    ws['A1'].fill = PatternFill(start_color="FFFF00", fill_type="solid")
    wb.save(filename='Comment Analyzer.xlsx')


# Run comment analyzer 1
if __name__ == "__main__":
    main()









