import openpyxl as xl
from datetime import datetime


def getDateFromCell(cell, format):
    text = cell.value
    start_index = text.index('-')
    date = text[start_index + 1:]
    date = datetime.strptime(date, format)
    return date


def fillBlack(ws, date_of_entering, all_end_date):
    start_col = 6
    black_fill = xl.styles.PatternFill(start_color='00000000',
                                       end_color='00000000',
                                       fill_type='solid')
    for y in all_end_date:
        if date_of_entering >= y:
            ws.cell(row=x, column=start_col).fill = black_fill
        start_col += 1


def fillRed(ws, top_length, row):
    should_fill_red = True
    red_fill = xl.styles.PatternFill(start_color='FFFF0000',
                                     end_color='FFFF0000',
                                     fill_type='solid')
    for i in range(top_length, 4, -1):
        if ws.cell(row=row, column=i).value:
            should_fill_red = False
            break
    if should_fill_red:
        ws.cell(row=row, column=1).fill = red_fill
        ws.cell(row=row, column=2).fill = red_fill


def fill_yellow(ws, top_length, row):
    should_fill_yellow = True
    yellow_fill = xl.styles.PatternFill(start_color='FFFFC000',
                                        end_color='FFFFC000',
                                        fill_type='solid')
    for i in range(top_length, top_length - 5, -1):
        if ws.cell(row=row, column=i).value or (
                ws.cell(row=row, column=i).fill.start_color.rgb == '00000000' and not ws.cell(row=row,
                                                                                        column=i).fill.patternType is None):
            should_fill_yellow = False
            break
    if should_fill_yellow:
        ws.cell(row=row, column=1).fill = yellow_fill
        ws.cell(row=row, column=2).fill = yellow_fill


if __name__ == '__main__':
    wb = xl.load_workbook(filename="6k萌新群成员段位及课题完成情况 2019.12.9.xlsx")
    ws = wb.active

    date_format1 = '%Y.%m.%d'
    date_format2 = '%Y/%m/%d'

    all_end_date = []
    top_row = ws[1]
    top_length = len(top_row)
    print(ws['O58'].fill)
    for x in range(6, top_length + 1):
        all_end_date.append(getDateFromCell(ws.cell(row=1, column=x), date_format1))
    for x in range(2, 59):
        date_of_entering = ws['C%d' % x].value
        current_date = datetime.now()

        fillBlack(ws, date_of_entering, all_end_date)
        fill_yellow(ws, top_length, x)
        # print(ws['L54'].fill.start_color.rgb)
        if (current_date - date_of_entering).days <= 30:
            fillRed(ws, top_length, x)
    print(ws['O58'].fill)
    wb.save('Sample.xlsx')
