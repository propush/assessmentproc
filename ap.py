import argparse

import pdfplumber
import pdfquery
import xlsxwriter
from pdfquery.cache import FileCache

table_starter = ['ШКАЛА', 'ОПИСАНИЕ ШКАЛЫ', 'Процентиль']
name_column_text = 'Тест потенциала Potential in Focus'


def is_legit_number(param):
    if param is not None and param != '':
        try:
            float(param)
            return True
        except ValueError:
            return False


def concat_value(val1, val2):
    if val1 is None or val1 == '':
        return val2
    if val2 is None or val2 == '':
        return val1
    return f'{val1} {val2}'


def concat_row(prev_row, row):
    if prev_row is None:
        return row
    else:
        return [
            f'{concat_value(prev_row[0], row[0])}',
            f'{concat_value(prev_row[1], row[1])}',
            f'{concat_value(prev_row[2], row[2])}'
        ]


def row_is_empty(row):
    return row[0] == '' and row[1] == '' and row[2] == ''


class AData:
    def __init__(self, parameter: str, description: str, value: int):
        self.parameter = parameter
        self.description = description
        self.value = value


class Assessment:
    def __init__(self, name: str, a_data_list: list[AData]):
        self.name = name
        self.a_data_list = a_data_list


def row_to_a_data(row) -> AData:
    return AData(row[0], row[1], int(row[2]))


def process_pdf(file, use_caching) -> Assessment:
    print(f"Processing {file}")
    if use_caching:
        pdf = pdfquery.PDFQuery(file, parse_tree_cacher=FileCache("/tmp/"))
    else:
        pdf = pdfquery.PDFQuery(file)
    pdf.load()
    name_column = pdf.pq(f'LTTextLineHorizontal:contains("{name_column_text}")')
    name = pdf.pq(
        f'LTTextLineHorizontal:in_bbox("{float(name_column.attr("x0"))}, '
        f'{float(name_column.attr("y0")) - 22}, '
        f'{float(name_column.attr("x1"))}, '
        f'{float(name_column.attr("y1")) - 22}")')
    # print(name)
    print('Name:', name.text())

    col1 = pdf.pq(f'LTTextLineHorizontal:contains("{table_starter[0]}")')
    col2 = pdf.pq(f'LTTextLineHorizontal:contains("{table_starter[1]}")')
    col3 = pdf.pq(f'LTTextLineHorizontal:contains("{table_starter[2]}")')
    table_settings = {
        "vertical_strategy": "explicit",
        "horizontal_strategy": "text",
        "explicit_vertical_lines": [float(col1.attr("x0")) - 5,
                                    float(col2.attr("x0")) - 5,
                                    float(col3.attr("x0")),
                                    float(col3.attr("x1"))],
        "snap_y_tolerance": 3,
    }

    data = []
    with pdfplumber.open(file) as pdf:
        table = pdf.pages[0].extract_table(table_settings=table_settings)
        prev_row = None
        table_started = False
        for row in table:
            if row == table_starter:
                table_started = True
                continue
            if table_started:
                if row_is_empty(row):
                    continue
                if is_legit_number(row[2]):
                    data.append(row_to_a_data(concat_row(prev_row, row)))
                    prev_row = None
                else:
                    prev_row = concat_row(prev_row, row)
        return Assessment(name.text(), data)


def process_pdfs(files, use_caching):
    if use_caching:
        print("WARNING!!! Using caching, info might be incorrect.")
    assessments = []
    for file in files:
        assessments.append(process_pdf(file, use_caching))
    return assessments


def export_xlsx(assessments, output):
    print(f"Exporting to {output}")
    workbook = xlsxwriter.Workbook(output)
    worksheet = workbook.add_worksheet()
    bold = workbook.add_format({"bold": True})
    worksheet.set_column("A:A", 50)
    worksheet.set_column(1, len(assessments) + 1, 20)
    worksheet.write('A1', 'Итоговый балл PIF', bold)
    column = 1
    for assessment in assessments:
        worksheet.write(0, column, assessment.name, bold)
        row = 1
        for a_data in assessment.a_data_list:
            worksheet.write(row, 0, a_data.parameter)
            worksheet.write(row, column, a_data.value)
            row += 1
        column += 1
    workbook.close()


def main():
    arg_parser = argparse.ArgumentParser(
        prog='ap',
        description='Assessments PDF parser and XLSX generator.\n'
                    'Copyright (c) Sergey Poziturin'
    )
    arg_parser.add_argument('-c', '--use-caching', help='Use pdf parsing cache (for debug only!)', action='store_true')
    arg_parser.add_argument('-o', '--output', help='XLSX output file name', default='output.xlsx')
    arg_parser.add_argument('files', metavar='file', type=str, help='PDF input file name', nargs='+')
    args = arg_parser.parse_args()
    assessments = process_pdfs(args.files, args.use_caching)
    export_xlsx(assessments, args.output)


if __name__ == '__main__':
    main()
