import argparse
import requests
import re
import xlwt
import xlrd

parser = argparse.ArgumentParser(description='Process scraper')
parser.add_argument('-g', '--graph', action='store', nargs='?', const=0, type=int,
                    help="Flag for plot Graphs or not (default - not).")
args = parser.parse_args()
args = vars(args)


def scrap_valute():
    years = range(2010, 2019)
    url = 'https://www.kursvaliut.ru/cредний-курс-валют-за-месяц-'
    regexp = r"<p><strong>Среднегодовой валютный курс за {year} для Доллар: <\/strong><\/p>\s*<p>(?P<value>\d+\,\d+)\s+RUB<\/p>"
    results = []
    for year in years:
        url_ = url+str(year)
        resp = requests.get(url_).text
        value = re.finditer(regexp.format(year=year), resp)
        try:
            res = [a.groupdict() for a in value][0]
            results.append((year, float(res['value'].replace(',', '.'))))
        except Exception as e:
            print(e)
            results.append((year, 'No parsing Data'))

    return results


def scrap_infl():
    years = range(2010, 2019)
    url = 'http://уровень-инфляции.рф/%D1%82%D0%B0%D0%B1%D0%BB%D0%B8%D1%86%D0%B0_%D0%B8%D0%BD%D1%84%D0%BB%D1%8F%D1%86%D0%B8%D0%B8.aspx'
    regexp = r"<td class=\"(?P<type>tableYear|tableSummary)\">(?P<value>\d+,?\d+)<\/td>"
    resp = requests.get('http://уровень-инфляции.рф/%D1%82%D0%B0%D0%B1%D0%BB%D0%B8%D1%86%D0%B0_%D0%B8%D0%BD%D1%84%D0%BB%D1%8F%D1%86%D0%B8%D0%B8.aspx').text
    results = re.finditer(regexp, resp)
    results = [i.groupdict() for i in results]
    results_ = [results[i:i+2] for i in range(0, len(results), 2)]
    results_ = [ (i[0]['value'], i[1]['value']) for i in results_ ][1:10]
    return results_[::-1]


def scrap_educ():

    url = 'https://www.minfin.ru/common/upload/library/2019/07/main/fedbud_year.xlsx'
    resp = requests.get(url)
    with open('tmp.xls', 'wb') as f:
        f.write(resp.content)
    workbook = xlrd.open_workbook("tmp.xls")
    sheet = workbook.sheet_by_index(0)
    for rowx in range(sheet.nrows):
        cols = sheet.row_values(rowx)
        if cols[1] == 'Образование':
            data = [[0, i] for i in cols[6:]]
            return data


def generate_table():

    columns = (('Курс Доллара, руб.', scrap_valute()), ('Уровень Инфляции, %', scrap_infl()),
               ('Государственные расходы на образование, млрд. руб.', scrap_educ()))
    wb = xlwt.Workbook(encoding='utf8')
    xl_style = xlwt.easyxf('font: name Arial, colour_index black, bold off, italic off; align: wrap on, vert center;')
    xl_style_title = xlwt.easyxf(
        'font: name Arial, colour_index black, bold on, italic off; align: wrap on, vert center;')
    ws = wb.add_sheet("Data")
    ws.write(0, 0, '', xl_style)
    for i, year in enumerate(range(2010, 2019)):
        ws.write(0, i+1, year, xl_style_title)
    for j, column in enumerate(columns):
        ws.write(j+1,0, column[0], xl_style_title)
        for i, data in enumerate(column[1]):
            ws.write(j+1, i+1, data[1])

    wb.save("results.xls")


generate_table()