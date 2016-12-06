import xlrd
import matplotlib.pyplot as plt

workbook = xlrd.open_workbook("thaitrade.xlsx")

exportsheet = workbook.sheet_by_name("EX")
importsheet = workbook.sheet_by_name("IM")
continent = [ "Europe", "North America", "Asia", "South America", "Africa", "Oceania"]
def main():
    print("Choose graph: ", end='')
    graph = input()
    if graph == 'continent':
        print("CHOOSE UR ACTIVITY: ",end='')
        activity = input()
        print("CHOOSE YEARS: ",end='')
        year = int(input())
        con(activity, year)
    if graph == 'Asia':
        print("CHOOSE YEARS: ",end='')
        year = int(input())
        asia(year)
    if graph == 'Europe':
        print("CHOOSE YEARS: ",end='')
        year = int(input())
        europe(year)
def readline(status, year):
    """read line from file"""
    dct_import = {}
    dct_export = {}
    table_export = []
    record_export = []
    table_import = []
    record_import = []
    as_export = []
    as_import = []
    na_export = []
    na_import = []
    sa_export = []
    sa_import = []
    eu_export = []
    eu_import = []
    oc_export = []
    oc_import = []
    af_export = []
    af_import = []

    for i in continent:
        dct_import.setdefault(i, 0)
        dct_export.setdefault(i, 0)

    total_rows1 = exportsheet.nrows
    total_cols1 = exportsheet.ncols

    total_rows2 = importsheet.nrows
    total_cols2 = importsheet.ncols

    for x in range(total_rows1):
        for y in range(total_cols1):
            record_export.append(exportsheet.cell(x, y).value)
        if record_export[17] in continent:
            dct_export[record_export[17]] += record_export[year]
            table_export.append(record_export)
        if record_export[17] == 'Asia':
            as_export.append(record_export)
        elif record_export[17] == 'North America':
            na_export.append(record_export)
        elif record_export[17] == 'South America':
            sa_export.append(record_export)
        elif record_export[17] == 'Europe':
            eu_export.append(record_export)
        elif record_export[17] == 'Oceania':
            oc_export.append(record_export)
        elif record_export[17] == 'Africa':
            af_export.append(record_export)
        record_export = []
        x += 1

    for x in range(total_rows2):
        for y in range(total_cols2):
            record_import.append(importsheet.cell(x, y).value)
        if record_import[17] in continent:
            dct_import[record_import[17]] += record_import[year]
            table_import.append(record_import)
        if record_import[17] == 'Asia':
            as_import.append(record_import)
        elif record_import[17] == 'North America':
            na_import.append(record_import)
        elif record_import[17] == 'South America':
            sa_import.append(record_import)
        elif record_import[17] == 'Europe':
            eu_import.append(record_import)
        elif record_import[17] == 'Oceania':
            oc_import.append(record_import)
        elif record_import[17] == 'Africa':
            af_import.append(record_import)
        record_import = []
        x += 1
    if status == '':
        return continent, dct_export, dct_import
    elif status == 'Asia':
        return as_import, as_export
    elif status == 'Europe':
        return eu_import, eu_export
    elif status == 'North America':
        return na_import, na_export
    elif status == 'South America':
        return sa_import, sa_export
    elif status == 'Africa':
        return af_import, af_export
    elif status == 'Oceania':
        return oc_import, oc_export
###export 2556 by continent###
def con(activity, y):
    """Import-Export between Thailand and continents"""
    if y == 2013:
        year = 2
    if y == 2014:
        year = 3
    if y == 2015:
        year = 4
    if y == 2016:
        year = 6
    continent, dct_export, dct_import = readline('',year)
    if activity == 'Export':
        export_y = [dct_export[i] for i in continent ]
        cols = ['red', 'green', 'blue', 'yellow', 'pink', 'orange']
        plt.title('Export between Thailand and continents ('+ str(y)+')')
        plt.pie(export_y, labels = continent, colors = cols, startangle = 90, autopct = '%1.1f%%')
    if activity == 'Import':
        import_y = [dct_import[i] for i in continent ]
        cols = ['red', 'green', 'blue', 'yellow', 'pink', 'orange']
        plt.title('import between Thailand and continents ('+ str(y)+')')
        plt.pie(import_y, labels = continent, colors = cols, startangle = 90, autopct = '%1.1f%%')
    #autopct = '%1.1f%%' จะเป็นการโชว์ % ในกราฟ
    plt.show()

def asia(y):
    """Imports - Exports between Thailand and other countries in Asia"""
    if y == 2013:
        year = 2
    if y == 2014:
        year = 3
    if y == 2015:
        year = 4
    if y == 2016:
        year = 6
    as_import, as_export = readline('Asia', year)
    export_x = [i for i in range(0, len(as_import)*2, 2)]
    import_x = [i for i in range(1, len(as_import)*2, 2)]
    export_y = [i[2] for i in as_export]
    import_y = [i[2] for i in as_import]
    country = [i[1] for i in as_export]
    plt.bar(import_x, import_y,label = 'Import', color = 'red')
    plt.bar(export_x, export_y,label = 'Export', color = 'blue')
    plt.title('Asia Import-Export '+str(y))
    plt.xticks(import_x, country, rotation = 90)
    plt.xlabel('Country')
    plt.ylabel('Values(million USD)')
    plt.legend()
    plt.show()

def europe(y):
    """Imports - Exports between Thailand and other countries in Asia"""
    if y == 2013:
        year = 2
    if y == 2014:
        year = 3
    if y == 2015:
        year = 4
    if y == 2016:
        year = 6
    eu_import, eu_export = readline('Europe', year)
    export_x = [i for i in range(0, len(eu_import)*2, 2)]
    import_x = [i for i in range(1, len(eu_import)*2, 2)]
    export_y = [i[2] for i in eu_export]
    import_y = [i[2] for i in eu_import]
    country = [i[1] for i in eu_export]
    plt.bar(import_x, import_y,label = 'Import', color = 'red')
    plt.bar(export_x, export_y,label = 'Export', color = 'blue')
    plt.title('Europe Import-Export '+str(y))
    plt.xticks(import_x, country, rotation = 90)
    plt.xlabel('Country')
    plt.ylabel('Values(million USD)')
    plt.legend()
    plt.show()

main()

