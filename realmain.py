import xlrd
import matplotlib.pyplot as plt

workbook = xlrd.open_workbook("thaitrade.xlsx")

exportsheet = workbook.sheet_by_name("EX")
importsheet = workbook.sheet_by_name("IM")
continent = [ "Europe", "North America", "Asia", "South America", "Africa", "Oceania"]
def readline(year):
    dct_import = {}
    dct_export = {}
    for i in continent:
        dct_import.setdefault(i, 0)
        dct_export.setdefault(i, 0)

    total_rows1 = exportsheet.nrows
    total_cols1 = exportsheet.ncols

    total_rows2 = importsheet.nrows
    total_cols2 = importsheet.ncols

    table_export = list()
    record_export = list()

    table_import = list()
    record_import = list()
    for x in range(total_rows1):
        for y in range(total_cols1):
            record_export.append(exportsheet.cell(x, y).value)
        if record_export[17] in continent:
            dct_export[record_export[17]] += record_export[year]
            table_export.append(record_export)
        record_export = []
        x += 1

    for x in range(total_rows2):
        for y in range(total_cols2):
            record_import.append(importsheet.cell(x, y).value)
        if record_import[17] in continent:
            dct_import[record_import[17]] += record_import[year]
            table_import.append(record_import)
        record_import = []
        x += 1
    return continent, dct_export, dct_import
###import - export 2556 by continent###
def con1():
    continent, dct_export, dct_import = readline(2)
    export_x = [i for i in range(0, len(continent)*3, 3)]
    import_x = [i for i in range(1, len(continent)*3, 3)]
    width = 1
    export_y = [dct_export[i] for i in continent ]
    import_y = [dct_import[i] for i in continent ]

    plt.bar(import_x, import_y, width, label = 'Import', color = 'red')
    plt.bar(export_x, export_y, width, label = 'Export', color = 'blue')
    plt.title('Import-Export 2013')
    plt.xticks(import_x, continent, rotation = 90)
    plt.xlabel('Continent')
    plt.ylabel('Values(million USD)')
    plt.legend()
    plt.show()
#############################################
###import - export 2557 by continent###
def con2():
    continent, dct_export, dct_import = readline(3)
    export_x = [i for i in range(0, len(continent)*3, 3)]
    import_x = [i for i in range(1, len(continent)*3, 3)]
    width = 1
    export_y = [dct_export[i] for i in continent ]
    import_y = [dct_import[i] for i in continent ]
    plt.title('Import-Export 2014')
    plt.bar(import_x, import_y, width, label = 'Import', color = 'red')
    plt.bar(export_x, export_y, width, label = 'Export', color = 'blue')
    plt.xticks(import_x, continent, rotation = 90)
    plt.xlabel('Continent')
    plt.ylabel('Values(million USD)')
    plt.legend()
    plt.show()

#############################################
###import - export 2558 by continent###
def con3():
    continent, dct_export, dct_import = readline(4)
    export_x = [i for i in range(0, len(continent)*3, 3)]
    import_x = [i for i in range(1, len(continent)*3, 3)]
    width = 1
    export_y = [dct_export[i] for i in continent ]
    import_y = [dct_import[i] for i in continent ]

    plt.bar(import_x, import_y, width, label = 'Import', color = 'red')
    plt.bar(export_x, export_y, width, label = 'Export', color = 'blue')
    plt.title('Import-Export 2015')
    plt.xticks(import_x, continent, rotation = 90)
    plt.xlabel('Continent')
    plt.ylabel('Values(million USD)')
    plt.legend()
    plt.show()

#############################################
###import - export 2559 by continent###
def con4():
    continent, dct_export, dct_import = readline(6)
    export_x = [i for i in range(0, len(continent)*3, 3)]
    import_x = [i for i in range(1, len(continent)*3, 3)]
    width = 1
    export_y = [dct_export[i] for i in continent ]
    import_y = [dct_import[i] for i in continent ]

    plt.bar(import_x, import_y, width, label = 'Import', color = 'red')
    plt.bar(export_x, export_y, width, label = 'Export', color = 'blue')
    plt.title('Import-Export 2016')
    plt.xticks(import_x, continent, rotation = 90)
    plt.xlabel('Continent')
    plt.ylabel('Values(million USD)')
    plt.legend()
    plt.show()

#############################################

