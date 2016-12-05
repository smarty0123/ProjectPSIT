import xlrd
import matplotlib.pyplot as plt

workbook = xlrd.open_workbook("thaitrade.xlsx")

exportsheet = workbook.sheet_by_name("EX")
importsheet = workbook.sheet_by_name("IM")
continent = [ "Europe", "North America", "Asia", "South America", "Africa", "Oceania"]
def readline(status, year):
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
    elif status == 'asia':
        return as_import, as_export
###import - export 2556 by continent###
def con1():
    continent, dct_export, dct_import = readline('',2)
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
import xlrd
import matplotlib.pyplot as plt

workbook = xlrd.open_workbook("thaitrade.xlsx")

exportsheet = workbook.sheet_by_name("EX")
importsheet = workbook.sheet_by_name("IM")
continent = [ "Europe", "North America", "Asia", "South America", "Africa", "Oceania"]
def readline(status, year):
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
    elif status == 'asia':
        return as_import, as_export
###import - export 2556 by continent###
def con1():
    continent, dct_export, dct_import = readline('',2)
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
    plt.show()
#############################################
###import - export 2557 by continent###
def con2():
    continent, dct_export, dct_import = readline('',3)
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
    continent, dct_export, dct_import = readline('',4)
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
    continent, dct_export, dct_import = readline('',6)
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
###ASIA 2013###
def asia_2013():
    as_import, as_export = readline('asia',2)
    export_x = [i for i in range(0, len(as_import)*2, 2)]
    import_x = [i for i in range(1, len(as_import)*2, 2)]
    export_y = [i[2] for i in as_export]
    import_y = [i[2] for i in as_import]
    country = [i[1] for i in as_export]
    plt.bar(import_x, import_y,label = 'Import', color = 'red')
    plt.bar(export_x, export_y,label = 'Export', color = 'blue')
    plt.title('Asia Import-Export 2013')
    plt.xticks(import_x, country, rotation = 90)
    plt.xlabel('Country')
    plt.ylabel('Values(million USD)')
    plt.show()
#################################
###ASIA 2014###
def asia_2014():
    as_import, as_export = readline('asia',3)
    export_x = [i for i in range(0, len(as_import)*2, 2)]
    import_x = [i for i in range(1, len(as_import)*2, 2)]
    export_y = [i[2] for i in as_export]
    import_y = [i[2] for i in as_import]
    country = [i[1] for i in as_export]
    plt.bar(import_x, import_y,label = 'Import', color = 'red')
    plt.bar(export_x, export_y,label = 'Export', color = 'blue')
    plt.title('Asia Import-Export 2014')
    plt.xticks(import_x, country, rotation = 90)
    plt.xlabel('Country')
    plt.ylabel('Values(million USD)')
    plt.show()
#################################
#####
