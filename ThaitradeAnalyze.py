import xlrd
import matplotlib.pyplot as plt

workbook = xlrd.open_workbook("thaitrade.xlsx")

exportsheet = workbook.sheet_by_name("EX")
importsheet = workbook.sheet_by_name("IM")
groupsheet = workbook.sheet_by_name("Group")
continent = ['EUROPE', 'NORTH AMERICA', 'ASIA', 'SOUTH AMERICA', 'AFRICA', 'OCEANIA']
def main():
    '''Interface'''
    while True:
        print("SELECT GRAPH")
        print("(Continent/Group/Europe/North America/Asia/South America/Africa/Oceania):")
        print('=>', end='')
        graph = input().strip().upper()
        if graph not in ['CONTINENT', 'GROUP', 'EUROPE', 'NORTH AMERICA', 'ASIA', 'SOUTH AMERICA', 'AFRICA', 'OCEANIA']:
            print("Invalid Message!!Please Try Again.")
        if graph in ['CONTINENT', 'GROUP', 'EUROPE', 'NORTH AMERICA', 'ASIA', 'SOUTH AMERICA', 'AFRICA', 'OCEANIA']:
            if graph == 'CONTINENT':
                while True:
                    print("SELECT ACTIVITY(Export/Import): ",end='')
                    activity = input().strip().upper()
                    if activity in ['EXPORT', 'IMPORT']:
                        while True:
                            print("SELECT YEAR(2013-2016): ",end='')
                            year = int(input())
                            if year >= 2013 and year <= 2016:
                                con(activity, year)
                                break
                            else:
                                print("Invalid year!!Please Try Again.")
                        break
                    else:
                        print("Invalid Message!!Please Try Again.")
            elif graph == 'GROUP':
                while True:
                    print("SELECT YEAR(2013-2016): ",end='')
                    year = int(input())
                    if year >= 2013 and year <= 2016:
                        group(graph, year)
                        break
                    else:
                        print("Invalid year!!Please Try Again.")
            elif graph in continent:
                while True:
                    print("SELECT YEAR(2013-2016): ",end='')
                    year = int(input())
                    if year >= 2013 and year <= 2016:
                        country(graph, year)
                        break
                    else:
                        print("Invalid year!!Please Try Again.")
            break
def readline(status, year):
    """read line from excel file"""
    group = ['WORLDWIDE', 'APEC', 'RCEP', 'TPP', 'ASEAN', 'EU', 'BIMSTEC', 'EFTA']
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
    group_import = []
    group_export = []
    record_group = []
    for i in continent:
        dct_import.setdefault(i, 0)
        dct_export.setdefault(i, 0)

    total_rows1 = exportsheet.nrows
    total_cols1 = exportsheet.ncols

    total_rows2 = importsheet.nrows
    total_cols2 = importsheet.ncols

    total_rows3 = groupsheet.nrows
    total_cols3 = groupsheet.ncols

    for x in range(total_rows1):
        for y in range(total_cols1):
            record_export.append(exportsheet.cell(x, y).value)
        if record_export[17].upper() in continent:
            dct_export[record_export[17].upper()] += record_export[year]
            table_export.append(record_export)
        if record_export[17].upper() == 'ASIA':
            as_export.append(record_export)
        elif record_export[17].upper() == 'NORTH AMERICA':
            na_export.append(record_export)
        elif record_export[17].upper() == 'SOUTH AMERICA':
            sa_export.append(record_export)
        elif record_export[17].upper() == 'EUROPE':
            eu_export.append(record_export)
        elif record_export[17].upper() == 'OCEANIA':
            oc_export.append(record_export)
        elif record_export[17].upper() == 'AFRICA':
            af_export.append(record_export)
        record_export = []
        x += 1

    for x in range(total_rows2):
        for y in range(total_cols2):
            record_import.append(importsheet.cell(x, y).value)
        if record_import[17].upper() in continent:
            dct_import[record_import[17].upper()] += record_import[year]
            table_import.append(record_import)
        if record_import[17].upper() == 'ASIA':
            as_import.append(record_import)
        elif record_import[17].upper() == 'NORTH AMERICA':
            na_import.append(record_import)
        elif record_import[17].upper() == 'SOUTH AMERICA':
            sa_import.append(record_import)
        elif record_import[17].upper() == 'EUROPE':
            eu_import.append(record_import)
        elif record_import[17].upper() == 'OCEANIA':
            oc_import.append(record_import)
        elif record_import[17].upper() == 'AFRICA':
            af_import.append(record_import)
        record_import = []
        x += 1
    for x in range(total_rows3):
        for y in range(total_cols3):
            record_group.append(groupsheet.cell(x, y).value)
        if record_group[0] in group:
            group_export.append(record_group[year])
            group_import.append(record_group[year+5])
        record_group = []
        x += 1
    if status == '':
        return continent, dct_export, dct_import
    elif status == 'ASIA':
        return as_import, as_export
    elif status == 'EUROPE':
        return eu_import, eu_export
    elif status == 'NORTH AMERICA':
        return na_import, na_export
    elif status == 'SOUTH AMERICA':
        return sa_import, sa_export
    elif status == 'AFRICA':
        return af_import, af_export
    elif status == 'OCEANIA':
        return oc_import, oc_export
    elif status == 'GROUP':
        return group_import, group_export

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
    if activity == 'EXPORT':
        export_y = [dct_export[i] for i in continent ]
        cols = ['red', 'green', 'blue', 'yellow', 'pink', 'orange']
        plt.title('Export between Thailand and continents ('+ str(y)+')')
        plt.pie(export_y, labels = continent, colors = cols, startangle = 90, autopct = '%1.1f%%')
    if activity == 'IMPORT':
        import_y = [dct_import[i] for i in continent ]
        cols = ['red', 'green', 'blue', 'yellow', 'pink', 'orange']
        plt.title('Import between Thailand and continents ('+ str(y)+')')
        plt.pie(import_y, labels = continent, colors = cols, startangle = 90, autopct = '%1.1f%%')
    plt.show()

def country(c, y):
    """Imports - Exports between Thailand and other countries"""
    if y == 2013:
        year = 2
    if y == 2014:
        year = 3
    if y == 2015:
        year = 4
    if y == 2016:
        year = 6
    import_c, export_c = readline(c, year)
    export_x = [i for i in range(0, len(import_c)*2, 2)]
    import_x = [i for i in range(1, len(import_c)*2, 2)]
    export_y = [i[2] for i in export_c]
    import_y = [i[2] for i in import_c]
    country = [i[1] for i in export_c]
    plt.bar(export_x, import_y,label = 'Import', color = 'red')
    plt.bar(import_x, export_y,label = 'Export', color = 'blue')
    plt.title('Imports - Exports between Thailand and other countries in '+str(c)+' '+str(y))
    plt.xticks(import_x, country, rotation = 90)
    plt.xlabel('Country')
    plt.ylabel('Values(million USD)')
    plt.legend()
    plt.show()
def group(c, y):
    """Imports - Exports between Thailand and Various Groups"""
    group = ['WORLDWIDE', 'APEC', 'RCEP', 'TPP', 'ASEAN', 'EU', 'BIMSTEC', 'EFTA']
    if y == 2013:
        year = 6
    if y == 2014:
        year = 7
    if y == 2015:
        year = 8
    if y == 2016:
        year = 10
    import_g, export_g = readline(c, year)
    export_x = [i for i in range(0, len(group)*3, 3)]
    import_x = [i for i in range(1, len(group)*3, 3)]
    export_y = export_g
    import_y = import_g
    width = 1.01
    plt.bar(export_x, import_y,width, label = 'Import', color = 'red')
    plt.bar(import_x, export_y,width, label = 'Export', color = 'blue')
    plt.title('Imports - Exports between Thailand and Various Groups '+str(y))
    plt.xticks(import_x, group, rotation = 90)
    plt.xlabel('Group')
    plt.ylabel('Values(million USD)')
    plt.legend()
    plt.show()

main()
