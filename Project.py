import xlrd

workbook = xlrd.open_workbook("thaitrade.xlsx")

worksheet = workbook.sheet_by_name("EX")
continent = ["เอเชีย", "ยุโรป", "อเมริกาเหนือ", "อเมริกาใต้", "แอฟริกา", "ออสเตรเลีย"]
dct_export = {}
for i in continent:
    dct_export.setdefault(i, 0)

total_rows = worksheet.nrows
total_cols = worksheet.ncols

table = list()
record = list()

for x in range(total_rows):
    for y in range(total_cols):
        record.append(worksheet.cell(x, y).value)
    if record[17] in continent:
        dct_export[record[17]] += record[2]
        table.append(record)
    record = []
    x += 1
##for i in range(len(table)):
##    print(table[i][1])
print(dct_export)
