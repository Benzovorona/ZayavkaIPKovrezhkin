import xlrd
import xlsxwriter

ostatki = []
limit = []
zayavka = []

# Читаем остатки

ostatkiBook = xlrd.open_workbook('Ostatki.xls')
ostatkiSheet = ostatkiBook.sheet_by_index(0)
for rownum in range(ostatkiSheet.nrows):
    row = ostatkiSheet.row_values(rownum)
    if row[0] != '' and (row[2] == '' or type(row[2]) == float):
        ostatki.append(row)

# Читаем лимит

limitBook = xlrd.open_workbook('Limit.xlsx')
limitSheet = limitBook.sheet_by_index(0)
for rownum in range(limitSheet.nrows):
    row = limitSheet.row_values(rownum)
    if row[0] != '' and (row[2] == '' or type(row[2]) == float):
        limit.append(row)

# Проверяем лимит - перестраиваем позиции

# Формируем список кодов товаров

limitList = []
for i in limit:
    limitList.append(i[0])

# Сверяемся со списком кодов товаров и добавляем недостающие

for i in ostatki:
    if i[0] not in limitList:
        limit.insert(ostatki.index(i), [i[0], i[1], i[2], i[3], i[4]])

# Перестраиваем лимиты в соответствии со списком остатков
for i in range(len(ostatki)):
    for j in range(len(limit)):
        if ostatki[i][0] == limit[j][0]:
            if i != j:
                old = limit.pop(limit.index(limit[j]))
                limit.insert(i, old)
        continue

# Проверяем лимит - добавляем недостающие

for i in range(len(ostatki)):
    if ostatki[i][4] != limit[i][4]:
        if type(limit[i][4]) == float:
            if limit[i][4] > ostatki[i][4]:
                ned = float(limit[i][4]) - float(ostatki[i][4])
                zayavka.append(
                    [ostatki[i][0], ostatki[i][1], ostatki[i][2], ostatki[i][3], ned, ostatki[i][4], limit[i][4]])

# Сортируем заявку по алфавиту

sortZayavka = []
for i in zayavka:
	sortZayavka.append(i[1])
	sortZayavka.sort()

for i in sortZayavka:
	for j in zayavka:
		if i == j[1]:
			old = zayavka.pop(zayavka.index(j))
			zayavka.insert(sortZayavka.index(i), old)

# Стили заявки

wb = xlsxwriter.Workbook('Zayavka.xlsx')
ws = wb.add_worksheet()

ws.set_column(0, 0, 11)
ws.set_column(1, 1, 75)
ws.set_column(2, 6, 9)

formatZayavkahead = wb.add_format()
formatZayavkahead.set_bg_color('#ffffcf')
formatZayavkahead.set_font_name('Times New Roman')
formatZayavkahead.set_bold()
formatZayavkahead.set_border()

formatZayavkafloat1 = wb.add_format()
formatZayavkafloat1.set_num_format('0.000')
formatZayavkafloat1.set_font_name('Times New Roman')
formatZayavkafloat1.set_align('left')
formatZayavkafloat1.set_bold()

formatZayavkafloat2 = wb.add_format()
formatZayavkafloat2.set_num_format('0.00')
formatZayavkafloat2.set_font_name('Times New Roman')
formatZayavkafloat2.set_align('left')
formatZayavkafloat2.set_bold()

formatZayavkaText = wb.add_format()
formatZayavkaText.set_font_name('Times New Roman')
formatZayavkaText.set_align('left')
formatZayavkaText.set_bold()
formatZayavkaText.set_text_wrap()

# Формируем заголовки заявки

ws.write(0, 0, 'Код', formatZayavkahead)
ws.write(0, 1, 'Наименование', formatZayavkahead)
ws.write(0, 2, 'Дефицит', formatZayavkahead)
ws.write(0, 3, 'Наличие', formatZayavkahead)
ws.write(0, 4, 'Лимит', formatZayavkahead)
ws.write(0, 5, 'Ед.изм.', formatZayavkahead)
ws.write(0, 6, 'Цена', formatZayavkahead)

# Формируем заявку

row = 1
for i in zayavka:
    ws.write(row, 0, i[0], formatZayavkaText)
    ws.write(row, 1, i[1], formatZayavkaText)
    ws.write(row, 2, i[4], formatZayavkaText)
    ws.write(row, 3, i[5], formatZayavkafloat1)
    ws.write(row, 4, i[6], formatZayavkafloat1)
    ws.write(row, 5, i[3], formatZayavkaText)
    ws.write(row, 6, i[2], formatZayavkafloat2)
    row += 1
wb.close()

# Стили лимита

workbook = xlsxwriter.Workbook('Limit.xlsx')
zs = workbook.add_worksheet()

zs.set_column(0, 0, 11)
zs.set_column(1, 1, 70)
zs.set_column(2, 4, 11)

formatLimit = workbook.add_format()
formatLimit.set_bold()
formatLimit.set_font_name('Times New Roman')
formatLimit.set_bg_color('#ccffcc')
formatLimit.set_border()

formatLimithead = workbook.add_format()
formatLimithead.set_bold()
formatLimithead.set_font_name('Times New Roman')
formatLimithead.set_bg_color('#ffffcf')
formatLimithead.set_border()

formatLimitText1 = workbook.add_format()
formatLimitText1.set_bold()
formatLimitText1.set_font_name('Times New Roman')
formatLimitText1.set_num_format('0.000')
formatLimitText1.set_align('left')

formatLimitText2 = workbook.add_format()
formatLimitText2.set_bold()
formatLimitText2.set_align('left')
formatLimitText2.set_font_name('Times New Roman')
formatLimitText2.set_num_format('0.00')
formatLimitText2.set_text_wrap()

# Записываем заголовки лимитов

zs.write(0, 0, 'Код', formatLimithead)
zs.write(0, 1, 'Наименование', formatLimithead)
zs.write(0, 2, 'Цена', formatLimithead)
zs.write(0, 3, 'Ед.изм.', formatLimithead)
zs.write(0, 4, 'Кол-во', formatLimithead)

# Записываем содержание массива лимитов
row = 1
for i in limit:
    if i[2] == '':
        zs.write(row, 0, i[0], formatLimit)
        zs.write(row, 1, i[1], formatLimit)
        zs.write(row, 2, i[2], formatLimit)
        zs.write(row, 3, i[3], formatLimit)
        zs.write(row, 4, i[4], formatLimit)
        row += 1
    else:
        zs.write(row, 0, i[0], formatLimitText2)
        zs.write(row, 1, i[1], formatLimitText2)
        zs.write(row, 2, i[2], formatLimitText2)
        zs.write(row, 3, i[3], formatLimitText2)
        zs.write(row, 4, i[4], formatLimitText1)
        row += 1
workbook.close()

print('Все  готово. Нажмите Enter')
input()
