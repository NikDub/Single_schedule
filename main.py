import re
import docx
from docxtpl import DocxTemplate

# Делай как нужно и будет как надо
# Do as you need and it will be as it should be
context = {}
replaceName = []
fileList = []
scheduleList = []

lectorName = input("Введите имя преподовалетя: ")

countName = int(input("Введите количество имен на замену(0+): "))
for i in range(countName):
    replaceName.append(input("Введите имя" + str(i) + " для замены(формата 'Фамилия ИО'): "))

countFile = int(input("Введите количество файлов расписания: "))
for i in range(countFile):
    fileList.append(input("Введите имя файла" + str(i) + ": ") + ".docx")

print("Ожидайте обрабатываем файлы ...")
for fileName in fileList:
    print("Обрабатка файла " + fileName + " началась ...")
    doc = docx.Document(fileName)
    table = doc.tables[1]
    previousString = ' '
    for i, row in enumerate(table.rows):
        for j, cell in enumerate(row.cells):
            if cell.text.find(lectorName) != -1:
                resultString = table.rows[i].cells[0].text[0:3] + table.rows[i].cells[1].text[0:2].replace(":", "")
                match = re.compile('w:fill=\"(\S*)\"').search(cell._tc.xml)
                if match:
                    if match.group(1) in ['FFFFFF', 'auto']:
                        resultString += 'Б'
                    else:
                        resultString += 'З'

                resultString += ' ' + table.rows[0].cells[j].text + ' '

                nameStr = cell.text.replace(".", "")
                for nameForReplace in replaceName:
                    nameStr = nameStr.replace(nameForReplace, "")
                resultString += nameStr

                resultString = resultString.replace("\n", " ").replace("   ", " ").replace("  ", " ")

                if previousString != resultString:
                    previousString = resultString

                    if table.rows[i].cells[1].text == table.rows[i + 1].cells[1].text and \
                            table.rows[i + 1].cells[j].text == cell.text or \
                            table.rows[i].cells[1].text == table.rows[i - 1].cells[1].text and \
                            table.rows[i - 1].cells[j].text == cell.text:
                        scheduleList.append(resultString.replace("Б ", "З ", 1))

                    elif table.rows[i - 1].cells[j].text != cell.text and \
                            table.rows[i + 1].cells[j].text != cell.text and \
                            table.rows[i - 1].cells[1].text != table.rows[i].cells[1].text and \
                            table.rows[i + 1].cells[1].text != table.rows[i].cells[1].text:
                        scheduleList.append(resultString.replace("Б ", "З ", 1))

                    scheduleList.append(resultString)
    print("Обрабатка файла " + fileName + " завершина!")

print("Все файлы обработаны, записываем итоговый файл ...")
doc = DocxTemplate("ШаблонРасписания.docx")
for i in scheduleList:
    context[i[0:6].strip()] = i[6:].strip()
    print(i)
doc.render(context)
doc.save("Расписание.docx")

stop = input("Итоговый файл готов, [Нажмите Enter для завершения]")