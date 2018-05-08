import xlwings as xw
import os
import time

path = input("请输入路径：")
dirs = os.listdir(path)
dirs.sort(key=int)
for i in range(len(os.listdir(path + str(1)))):
    print(str(i + 1) + '. ' + os.listdir(path + str(1))[i])

choose = int(input("请选择汇总文件：")) - 1
col = (input("请输入要复试数据的列号："))
row_1 = int(input("请输入要复试数据的第一行："))
row_2 = int(input("请输入要复试数据的最后一行："))
print("马上为你汇总 <" + os.listdir(path + str(1))[choose] + "> 的数据")

for i in range(1, 20):
    wb_path = path + str(i) + "\\" + os.listdir(path + str(i))[choose]
    wb = xw.Book(wb_path)
    wb_2 = xw.Book(
        "C:\\Users\\zhong\\Documents\\VS Code\\Python\\xxx\\vba.xlsx")
    wb_2.sheets[0].range(
        chr(ord('A') + i - 1) + str(1)).value = "第 " + str(i) + " 次"

    wb_2.sheets[0].range(
        chr(ord('A') + i - 1) + str(2) + ':' + chr(ord('A') + i - 1) +
        str(row_2 - row_1 + 2)).value = wb.sheets[0].range(
            chr(ord(col)) + str(row_1) + ':' + chr(ord(col)) + str(row_2)
        ).options(ndim=2).value

    print("第 " + str(i) + " 次 " + chr(ord(col)) + str(row_1) + ':' +
          chr(ord(col)) + str(row_2) + " 的数据已被汇总到 " + chr(ord('A') + i - 1) +
          str(2) + ':' + chr(ord('A') + i - 1) + str(row_2 - row_1 + 2))

    wb_2.save()
    app = xw.apps.active
    app.quit()
    time.sleep(2)
