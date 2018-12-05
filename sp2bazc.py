import csv
import sys
from tkinter import filedialog, Tk

NUMBER_OF_COLUMNS = 15
ART, NAME, MAT, TH, L, W, L0, W0, X_1, X_2, Y_1, Y_2, QTY, UNIT, NOTE = range(NUMBER_OF_COLUMNS)

def read_csv(path):
    with open(path, 'r') as f:
        r = csv.reader(f, delimiter = '\t')
        rows = []
        for row in r:
            rows.append(row)
        return rows

def read_xlsx(path):
    import openpyxl as xl

    wb = xl.load_workbook(path)
    ws = wb.active
    sp_rows = []

    def to_sp_row(row):
        sp_row = []
        for cell in row[:NUMBER_OF_COLUMNS]:
            v = cell.value
            if v == None:
                v = ""
            sp_row.append(str(v))
        return sp_row

    def is_header(row):
        return row[ART] != "" and all([s == "" for s in row[1:]])

    rows = (to_sp_row(row) for row in ws.rows)
    while rows.__next__()[ART] != "Детали":
        pass
    for row in rows:
        if is_header(row):
            break
        sp_rows.append(row)
    return sp_rows

def get_squeezed_rows(rows):
    squeezed_rows = []
    for row in rows:
        if row[ART] == "":
            for i in range(NUMBER_OF_COLUMNS):
                if row[i] != "":
                    squeezed_rows[-1][i] += (' / ' + row[i])
        else:
            squeezed_rows.append(row)
    return squeezed_rows

def write_bazis_txt(path, rows):
    with open(path, 'w', encoding = "cp1251") as f:
        f.write("List of panels for cutting\n")
        curr_mat = ""
        for row in rows:
            if row[MAT] != curr_mat:
                f.write("Material\t{0}\tSlab\n".format(row[MAT]))
                curr_mat = row[MAT]
            f.write("{0}\t{1}\t{2}\t{3}\t{4}\t{5}\n".
                format(row[ART], row[L], row[W], row[QTY], row[NAME],
                       '\t'.join(row[X_1 : QTY])))

def input_file_path():
    path = filedialog.askopenfilename(filetypes=[('xlsx files', '.xlsx')])
    if not path in ("", ()):
        return path
    raise Exception("Программа остановлена -- не выбран файл спецификации.")

def output_file_path():
    path = filedialog.asksaveasfilename(filetypes=[('text files', '.txt')])
    if not path in ("", ()):
        return path
    raise Exception("Программа остановлена: не введено имя выходного файла.") 

#------------------------------------------------------------------------------

def main():
    try:
        root = Tk()
        root.withdraw()
        infile = input_file_path()
        rows = read_xlsx(infile)
        rows = get_squeezed_rows(rows)
        print("Из спецификации {} считаны следующие строки".format(infile))
        for row in rows:
            print("{} -- {} -- {} {}".format(row[ART], row[NAME], row[QTY], row[UNIT]))
        print("Общее кол. позиций: {}".format(len(rows)))

        write_bazis_txt(output_file_path(), rows)
        root.destroy()
    except Exception as e:
        print(e)

if __name__ == '__main__':
    main()