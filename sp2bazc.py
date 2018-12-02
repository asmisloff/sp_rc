import csv

NUMBER_OF_COLUMNS = 14
ART, NAME, MAT, TH, L, W, L0, W0, X_1, X_2, Y_1, Y_2, QTY, NOTE = range(NUMBER_OF_COLUMNS)

def read_csv(path):
    with open(path, 'r') as f:
        r = csv.reader(f, delimiter = '\t')
        rows = []
        for row in r:
            rows.append(row)
        return rows

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

#------------------------------------------------------------------------------

def main():
    rows = read_csv("test.csv")
    rows = get_squeezed_rows(rows)
    write_bazis_txt("q.txt", rows)

if __name__ == '__main__':
    main()