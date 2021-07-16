# This is a sample Python script.

# Press ⇧F10 to execute it or replace it with your code.
# Press Double ⇧ to search everywhere for classes, files, tool windows, actions, and settings.
from bs4 import BeautifulSoup
import requests
import xlrd
import xlwt
from xlutils.copy import copy
style1 = xlwt.easyxf('pattern: pattern solid, fore_colour green;')


def getResult(query):
    try:
        newquery = query.replace(' ', '+')
        r = requests.get(
            "https://www.uniprot.org/uniprot/?query=" + newquery + "&fil=organism%3A%22Homo+sapiens+%28Human%29+%5B9606%5D%22&sort=score")
        bs = BeautifulSoup(r.content, "html.parser")
        protein_names, gene_names = ["not Found"] * 20, ["not Found"] * 20
        a = bs.find_all(class_="protein_names")
        b = bs.find_all(class_="gene-names")
        allmatch = -1
        for i in range(20):
            try:
                protein_names[i] = a[i].div['title']
                gene_names[i] = b[i].strong.text
                if protein_names[i].lower() == query.lower():
                    gene_names[0] = gene_names[i]
                    protein_names[0] = protein_names[i]
                    allmatch = 0
            except Exception:
                break
    except Exception as ex:
        print(ex)
        protein_names, gene_names = ["not Found"] * 3, ["not Found"] * 3
    print(query, protein_names[:3], gene_names[:3])
    return protein_names[:5], gene_names[:5],allmatch


# print(protein_names, gene_names)


file = 'test.xls'
wb = xlrd.open_workbook(filename=file, formatting_info=True)
excel = copy(wb)
for i in range(11, 13, 2):
    sheet = wb.sheet_by_index(i)
    sheetw = excel.get_sheet(i)
    for j in range(1, 500):
        try:
            message = sheet.cell_value(j, 2)
            gene_names,protein_names,  allmatch = getResult(message)
            if allmatch == 0:
                sheetw.write(j, 4, protein_names[0],style1)
            else:
                sheetw.write(j, 4, protein_names[0])
            sheetw.write(j, 5, gene_names[0])
            sheetw.write(j, 6, protein_names[1])
            sheetw.write(j, 7, gene_names[1])
            sheetw.write(j, 8, protein_names[2])
            sheetw.write(j, 9, gene_names[2])
            sheetw.write(j, 10, protein_names[3])
            sheetw.write(j, 11, gene_names[3])
            sheetw.write(j, 12, protein_names[4])
            sheetw.write(j, 13, gene_names[4])
        except Exception as ex:
            print(ex)
            break
        excel.save('test.xls')
