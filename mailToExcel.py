import os
import xlwt
from tempfile import TemporaryFile


class Person(object):
    def __init__(self, vname, nname, email):
        self.vname = vname
        self.nname = nname
        self.email = email


people = []


def check_duplicates(newperson):
    for person in people:
        if person.email == newperson.email:
            return False
    return True


with open('data.txt', 'r') as file:
    data = file.read().replace('\n', '')
    data_split = data.split()
    for part in data_split:
        if part.__contains__('<'):
            email = part[1: len(part) - 2]
            if email.__contains__("@mgb.ch"):
                name = email[0:len(email) - 7].split('.')
                vname = name[0]
                nname = name[1]
                if "-MITS" not in nname:
                    newperson = Person(vname, nname, email)
                    if check_duplicates(newperson):
                        people.append(newperson)

    book = xlwt.Workbook()
    sheet = book.add_sheet('sheet')
    sheet.write(0, 0, "Vorname")
    sheet.write(0, 1, "Nachname")
    sheet.write(0, 2, "Email")
    for i, person in enumerate(people):
        sheet.write(i+1, 0, person.vname.title())
        sheet.write(i+1, 1, person.nname.title())
        sheet.write(i+1, 2, person.email)

    filename = "output.xls"
    if os.path.isfile("output.xls"):
        os.remove("output.xls")
    book.save(filename)
    book.save(TemporaryFile())
