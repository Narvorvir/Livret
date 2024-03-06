from json import dump, load, loads
from PIL import Image
from os.path import exists

from openpyxl import load_workbook, Workbook

from datetime import datetime

from shutil import copyfile

# =========== JSON =============
def checkFileExistance(filePath):
    try:
        with open(filePath, 'r') as f:
            return True
    except FileNotFoundError as e:
        return False
    except IOError as e:
        return False

def CheckFileFormatGood():
    structureBase = { "formatsNames" : [],"formats" :{}}

    if checkFileExistance("format.json") == False:
        file =  open("format.json", 'w')
        dump(structureBase, file)
    else:
        file =  open("format.json", 'r+')
        try:
            content = loads(file.read())
        except:
            file.close()
            file =  open("format.json", 'w+')
            dump(structureBase, file)
            content = structureBase

        if content.get("formats") == None or content.get("formatsNames") == None:
            file.close()
            file =  open("format.json", 'w')
            dump(structureBase, file)

def setFormat(title, subjects, group, picture, *args):
    CheckFileFormatGood()
    try: 
        Image.open(picture)
    except:    
        picture = "pictures/reportBook.png"

    tilePart = ''.join(filter(str.isalnum, title)) .lower()
    
    with open("format.json", "r") as file:
        data = load(file)  
    try:
        data.get("formatsNames", {}).remove(tilePart)
    except: pass
    data.get("formatsNames", {}).append(tilePart)

    data.get("formats", {})[tilePart] = {"title": title, "picture" : picture, "subjects": subjects, "group" : group}
   
    with open("format.json", 'w') as file:
        dump(data, file, indent=4)

def formatFile():
    CheckFileFormatGood()
    setFormat("Styléé", ["math", "Pc"], True, "https////")

def getFormats():
    CheckFileFormatGood()
    with open("format.json", "r") as file:
        data = load(file) 
    return data["formats"]

def RemoveFormat(title_format, *args):
    if title_format != None and title_format != "":
        CheckFileFormatGood()

        title_format = ''.join(filter(str.isalnum, title_format)) .lower()
        
        with open("format.json", "r") as file:
            data = load(file)  
        try:
            data.get("formatsNames", {}).remove(title_format)
            data.get("formats", {}).pop(title_format)
        except: pass
        

    
        with open("format.json", 'w') as file:
            dump(data, file, indent=4)

def SetFormatVar(format_name, var_name, new_value):
    CheckFileFormatGood()
    
    if var_name != "title":

        
        with open("format.json", "r") as file:
            data = load(file)  

        data.get("formats", {}).get(format_name, {})[var_name] =  new_value
    
        with open("format.json", 'w') as file:
            dump(data, file, indent=4)


# ========== EXCEL ===============
def CreateFile(trimester, level_class, year, subjects, names, *args):
    wb = Workbook()
    setting = [("Trimestre", int (trimester)), ("Niveau", str(level_class)), ("Annee", str (year)), ("Matieres", str(subjects)), ("Noms", str(names))]

    i = 1
    sheet = wb.active
    sheet.title = "Parametre"

    for key, value in setting:
        sheet.cell(row=i, column=1).value = key
        sheet.cell(row=i, column=2).value = value

        i+=1
    if False == (isinstance(names, tuple) or  isinstance(names, list)):
        x = names
        names = []
        names.append(x)

    for name in names:
        i = 1
        sheet2 = wb.create_sheet(name)
        for subject in subjects:
            sheet2.cell(row=i, column=1).value = subject
            sheet2.cell(row=i, column=2).value = ""
            i += 1

    wb.save("back.xlsx")
    return wb

def SetName(file, name):
    setting = GetSetting(file)
    sheet = file.create_sheet(name)
    i = 1

    for subject in setting["Matieres"]:
        sheet.cell(row=i, column=1).value = subject
        sheet.cell(row=i, column=2).value = ""
        i += 1
    return file

def removeName(file, name):
    try:
        sheet = file[name]
        file.remove(sheet)
    except: pass

def GetSetting(file):
    try:
        return GetTableauSheet(file , "Parametre")
    except:
        return None

def GetTableauSheet(file, name_sheet):
    sheet = file[name_sheet]
    dico = {}
    for i in range(1, sheet.max_row + 1):
        list_value = []
        for j in range(2, sheet.max_column + 1):
            list_value.append(sheet.cell(row = i, column = j).value)
        # if len(list_value) == 1:
        if len(list_value) <= 1:
            try:
                list_value = list_value[0]
            except: list_value = ""
            try:
                list_value = eval(list_value)
            except: pass
        

     

            dico[sheet.cell(row = i, column = 1).value] =  list_value
    return dico

def GetAllTableauSheet(file):
    setting = GetSetting(file)
    dico_sheet = []

    if setting != None:
        for name in setting["Noms"]:
            sheet_value = GetTableauSheet(file, name)
            sheet_value["name"] = name
            dico_sheet.append(sheet_value)
        return dico_sheet

def OpenFileExcel(file):
    return load_workbook(file)

def OpenFileExcelFormat(file_name):
    if file_name.find(".xlsx") != -1:
        file = OpenFileExcel(file_name)
        if GetSetting(file) != None and GetAllTableauSheet(file) != None:
            return file

def SetSubjects(file, name, dico_subjects, file_name):
    
    i = 1

    settings = GetSetting(file)
    n = len(settings["Matieres"])

    for key , subject in dico_subjects.items():
        if i <= n:
            file[name].cell(row=i, column=2).value = subject.strip()
            i += 1
    file.save(file_name)

    return file

def SaveFile(file_name, file_base="back.xlsx"):
    copyfile(file_base, file_name)

# =============== CSV ==================
def CheckFileGood(name_file, mode="r"):
    if checkFileExistance(name_file) == False:
        file =  open(name_file, 'w')
        file.close()
           
    file =  open(name_file, mode)
    return file

def GetListFromCsv(name_file):
    list_element = []
    file = CheckFileGood(name_file)
    
    for line in file.readlines():
        if line != "":
            elmt = line.replace("\n", "", 100).split(";;")
            if len(elmt) > 0 and elmt != [""]:
                list_element.append(elmt)
    return list_element

def GetListLastFile():
    result_list = []
    values = GetListFromCsv("last_file.csv")
    if values != None or value != "":

        for value in values:   
            if len(value) == 3 and value[0].find(".xlsx") != -1 and exists(value[1] + "\\" + value[0]) and len(value[2]) == 16:
                result_list.append(value)
    
    return result_list

def SetLastFile(name_file, way_file): 
    if name_file != None and name_file != "" and name_file.find(".xlsx") != -1:

        date = datetime.today().strftime('%d/%m/%Y %H:%M')
        result = f"{name_file};;{way_file};;{date}" 


        file = CheckFileGood("last_file.csv")
        lines = file.readlines()

        while len(lines) >= 15:
            del lines[0]
        
        new_file = open("last_file.csv", "w")
        for line in lines:
            if line.find(f"{name_file};;{way_file}") == -1 and len(line) > 3:

                new_file.write(line)
        new_file.write(result + "\n")

        new_file.close()

def SplitFile(file_name):
    file_name = file_name.replace("\\", "/")
    if file_name.find("/") != -1:
        name = file_name.split("/")[-1]
        way =  file_name.split("/")
        del way[-1]
        way =  "/".join(way)
        return name, way
    else:
        return None, None