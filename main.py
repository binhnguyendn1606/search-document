##Requirements:
##Search for keyWord in a certain folders
##fileExtensions include ['txt', 'doc', 'docx', 'xlsx', 'pdf']
##folderAddress is input
##keyWord is input
##result includes: keyWord, fileName, pageNumber (for other file extensions) or line (for txt files)
##result error: 'No keyword found in input folder'

##If folder has sub-folders: skip

import os
import docx2txt
import xlrd 

fileExtensions = ['txt', 'doc', 'docx', 'xlsx', 'pdf', 'csv', 'ppt', 'pptx']

#Function 1: search keyWord in TXT file

def search_txt(path, keyWord):
    txtFiles = []
    for file in os.listdir(path):
        if file.endswith('txt'):
            f = open(path + '\\' + file, 'r')
            openFile = f.read()
            if keyWord in openFile:
                txtFiles.append(file)
    return(txtFiles)

#Function 2: search keyWord in WORD file

def search_doc(path, keyWord):
    docFiles = []
    for file in os.listdir(path):
        if file.endswith('doc') or file.endswith('docx'):
            openFile = docx2txt.process(path + '\\' + file, 'r')
            if keyWord in openFile:
                docFiles.append(file)
    return(docFiles)

#Function 3: read EXCEL file and search keyWord in EXCEL file

def get_excel_data(filePath):
    openFile = xlrd.open_workbook(filePath, 'rb')
    data = []
    for sheet in openFile.sheets():
        for row in range(sheet.nrows):
            #values = []
            for column in range(sheet.ncols):
                content = sheet.cell(row, column).value
                if content != xlrd.empty_cell.value:
                    #values.append(str(sheet.cell(row, column).value))
                    data.append(str(content))
                    #data.append(values)
    return(data)

def search_excel(path, keyWord):
    excelFiles = []
    for file in os.listdir(path):
        filePath = path + '\\' + file
        if file.endswith('xls') or file.endswith('xlsx'):
            for x in get_excel_data(filePath):
                if keyWord in x:
                    excelFiles.append(file)
    return(excelFiles)


#Function 4: combine all functions

def search_files(path, keyWord):
    txtFiles = search_txt(path, keyWord)
    docFiles = search_doc(path, keyWord)
    excelFiles = search_excel(path, keyWord)

    result = ''
    
    if len(txtFiles) == 0:
        result += ('No key word found in .txt file(s)\n')
    else:
        result += ("Key word '{}' is in {}\n".format(keyWord, ', '.join(txtFiles)))
    
    if len(docFiles) == 0:
        result += ('No key word found in .doc or .docx file(s)\n')
    else:
        result += ("Key word '{}' is in {}\n".format(keyWord, ', '.join(docFiles)))
    
    if len(excelFiles) == 0:
        result += ('No key word found in .xls or .xlsx file(s)\n')
    else:
        result += ("Key word '{}' is in {}\n".format(keyWord, ', '.join(excelFiles)))

    return(result)

#print result
print(search_files(input(r'Please input folder path: '), input(r'Please input key word: ')))
#print(search_files(r'C:\Users\Hi\Desktop\MOE', 'thisisnotthetrue'))


