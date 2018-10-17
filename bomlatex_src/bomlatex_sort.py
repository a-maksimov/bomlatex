#!/usr/bin/env python
# -*- coding: utf-8 -*-
'''
Created on Jul 30, 2018

@author: Administrator
'''
import csv
import codecs
import re
from io import StringIO
from collections import OrderedDict

# Input and output filenames
inBomName  = 'bom.csv'
fullBomName = 'bom-full.csv'
outBomName = 'bom-gost.csv'
errorLog = 'error.csv'

# Convert bom file from Windows-1251 to UTF-8
with codecs.open(inBomName, 'r', 'cp1251') as f:
    u = f.read()   # now the contents have been transformed to a Unicode string
with codecs.open('bom-utf8.csv', 'w', 'utf-8') as out:
    out.write(u)   # and now the contents have been output as UTF-8

inBomName = 'bom-utf8.csv'

# Corresponding between BOM and Gost-BOM CSV columns
gostDesColumn  = 'LogicalDesignator'
gostCompColumn = ['ManufacturerPartNumber', 'Manufacturer']
gostCommentColumn = 'Note'
FuncGroup = 'FunctionalGroupTitle'
#
# Functions
#
# Get designator of a functional group (A1 for main\A1 etc)
# to do: hierarchy
def mkgrp(path,count):
        return path.split('\\')[count]

# Get base of designator (R for R12, C for C2, etc)
def base(designator):
        return ''.join([i for i in designator if not i.isdigit()])

# Add comma before string if not empty
def commatize(string):
        if string == '':
                return ''
        else:
                return ', ' + string

# Make component Name column
def mkdesc(rawEntry, columns):
        desc = rawEntry[columns[0]]
        for field in columns[1:]:
                desc = desc + commatize(rawEntry[field])
        return desc

# Make designator cell content for GOST bom
def mkdes(first, last, count):
        if count == 1:
                return first
        elif count == 2:
                return first + ', ' + last
        else:
                return first + '-' + last

# Make a list of functional groups
funcgroups = []
bom = []
with open(inBomName, 'r', encoding='utf-8') as bomcsv:
    bomreader = csv.DictReader(bomcsv, delimiter=',', quotechar='"')
    for row in bomreader:
        bom.append(row)
    for entry in bom:
        # Check if there is there is anything after main\ in PhysicalPath
        if len(entry['PhysicalPath'].split('\\')) > 1:
            for i in range(len(entry['PhysicalPath'].split('\\'))-1):
                funcgroup = OrderedDict({'LogicalDesignator': mkgrp(entry['PhysicalPath'],i + 1),
                                         'ManufacturerPartNumber': entry['FunctionalGroupTitle']})
            if funcgroup not in funcgroups: 
                funcgroups.append(funcgroup)
        else:
            # Remove FunctionGroupTitle from components on main
            entry['FunctionalGroupTitle'] = ''
            
# Split paths
paths = []
for entry in bom:
    path = entry['PhysicalPath'].split('\\')[0:len(entry['PhysicalPath'].split('\\'))]
    if path not in paths:
        paths.append(path)        

if paths[0][0]: # check if scheme is hierarchical
    # берем первое значение первого элемента,
    # сравниваем с каждым первым значением следующих элементов
    # если совпадает -- это избыточный элемент, выкидываем его
    # повторяем
    # получаем из 
    # paths = [['main'], ['main', 'A1'], ['main', 'A2'], ['main', 'A3'], ['main', 'A2', 'A4'], ['main', 'A3', 'A4']]
    # paths = [['A1'], ['A2', 'A4', 'A6'], ['A2', 'A4', 'A6']]
    h = []
    j = [1]
    while len(j) > 0: # check if there were any excessive paths remaining stored in j
        for num,path in enumerate(paths): # for every item in paths
            for i in range(num+1,len(paths)): # for every next item in paths
                if paths[num][0] == paths[i][0]: # if this item first element == next item first element
                    if num not in h: # if number of this path is not in list h
                        h.append(num) # append list h with number of this item
                        break
        if paths[len(paths)-1][0] == paths[len(paths)-2][0]: # if last item's first elenent == previous's
            h.append(len(paths)-1) # append h with number of last item
        for num in h: # for every number stored in h
            del paths[num][0] # delete item in paths with this number
        paths = [x for x in paths if x != []] # delete empty paths
        j = h # remember h in j for repeat
        h = [] # reset h
    
    # Assign numbers to funcgroups
    numbers = []  
    for i,path in enumerate(paths): # for every item in paths
        for k in range(len(paths[i])): # for every element of item in paths
            if k == 0: # if it is single
                number = str(i+1) # number is i etc
            else: # or if it is composite
                number = str(i+1) + '.' + str(k) # number is i+1.k etc
            numbers.append(OrderedDict({'LogicalDesignator': path[k], # assign designator = element k of number
                                        'Number': number})) # assign number
    
    # Append bom with fuctional groups
    for entry in funcgroups:
        bom.append(entry)
    
with open(fullBomName,'w', encoding='utf-8', newline='') as bomcsv:
    fieldnames = ['LogicalDesignator','ManufacturerPartNumber','Manufacturer','Note','PhysicalPath','FunctionalGroupTitle']
    writer = csv.DictWriter(bomcsv, fieldnames,extrasaction='ignore')
    writer.writeheader()
    for entry in bom:
        writer.writerow(entry)
#
# First conversion pass to temp BOM (without group headers)
#
# Define BOM-conversion structure
class Entry:
    pass

convBomEntry = Entry()          # Conversion BOM entry (made from input BOM)
tempBomEntry = Entry()          # Temp BOM entry

convBomEntry.desBase  = ''      # Designator base (R for R12, C for C2, etc)
convBomEntry.compName = ''      # Component Name
convBomEntry.compNote = ''      # Component Note

tempBomEntry.desBase   = ''     # Designator base (R for R12, C for C2, etc)
tempBomEntry.desFirst  = ''     # First designator in entry
tempBomEntry.desLast   = ''     # Last designator in entry
tempBomEntry.compDesc  = ''     # Component Name
tempBomEntry.compCount = 0      # Component count
tempBomEntry.compNote  = ''     # Component Note
tempBomEntry.FuncGroup = ''     # Component Functional Group
tempBomEntry.desList = []       # List of designators

# Temporary BOM. Fields: 'Designator', 'Name', 'Count', 'Note', 'Base', 'FunctionalGroup'
# 'Base' field needed for second conversion pass
#
tempbom = []

with open(fullBomName, 'r', encoding='utf-8') as bomcsv:
        bomreader = csv.DictReader(bomcsv, delimiter=',', quotechar='"')
        for rawEntry in bomreader:
                # Begin fill tempBomEntry if empty
                if tempBomEntry.compCount == 0:
                        tempBomEntry.desBase   = base(rawEntry[gostDesColumn])
                        tempBomEntry.desFirst  = rawEntry[gostDesColumn]
                        tempBomEntry.desLast   = rawEntry[gostDesColumn]
                        tempBomEntry.compDesc  = mkdesc(rawEntry, gostCompColumn)
                        tempBomEntry.compCount = 1
                        tempBomEntry.compComment  = rawEntry[gostCommentColumn]
                        tempBomEntry.FuncGroup  = rawEntry[FuncGroup]
                        tempBomEntry.desList.append(rawEntry[gostDesColumn])
                else:
                        # Fill conversion structure
                        convBomEntry.desBase   = base(rawEntry[gostDesColumn])
                        convBomEntry.compDesc  = mkdesc(rawEntry, gostCompColumn)
                        convBomEntry.compComment  = rawEntry[gostCommentColumn]
                        convBomEntry.FuncGroup  = rawEntry[FuncGroup]

                        # Update temp entry if data similar
                        if (convBomEntry.desBase + convBomEntry.compDesc + convBomEntry.compComment + convBomEntry.FuncGroup) == \
                           (tempBomEntry.desBase + tempBomEntry.compDesc + tempBomEntry.compComment + tempBomEntry.FuncGroup):
                                tempBomEntry.desLast = rawEntry[gostDesColumn]
                                # Check of Designator already included
                                if  tempBomEntry.desLast not in tempBomEntry.desList:
                                    tempBomEntry.compCount = tempBomEntry.compCount + 1
                                    tempBomEntry.desList.append(tempBomEntry.desLast)
                        else:
                                # Move tempBomEntry to tempbom
                                tempbom.append({'Designator': mkdes(tempBomEntry.desFirst,
                                                                    tempBomEntry.desLast,
                                                                    tempBomEntry.compCount),
                                                'Name': tempBomEntry.compDesc,
                                                'Count': tempBomEntry.compCount,
                                                'Comment': tempBomEntry.compComment,
                                                'Base': tempBomEntry.desBase,
                                                'FunctionalGroupTitle': tempBomEntry.FuncGroup})

                                # Begin fill next tempBomEntry
                                tempBomEntry.desBase   = base(rawEntry[gostDesColumn])
                                tempBomEntry.desFirst  = rawEntry[gostDesColumn]
                                tempBomEntry.desLast   = rawEntry[gostDesColumn]
                                tempBomEntry.compDesc  = mkdesc(rawEntry, gostCompColumn)
                                tempBomEntry.compCount = 1
                                tempBomEntry.compComment  = rawEntry[gostCommentColumn]
                                tempBomEntry.FuncGroup  = rawEntry[FuncGroup]
                                tempBomEntry.desList.append(tempBomEntry.desLast)

        # Move last tempBomEntry to tempbom
        tempbom.append({'Designator': mkdes(tempBomEntry.desFirst,
                                            tempBomEntry.desLast,
                                            tempBomEntry.compCount),
                        'Name': tempBomEntry.compDesc,
                        'Count': tempBomEntry.compCount,
                        'Comment': tempBomEntry.compComment,
                        'Base': tempBomEntry.desBase,
                        'FunctionalGroupTitle': tempBomEntry.FuncGroup})
#
# Second conversion pass to output BOM
#
# Read BOM headers before conversion
# cgroups fields: 'Base', 'Single', 'Multiple'
cgroups = []
        
with open('cgroups.csv', 'r', encoding='utf-8') as cgcsv:
        cgreader = csv.DictReader(cgcsv, delimiter=',', quotechar='"')
        for entry in cgreader:
                cgroups.append(entry)
#                
# Output BOM generation data
#
outbom      = []    # Output BOM
lastgroup   = []    # Last component group
nodesgroup  = []    # components with no designator
funcgroups  = []    # functional groups
funcgrouped  = []   # function grouped components
singletitle = ''    # Last group titles
multititle  = ''
#
# Second pass functions
#
def getsingle(base, cgroups):
        for entry in cgroups:
                if entry['Base'] == base:
                        return entry['Single']
        return 'Неизвестный компонент'

def getmulti(base, cgroups):
        for entry in cgroups:
                if entry['Base'] == base:
                        return entry['Multiple']
        return 'Неизвестные компоненты'

def outputgroup(outbom, lastgroup):
        if len(lastgroup) == 1:
                outbom.append({'Designator': '',
                                'Name': '',
                                'Count': '',
                                'Comment': ''})
                outbom.append({'Designator': lastgroup[0]['Designator'],
                                'Name': singletitle + ' ' + lastgroup[0]['Name'],
                                'Count': lastgroup[0]['Count'],
                                'Comment': lastgroup[0]['Comment']})
        else:
                outbom.append({'Designator': '',
                                'Name': '',
                                'Count': '',
                                'Comment': ''})
                outbom.append({'Designator': '',
                                'Name': multititle,
                                'Count': '',
                                'Comment': ''})
                for row in lastgroup:
                        outbom.append({'Designator': row['Designator'],
                                        'Name': row['Name'],
                                        'Count': row['Count'],
                                        'Comment': row['Comment']})
#                       
# Output BOM generation
#
for entry in tempbom:
#        First iteration
        if entry['Designator'] == '':
            nodesgroup.append(entry['Name'])
            continue
        if entry['Base'] == 'A':
            funcgroups.append({'Designator': entry['Designator'],
                                  'Name': entry['Name'],
                                  'Count': entry['Count'],
                                  'Comment': entry['Comment']})
            continue
        if entry['FunctionalGroupTitle'] != '':
            funcgrouped.append({'Designator': entry['Designator'],
                                  'Name': entry['Name'],
                                  'Count': entry['Count'],
                                  'Comment': entry['Comment'],
                                  'Base': entry['Base'],
                                  'FunctionalGroupTitle': entry['FunctionalGroupTitle']})
            continue
        if multititle == '':
                lastgroup.append({'Designator': entry['Designator'],
                                  'Name': entry['Name'],
                                  'Count': entry['Count'],
                                  'Comment': entry['Comment']})
                singletitle = getsingle(entry['Base'], cgroups)
                multititle =  getmulti(entry['Base'], cgroups)
        # Extend group
        elif getmulti(entry['Base'], cgroups) == multititle:
                lastgroup.append({'Designator': entry['Designator'],
                                  'Name': entry['Name'],
                                  'Count': entry['Count'],
                                  'Comment': entry['Comment']})
        # Output group, make new group
        else:
                # Output group
                outputgroup(outbom, lastgroup)

                # Flush and fill group
                lastgroup = []
                lastgroup.append({'Designator': entry['Designator'],
                                  'Name': entry['Name'],
                                  'Count': entry['Count'],
                                  'Comment': entry['Comment']})
                singletitle = getsingle(entry['Base'], cgroups)
                multititle =  getmulti(entry['Base'], cgroups)

# Output group
outputgroup(outbom, lastgroup)

# Write error output
with open(errorLog, 'w', encoding='utf-8', newline='') as errorlog:
    errorlog.write("Components without Designators:\n")
    for item in nodesgroup:
        errorlog.write("%s\n" % item)


# Append numbers to funcgroups
for funcgroup in funcgroups:
    for number in numbers:
        if len(funcgroup['Designator'].split(',')) > 1:
            if funcgroup['Designator'].split(',')[0] == number['LogicalDesignator']:
                funcgroup['Number'] = number['Number']
                break
        elif len(funcgroup['Designator'].split('-')) > 1:
            if funcgroup['Designator'].split('-')[0] == number['LogicalDesignator']:
                funcgroup['Number'] = number['Number']
                break   
        else:
            if funcgroup['Designator'] == number['LogicalDesignator']:
                funcgroup['Number'] = number['Number']
                break

# Append outbom with functional groups
for funcgroup in funcgroups:
    lastgroup  = []
    multititle = ''
    singletile = ''
    outbom.append({'Designator': '',
                    'Name': '',
                    'Count': '',
                    'Comment': ''})
    outbom.append({'Designator': funcgroup['Designator'],
                    'Name': funcgroup['Number'] + ' ' + funcgroup['Name'],
                    'Count': funcgroup['Count'],
                    'Comment': funcgroup['Comment']})
    for entry in funcgrouped:
        if entry['FunctionalGroupTitle'] == funcgroup['Name']:
            if multititle == '':
                lastgroup.append({'Designator': entry['Designator'],
                                  'Name': entry['Name'],
                                  'Count': entry['Count'],
                                  'Comment': entry['Comment']})
                singletitle = getsingle(entry['Base'], cgroups)
                multititle =  getmulti(entry['Base'], cgroups)
            # Extend group
            elif getmulti(entry['Base'], cgroups) == multititle:
                lastgroup.append({'Designator': entry['Designator'],
                                  'Name': entry['Name'],
                                  'Count': entry['Count'],
                                  'Comment': entry['Comment']})
            # Output group, make new group
            else:
                # Output group
                outputgroup(outbom, lastgroup)

                # Flush and fill group
                lastgroup = []
                lastgroup.append({'Designator': entry['Designator'],
                                  'Name': entry['Name'],
                                  'Count': entry['Count'],
                                  'Comment': entry['Comment']})
                singletitle = getsingle(entry['Base'], cgroups)
                multititle =  getmulti(entry['Base'], cgroups)
    outputgroup(outbom, lastgroup)

# Print result
for entry in outbom:
        print(entry['Designator'], '\t',
              entry['Name'], '\t',
              entry['Count'], '\t',
              entry['Comment'])

# Write output BOM
with open(outBomName, 'w', encoding='utf-8', newline='') as gostbomcsv:
    fieldnames = ['Designator', 'Name', 'Count', 'Comment']
    writer = csv.DictWriter(gostbomcsv, fieldnames=fieldnames)
    writer.writeheader()
    for entry in outbom:
            if entry['Comment'] == 'Не уст.': # Нет в ГОСТ 2.0.12
                entry['Comment'] = 'Не устанавливать'
            writer.writerow({'Designator': entry['Designator'],
                             'Name': entry['Name'],
                             'Count': entry['Count'],
                             'Comment': entry['Comment']})
#
# Parse for LaTeX package csvsimple
#
latexBomName = 'latex\\bom-gost-latex.csv'
#
# Center and underline headers
with open(outBomName, 'r', encoding='utf-8') as bomcsv:
    bomreader = csv.DictReader(bomcsv, delimiter=',', quotechar='"')
    with open(latexBomName,'w', encoding='utf-8', newline='') as bomcsv:
        writer = csv.DictWriter(bomcsv,fieldnames=bomreader.fieldnames)
        writer.writeheader()
        for entry in bomreader:
            # Check if it's a component group header and parse
            if entry['Designator'] == '':
                writer.writerow({'Designator': entry['Designator'],
                                 'Name': '\\centering{' + entry['Name'] + '}',
                                 'Count': entry['Count'],
                                 'Comment': entry['Comment']})
            # Check if it's a functional group header and parse
            elif ''.join([i for i in entry['Designator'] if not i.isdigit()])[0] == 'A':
                writer.writerow({'Designator': entry['Designator'],
                                 'Name': '\\centering{\\underline{' + entry['Name'] + '}}',
                                 'Count': entry['Count'],
                                 'Comment': entry['Comment']})
            else:
                writer.writerow({'Designator': entry['Designator'],
                                 'Name': entry['Name'],
                                 'Count': entry['Count'],
                                 'Comment': entry['Comment']})


# Change delimiter " " to { }
latexBom = open(latexBomName, 'r', encoding='utf-8').read()
buffer = StringIO()
quote_index = 0

with open(latexBomName,'r',encoding='utf-8') as file:
    for c in latexBom:
        if c == '"':
            buffer.write('{' if quote_index % 2 == 0 else '}')
            quote_index += 1
        else:
            buffer.write(c)

with open(latexBomName, 'w', encoding='utf-8') as file:
    file.write(buffer.getvalue())
    
# Find and replace special symbols ? to \? 
filedata = None
with open(latexBomName, 'r', encoding='utf-8') as file:
    filedata = file.read()
    filedata = filedata.replace('{}', '')
    filedata = filedata.replace('-', '"=')
    filedata = filedata.replace('"', '\"')
    filedata = filedata.replace('ф. ', 'ф.~')
    filedata = filedata.replace('фирма ', 'фирма~')
    filedata = filedata.replace('&', '\&')
    filedata = filedata.replace('%', '\%')
    filedata = filedata.replace('_', ' ')
    filedata = filedata.replace('#', '\#')
with open(latexBomName, 'w', encoding='utf-8') as file:
    file.write(filedata)

# Write ESKD data into \latex\data.tex
eskd_data = 'latex\\data.tex'

with open(eskd_data, 'r', encoding='utf-8') as file:
    content = file.read()
with open(eskd_data, 'w', encoding='utf-8') as file:
    content = re.sub('org\}\{(.*?)\}','org}{'+ bom[0]['DocumentNumberE3'].split('.',1)[0] +'}',content)
    content = re.sub('num\}\{(.*?)\}','num}{'+ bom[0]['DocumentNumberE3'].split('.',1)[1].split('Э3')[0] +'}',content)
    content = re.sub('ESKDtitle\{(.*?)\}','ESKDtitle{'+ bom[0]['Title'] +'}',content)
    content = re.sub('ESKDauthor\{(.*?)\}','ESKDauthor{'+ bom[0]['Author'] +'}',content)
    content = re.sub('ESKDchecker\{(.*?)\}','ESKDchecker{'+ bom[0]['CheckedBy'] +'}',content)
    content = re.sub('ESKDnormContr\{(.*?)\}','ESKDnormContr{'+ bom[0]['NormInspection'] +'}',content)
    content = re.sub('ESKDapprovedBy\{(.*?)\}','ESKDapprovedBy{'+ bom[0]['ApprovedBy'] +'}',content)
    content = re.sub('ESKDgroup\{(.*?)\}','ESKDgroup{'+ bom[0]['Organization'] +'}',content)
    content = re.sub('ESKDapprovedBy\{(.*?)\}','ESKDapprovedBy{'+ bom[0]['ApprovedBy'] +'}',content)
    content = re.sub('ESKDcolumnXXV\{(.*?)\}','ESKDcolumnXXV{'+ bom[0]['FirstApply'] +'}',content)
    file.write(content)