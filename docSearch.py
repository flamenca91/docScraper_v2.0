#from __future__ import annotations
import docx
#from docx import Document
#from docx.shared import RGBColor
import re
#from docExtract import *
#import xlwings

docFile = {}

f = open("DocTagsList.txt")
for line in f:
    line = line.split()
    docFile[line[0]] = line[1]

filePath = "C:/Users/steph/OneDrive/Desktop/Docs_Project/"
docFileList = list(docFile.keys())                  # This is a list of all main tags found in each document

# ************************************* File Handling ***********************************************

def GetText(filename):                      # Opens the document and places each paragraph into a list
    doc = docx.Document(filename)
    fullText = []
    for para in doc.paragraphs:
        fullText.append(para.text)
    fullText = [ele for ele in fullText if ele.strip()]   # Eliminates empty paragraphs or spaces
    return fullText                                       # List of all document content

# ************************************ Raw Leading Tags *********************************************

def GetLeadingTags():                   # Returns only valid parent tags
    leadingTags = []                    # List of Leading Tags
    tagDescriptions = []
    for tag in docFileList:             # Tags are used to open the corresponding file
        textList = GetText(filePath + docFile[tag])
        index = 0
        ind = []
        for t in textList:
            if tag == "BOLUS" or tag == "ACE":              # SRS has three additional tags - AID, BOLUS and ACE
                if re.search('.*[:\s]' + "SRS" + '[:\s]', t):   # This if block does the same as the else block
                    ind.append(index)
                    tagDescriptions.append(t)
                    y = re.findall('\S*[:\s]' + "SRS" + '[:\s]\S*', t)
                    leadingTags.append(y[0])
                index = index + 1
            else:
                if re.search('.*[:\s]' + re.escape(tag) + '[:\s]', t):
                    ind.append(index)
                    tagDescriptions.append(t)
                    y = re.findall('\S*[:\s]' + re.escape(tag) + '[:\s]\S*', t)
                    leadingTags.append(y[0])
                index = index + 1
    leadTagsAndDescriptions = [leadingTags, tagDescriptions]
    return leadTagsAndDescriptions


def GetRelLeadTags():        # Eliminates all leading tags with that have no trailing tags (RISK & URS)
    indx = 0
    indxList = []
    relationalLeadingTags = []
    leadTagsList = GetLeadingTags()
    for tag in leadTagsList[0]:
        if re.search('[:\s]' + "RISK" + '[:\s]', tag) or re.search('[:\s]' + "URS" + '[:\s]', tag):
            indxList.append(indx)
        else:
            relationalLeadingTags.append(tag)
        indx = indx + 1
    relationalLeadingTags.sort()
    indxList.reverse()
    for i in indxList:
        leadTagsList[1].pop(i)
    print(indxList)
    print(len(leadTagsList[1]))
    return [relationalLeadingTags, leadTagsList[1]]


#************************************* Trailing Tags ******************************************

def GetTrailingTags():
    trailingTags = []                   # List of Trailing Tags
    for tag in docFileList:             # Tags are used to open the corresponding file
        textList = GetText(filePath + docFile[tag])
        index = 0
        ind = []
        for t in textList:
            if tag == "BOLUS" or tag == "ACE":
                if re.search('.*[:\s]' + "SRS" + '[:\s]', t):
                    ind.append(index)
                    y = re.findall('[\[{].+[\]}]', t)
                    if len(y) != 0:
                        trailingTags.append(y[0])
                index = index + 1
            else:
                if re.search('.*[:\s]' + re.escape(tag) + '[:\s]', t):
                    ind.append(index)
                    y = re.findall('[\[{].+[\]}]', t)
                    if len(y) != 0:
                        trailingTags.append(y[0])
                index = index + 1
    return trailingTags

#********** Eliminates any duplicate trailing tags or PASS/FAIL descriptors and sorts the list alphabetically **********
def GetUniqueTrailingTags():
    uniqueParentTags = []
    trailingTagsList = GetTrailingTags()
    for tags in trailingTagsList:
        tags = tags.replace('[', "").replace(']', "")
        tags = tags.split()
        for item in tags:
            uniqueParentTags.append(item)
    uniqueTags = list(set(uniqueParentTags))
    uniqueTags.remove('{NA}')
    uniqueTags.remove('{PASS}')
    uniqueTags.remove('{FAIL}')
    uniqueTags.sort()
    return uniqueTags

def GetVerifiedUniqueTrailingTags():        # Returns only valid tags and extracts Orphan tags
    verifiedTrailingTags = []
    orphanTags = []
    tempTags = GetUniqueTrailingTags()
    for tt in tempTags:
        tag = re.findall('[:\s]' + ".*" + '[:\s]', tt)
        tag = tag[0].replace(':', "")
        if tag in docFileList:
            verifiedTrailingTags.append(tt)
        else:
            orphanTags.append(tt)
    verifiedTags = [verifiedTrailingTags,orphanTags]
    return verifiedTags