import docx
#from docx import Document
#from docx.shared import RGBColor
import re
from docSearch import *
#import xlwings

#************************************ Uses GetVerifiedUniqueTrailingTags() to find *************************************

def GetRelationalList():
    fList = []
    sList = []
    thList = []
    uniqueTagsList = GetVerifiedUniqueTrailingTags()
    trailingTagsList = GetTrailingTags()
    relLeadTags = GetRelLeadTags()
    for uTag in uniqueTagsList[0]:
        part1 = re.findall('\S.*?[\s:]', uTag)
        fList.append(part1[0].replace(':', ''))
        sList.append(part1[1].replace(':', ''))
        part2 = re.findall('[\s:].*\S', uTag)
        part2 = part2[0].split(':')
        thList.append(part2[2])
    indexList = []
    ind = 0
    for i in range(len(fList)):
        indexList.clear()
        ind = 0
        printItem = 0
        for tTag in trailingTagsList:
            tTag = tTag.split("] ")
            for item in tTag:
                if fList[i] in item and sList[i] in item and re.search(r'\b(' + thList[i] + r')\b', item):
                    if printItem == 0:
                        print(fList[i] + ":" + sList[i] + ":" + thList[i])
                    printItem = printItem + 1
                    indexList.append(ind)
            ind = ind + 1

        for id in indexList:
            indexing = relLeadTags[1][id]
            print(indexing)
        print("\n")