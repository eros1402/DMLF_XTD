# -*- coding: utf-8 -*-
# Make sure that the xlsxwriter package was installed. If not:
# $cd /home/cph/Python/Softwares/XlsxWriter-1.0.4/
# $sudo python setup.py install

#Example: compare 2 DMLF files:
# $python compareDMLF_XTD.py Sample_DMLF/90337BA.PR35.002.05 \
#                            Sample_DMLF/90337BA.PR35.002.06 \
#                            --hide PROMPT,RES,MIN,MAX,ARRAY_SIZE,LOG,DISPLAY,STATISTICS,ERROR,TEST \
#                            --ignore NUM,DESC \
#                            --renameFile RenameParam_90337.csv

# Compare 2 DMLF folders
# $python compareDMLF_XTD.py -d Sample_DMLF \
#                               Sample_DMLF \
#                            --dev 90337BA \
#                            --cond PR150,PR35,PR175 \
#                            --spec 002.05,002.06 \
#                            --hide PROMPT,RES,MIN,MAX,ARRAY_SIZE,LOG,DISPLAY,STATISTICS,ERROR,TEST \
#                            --ignore NUM,DESC \
#                            --renameFile RenameParam_90337.csv

import sys
import argparse
import os
import io
import xlsxwriter
from collections import namedtuple
from collections import OrderedDict
from datetime import datetime


#### Global variables ####
# Note: XTD DMLF file is separated in 3 parts: 1-Bin Code ; 2-Limits Parameter ;  3-Inputs Parameter
BinStruct = namedtuple ("BinStruct", "BinCode BinType BinDesc")
LimitsParamStruct = namedtuple ("LimitsParamStruct", "GROUP NUM MEAS MEAS_TYPE UNIT ARRAY_SIZE BIN LOG DISPLAY STATISTICS RANGE ERROR PROMPT TEST MIN_VALUE MAX_VALUE")
InputsParamStruct = namedtuple ("InputsParamStruct", "GROUP NUM NAME NOM_TYPE UNIT LOG DISPLAY DESC TEST NOM_VALUE")

#### Support function ####
def get1stColumnPosOfOldParamsInDiffSheet ():
  return 1


def get1stColumnPosOfNewParamsInDiffSheet (diffType):
  paramFields = BinStruct._fields
  if diffType == 2:
    paramFields = LimitsParamStruct._fields
  elif diffType == 3:
    paramFields = InputsParamStruct._fields

  colPos = get1stColumnPosOfOldParamsInDiffSheet()
  for field in paramFields:
    colPos += 1
  return colPos


def LogMessage (str):
  global LogSheet
  global PrintInTerminal
  global LogRow

  timestamp = "%s" % datetime.now ()
  outputFileLogSheet.write (LogRow, 0, LogRow)
  outputFileLogSheet.write (LogRow, 1, timestamp)
  outputFileLogSheet.write (LogRow, 2, str)

  LogRow += 1
  return


def parseRenameFile (pathFile, renameTable):
#   print ("Process file: " + pathFile)
  file = io.open (pathFile, "r", encoding = 'utf-8')
  row = 0
  for line in file:
      splittedLine = line.strip ('\n').split (',', 1)
      if (len (splittedLine) >= 2):
          old_param = splittedLine[0]
          new_param = splittedLine[1].strip('\r').strip(' ').replace(',','')
          renameTable[old_param] = new_param
          outputFileRenameSheet.write(row, 0, old_param)
          outputFileRenameSheet.write(row, 1, new_param)
          row += 1

  file.close ()
  return


def parseDMLFFile (pathFile, binDict, limitsParamDict, inputsParamDict):
#     print ("Process file: " + pathFile)
    dmlfFile = io.open (pathFile, "r", encoding = 'utf-8')
    lineNum = 1
    readPart = 0  # readPart = 1 : Start reading bincodes
                  # readPart = 2 : Start reading fields of limits parameters
                  # readPart = 3 : Start reading fields of inputs parameters
    limitsParam = dict ()
    inputsParam = dict ()
    paramName = ''
    for line in dmlfFile:
        splittedLine = line.strip ('\n').split ('=', 1) # Split the line in 2 parts by '=' as splitter

        if splittedLine[0] == 'BINNUM':
#           print ("Start reading bincode at lineNum: " + lineNum.__str__())
          readPart = 1
        elif splittedLine[0] == 'PARAM_NUM':
#           print ("Start reading Limits Parameter at lineNum: " + lineNum.__str__())
          readPart = 2
        elif splittedLine[0] == 'IN_PARAM_NUM':
#           print ("Start reading Inputs Parameter at lineNum: " + lineNum.__str__())
          readPart = 3
        else:
          if (len (splittedLine) == 2):
            paramField = splittedLine[0]
            paramValue = splittedLine[1]

            ### Read bin code
            if readPart == 1:
              binCode = int (paramField.strip ('BIN'))     # get bin Code
              binType_Code = paramValue.split (',', 1)[0] # get bin Type code: 0-Pass , 1-Fail, 2-Retest
              if (binType_Code == '0')   : binType = "Pass"
              elif (binType_Code == '1') : binType = "Failed"
              elif (binType_Code == '2') : binType = "Retest"
              binDesc = paramValue.split (',', 1)[1].strip () # get bin Description
              param = binDesc[:-6]  # Remove 'Failed' or 'Passed' or 'Retest' at the end of binDisc to get parameter name
#               print (binDesc)
              binDict[param] = BinStruct (BinCode = binCode, BinType = binType, BinDesc = binDesc)

            ### Read Limits parameters
            elif readPart == 2:
              if paramField in LimitsParamStruct._fields: ## Only check fields that are defined in LimitsParamStruct
                if paramField == 'NUM':
                  if 0 < len (limitsParam):
                    limitsParamDict[paramName] = LimitsParamStruct (**limitsParam)
                  limitsParam.clear ()    # Clear parameter struct before reading for each parameter
                  paramName = ''

                limitsParam[paramField] = paramValue
                if paramField == 'MEAS':
                  paramName = paramValue
                else:
                  if (paramField in ('NUM', 'BIN', 'ARRAY_SIZE')):
                    if paramValue != '':
                      limitsParam[paramField] = int (paramValue)
                  elif (paramField in ('MIN_VALUE', 'MAX_VALUE')):
                    if paramValue != '':
                      if limitsParam['MEAS_TYPE'] == 'FLOAT':
                        limitsParam[paramField] = float (paramValue)
                      elif limitsParam['MEAS_TYPE'] == 'INTEGER':
                        limitsParam[paramField] = int (paramValue)

#               if (lineNum > 7000) & (lineNum < 7050):
#                 print (limitsParam)

            ### Read Inputs parameters
            elif readPart == 3:
              if paramField in InputsParamStruct._fields: ## Only check fields that are defined in InputssParamStruct
                if paramField == 'NUM':
                  if 0 < len (inputsParam):
                    inputsParamDict[paramName] = InputsParamStruct (**inputsParam)
                  inputsParam.clear ()    # Clear parameter struct before reading for each parameter
                  paramName = ''

                inputsParam[paramField] = paramValue
                if paramField == 'NAME':
                  paramName = paramValue
                elif paramField == 'NUM':
                  inputsParam[paramField] = int (paramValue)
                elif paramField == 'NOM_VALUE':
                  if inputsParam['NOM_TYPE'] == 'FLOAT':
                    inputsParam[paramField] = float (paramValue)
                  elif inputsParam['NOM_TYPE'] == 'INTEGER':
                    inputsParam[paramField] = int (paramValue)

#               if (lineNum > 26100) & (lineNum < 26150):
#                 print (inputsParam)

        lineNum += 1
    dmlfFile.close ()
    return


def isParamRenamed (param, renameTable):
  global args

  doRename = False
  if args.renameFile:
    if param in renameTable.keys():
      doRename = True

  return doRename

def getParam1ByParam2InRenameTable (RenameTable, param2):
  for p1, p2 in RenameTable.items():
    if p2 == param2:
      return p1


def doDiff (diffSheet, startRow, renameTable, nameFile1, paramDict1, nameFile2, paramDict2, diffType = 2):
  global args
  global xlsxFormatRemoved
  global xlsxFormatAdded
  global xlsxFormatRenamed
  global xlsxFormatChanged

  rowPos = startRow
  sortedParamDict1 = dict()
  sortedParamDict2 = dict()
  paramNameFields = ['BinDesc', 'MEAS', 'PROMPT', 'RES', 'MIN', 'MAX', 'NOM', 'NAME', 'DESC']

  if diffType == 1:
    sortedParamDict1 = OrderedDict (sorted (paramDict1.items (), key = lambda x: x[1].BinCode))
    sortedParamDict2 = OrderedDict (sorted (paramDict2.items (), key = lambda x: x[1].BinCode))
    paramFields = BinStruct._fields
  elif diffType == 2:
    sortedParamDict1 = OrderedDict (sorted (paramDict1.items (), key = lambda x: x[1].NUM))
    sortedParamDict2 = OrderedDict (sorted (paramDict2.items (), key = lambda x: x[1].NUM))
    paramFields = LimitsParamStruct._fields
  elif diffType == 3:
    sortedParamDict1 = OrderedDict (sorted (paramDict1.items (), key = lambda x: x[1].NUM))
    sortedParamDict2 = OrderedDict (sorted (paramDict2.items (), key = lambda x: x[1].NUM))
    paramFields = InputsParamStruct._fields


  #*** Parameters ***
  if diffType == 1:
    diffSheet.write (rowPos, 0, 'Compare Bincodes: ' + nameFile1 + ' vs ' + nameFile2, xlsxBoldTextFormat)
  elif diffType == 2:
    diffSheet.write (rowPos, 0, 'Compare Limits Parameters: ' + nameFile1 + ' vs ' + nameFile2, xlsxBoldTextFormat)
  else:
    diffSheet.write (rowPos, 0, 'Compare Inputs Parameters: ' + nameFile1 + ' vs ' + nameFile2, xlsxBoldTextFormat)
  rowPos += 1

  # Write Legends
  diffSheet.write (rowPos, 1, 'Legends')
  rowPos += 1

  diffSheet.write (rowPos, 0, 'Removed parameters')
  numOfRemovedParameters = 0
  rowParamRemoved = rowPos
  rowPos += 1

  diffSheet.write (rowPos, 0, 'Added parameters')
  numOfAddedParameters = 0
  rowParamAdded = rowPos
  rowPos += 1

  diffSheet.write (rowPos, 0, 'Changed parameters')
  numOfChangedParameters = 0
  rowParamChanged = rowPos
  rowPos += 1

  rowParamRenamed = 0
  numOfRenamedParameters = 0
  if args.renameFile:
    diffSheet.write (rowPos, 0, 'Renamed parameters')
    rowParamRenamed = rowPos
    rowPos += 1

  diffSheet.write (rowPos, 0, 'Not changed parameters')
  numOfNoChangedParameters = 0
  rowParamNotChanged = rowPos
  rowPos += 1

  # Add headers
  diffSheet.write (rowPos, get1stColumnPosOfOldParamsInDiffSheet(), nameFile1)
  diffSheet.write (rowPos, get1stColumnPosOfNewParamsInDiffSheet(diffType), nameFile2)
  rowPos += 1
  diffSheet.write (rowPos, 0, 'RESULT')

  colCnt = 1;
  for i in range (1, 3):
    for field in paramFields:
      diffSheet.write (rowPos, colCnt, field + '_' + str(i))
      colCnt += 1

  rowPos += 1

  ### Check removed parameters
  for param1, paramStruct1 in sortedParamDict1.items ():
    if (isParamRenamed(param1, renameTable)):
      param1 = renameTable[param1]

    if param1 not in sortedParamDict2:
#         print("Removed parameter: " + param1)
        diffSheet.write (rowPos, 0, 'Removed')
        colPos = get1stColumnPosOfOldParamsInDiffSheet ()
        for field in paramStruct1:
          diffSheet.write (rowPos, colPos, field, xlsxFormatRemoved)
          colPos += 1

        numOfRemovedParameters += 1
        rowPos += 1

  ### Check added parameters
  for param2, paramStruct2 in sortedParamDict2.items ():
    if (args.renameFile) and (param2 in renameTable.values()):
      param2 = getParam1ByParam2InRenameTable (renameTable, param2)

    if param2 not in sortedParamDict1:
#       print("Added parameter: " + param2)
        diffSheet.write (rowPos, 0, 'Added')

        colPos = get1stColumnPosOfNewParamsInDiffSheet (diffType)
        for field in paramStruct2:
          diffSheet.write (rowPos, colPos, field, xlsxFormatAdded)
          colPos += 1
        numOfAddedParameters += 1
        rowPos += 1

  ### Check changed parameters
  for param1, paramStruct1 in sortedParamDict1.items ():
    renamedParam1 = param1
    doRename = False
    isParamNameField = False
    if (args.renameFile) and (param1 in renameTable.keys()):
      doRename = True
      renamedParam1 = renameTable[param1]

    if renamedParam1 in sortedParamDict2:
        if (doRename): numOfRenamedParameters += 1
        paramStruct2 = sortedParamDict2[renamedParam1]
        changedFields = dict ()
        isParamChanged = False;

        for field in paramFields:
          isParamNameField = field in paramNameFields
          changedFields[field] = False if ((field in args.ignore) or ((doRename) and (isParamNameField))) else getattr(paramStruct1, field) != getattr(paramStruct2, field)
          isParamChanged = isParamChanged or changedFields[field]

        if (isParamChanged or doRename):
#           print("Changed parameter: " + param1)
          if isParamChanged:
            diffSheet.write (rowPos, 0, 'Changed')
          else:
            diffSheet.write (rowPos, 0, 'Changed name only')
          colPos1 = get1stColumnPosOfOldParamsInDiffSheet ()
          colPos2 = get1stColumnPosOfNewParamsInDiffSheet (diffType)
          for field in paramFields:
            isParamNameField = field in paramNameFields
            if ((doRename) and (isParamNameField)):
              diffSheet.write (rowPos, colPos1, getattr(paramStruct1, field), xlsxFormatRenamed)
              diffSheet.write (rowPos, colPos2, getattr(paramStruct2, field), xlsxFormatRenamed)
            else:
              diffSheet.write (rowPos, colPos1, getattr(paramStruct1, field), xlsxFormatChanged if changedFields[field] else None)
              diffSheet.write (rowPos, colPos2, getattr(paramStruct2, field), xlsxFormatChanged if changedFields[field] else None)
            colPos1 += 1
            colPos2 += 1

          if isParamChanged:
            numOfChangedParameters += 1
          rowPos += 1

  # Report not changed parameters
  for param1, paramStruct1 in sortedParamDict1.items ():
    if param1 in sortedParamDict2:
      paramStruct2 = sortedParamDict2[param1]
      changedFields = dict ()
      isParamChanged = False;
      for field in paramFields:
        changedFields[field] = False if field in args.ignore else getattr(paramStruct1, field) != getattr(paramStruct2, field)
        isParamChanged = isParamChanged or changedFields[field]

      if not isParamChanged:
        diffSheet.write (rowPos, 0, 'Not changed')
        colPos1 = get1stColumnPosOfOldParamsInDiffSheet ()
        colPos2 = get1stColumnPosOfNewParamsInDiffSheet (diffType)
        for field in paramFields:
          diffSheet.write (rowPos, colPos1, getattr(paramStruct1, field))
          diffSheet.write (rowPos, colPos2, getattr(paramStruct2, field))
          colPos1 += 1
          colPos2 += 1

        # Hide no changed rows
        diffSheet.set_row (rowPos,None, None, {'hidden': True})

        rowPos += 1
        numOfNoChangedParameters += 1


  diffSheet.write (rowParamRemoved, 1, numOfRemovedParameters, xlsxFormatRemoved)
  diffSheet.write (rowParamAdded, 1, numOfAddedParameters, xlsxFormatAdded)
  diffSheet.write (rowParamChanged, 1, numOfChangedParameters, xlsxFormatChanged)
  if args.renameFile:
    diffSheet.write (rowParamRenamed, 1, numOfRenamedParameters, xlsxFormatRenamed)
  diffSheet.write (rowParamNotChanged, 1, numOfNoChangedParameters, None)

  return (rowPos + 1)


def hideParamFields (outputFileDiffSheet, hideFields, isLimitsParam = True):
  colParamsDict1 = dict ()
  colParamsDict2 = dict ()

  paramFields = LimitsParamStruct._fields
  if not isLimitsParam:
    paramFields = InputsParamStruct._fields

  colPos = 1
  for field in paramFields:
    colParamsDict1[field] = colPos
    colPos += 1
  for field in paramFields:
    colParamsDict2[field] = colPos
    colPos += 1

  for field in hideFields:
#       print("Hidden field:" + field)
    if field in paramFields:
      outputFileDiffSheet.set_column(colParamsDict1[field], colParamsDict1[field],None, None, {'hidden': True})
      outputFileDiffSheet.set_column(colParamsDict2[field], colParamsDict2[field],None, None, {'hidden': True})

  return



#### __main__ ####

# Adding arguments from cmd line

parser = argparse.ArgumentParser (description='Compare XTD DMLF files')
parser.add_argument ('path1', help='path of the old DMLF file (default) or folder (with -d/--dir)')
parser.add_argument ('path2', help='path of the new DMLF file (default) or folder (with -d/--dir)')
parser.add_argument ('-d', '--dir', action='store_true', help='Indicate path1 & path2 are to directories')
parser.add_argument ('--dev', default = '', help = 'Device name, only needed when paths are directories. E.g --dev 90337BA')
parser.add_argument ('--cond', default = '', help = 'The list of compared condtions, only needed when paths are directories. E.g --cond PR35,PR150,PR175')
parser.add_argument ('--spec', default = '', help = 'The list of compared spec versions, only needed when paths are directories and can only put max 2 spec versions. E.g --spec 002.05,002.06')
parser.add_argument ('--renameFile', default = '', help='The link of renamed parameter file. E.g --renameFile renamedParam_12125.csv')
parser.add_argument ('--ignore', default='', choices=LimitsParamStruct._fields, help='The list of ignored fields in the parameter comparison')
parser.add_argument ('--hide', default='', help='The list of hided fields in the report file. E.g --hide PROMPT,MIN,MAX,ERROR,DISPLAY,RES,BIN,ARRAY_SIZE,LOG,ERROR,GROUP,TEST')
args = parser.parse_args ()

path1 = os.path.abspath (args.path1)
path2 = os.path.abspath (args.path2)
baseName1 = os.path.basename (args.path1)
baseName2 = os.path.basename (args.path2)

# Check the Paths are valid & other arguments
if (args.dir):
    if ('.' in baseName1) :
        sys.exit ("Error: Invalid directory path : " + path1)
    elif ('.' in baseName2):
        sys.exit ("Error: Invalid directory path : " + path2)

    if not (args.dev):
        sys.exit ("Error: Please specify device name with --dev=deviceName")
    else:
        device = args.dev

    if not (args.cond):
        sys.exit ("Error: Please specify conditions with --cond='listOfConditions'")
    else:
        conds = args.cond.strip('\'')
        conditions = conds.split(',')

    if not (args.spec):
        sys.exit ("Error: Please spec versions with --spec='listOfSpecVersion'")
    else:
        specs = args.spec.strip('\'')
        versions = specs.split(',')
        if (len(versions) > 2) :
          sys.exit ("Error: Only can put max 2 spec conditions")
else:
    if ('.' not in baseName1) :
        sys.exit("Error: Invalid DMLF file path : " + path1)
    elif ('.' not in baseName2):
        sys.exit("Error: Invalid DMLF file path : " + path2)

comparedFiles = dict ()
comparedFiles.clear()
ver1 = ''
ver2 = ''
if(args.dir):
    if len(versions) == 2:
      ver1 = versions[0]
      ver2 = versions[1]
    elif len(versions) == 1:
      ver1 = versions[0]
      ver2 = versions[0]

    comparedFiles.clear()
    for cond in conditions:
      fileVer1 = device + '.' + cond + '.' + ver1
      fileVer2 = device + '.' + cond + '.' + ver2
      comparedFiles[fileVer1] = fileVer2
else:
  comparedFiles[baseName1] = baseName2

# Define output file:
#        Output file sheets : Log sheet - Diff_Bincodes - Diff_Limits - Diff_Inputs
name1 = baseName1.replace('.','_')
name2 = baseName2.replace('.','_')
pathOutputFile = os.path.abspath (os.path.join (os.getcwd (), 'diff_' + name1 + '_vs_' + name2 + '.xlsx'))
if (args.dir):
  pathOutputFile = os.path.abspath (os.path.join (os.getcwd (), 'diff_' + device + '_SP' + ver1 + '_vs_SP' + ver2 + '.xlsx'))
outputFile = xlsxwriter.Workbook (pathOutputFile)
outputFileLogSheet = outputFile.add_worksheet ('Log')   # pointer to Log sheet
outputFileDiffBincodesSheet = outputFile.add_worksheet ('Diff_Bincodes') # pointer to Diff_Limits sheet
outputFileDiffLimitsSheet = outputFile.add_worksheet ('Diff_Limits') # pointer to Diff_Limits sheet
outputFileDiffInputsSheet = outputFile.add_worksheet ('Diff_Inputs') # pointer to Diff_Inputs sheet
# outputFileDiffSheet = outputFile.add_worksheet ('Diff') # pointer to Diff_Limits sheet

xlsxBoldTextFormat = outputFile.add_format ({'bold': True})
# Define format for the comparison:
#    Removed field: Red filled
#    Added field: Cyan filled
#    Changed field: Yellow filled
#    Renamed field: Green filled
#    No Changed field: No color
xlsxFormatRemoved = outputFile.add_format ()
xlsxFormatRemoved.set_bg_color ('#FF0000')
xlsxFormatAdded = outputFile.add_format ()
xlsxFormatAdded.set_bg_color ('#00FFFF')
xlsxFormatChanged = outputFile.add_format ()
xlsxFormatChanged.set_bg_color ('#FFFF00')
xlsxFormatRenamed = outputFile.add_format ()
xlsxFormatRenamed.set_bg_color ('#00FF00')

LogRow = 0
rowDiffBinSheet = 0
rowDiffLimitSheet = 0
rowDiffInputSheet = 0

# Read rename file
renameTable = dict ()
renameTable['Old_ParamName'] = 'New_ParamName' # Dummy rename table
if args.renameFile:
  # Update rename table & write to the output file
  pathRenameFile = os.path.abspath (args.renameFile)
  outputFileRenameSheet = outputFile.add_worksheet ('RenameTable')
  parseRenameFile (pathRenameFile, renameTable)
  LogMessage ('Read ' + pathRenameFile)


for file1, file2 in comparedFiles.items():
  if(args.dir):
    pathFile1 = os.path.join(path1, file1)
    pathFile2 = os.path.join(path2, file2)
  else:
    pathFile1 = path1
    pathFile2 = path2

  binDict1 = dict ()
  limitsParamDict1 = dict ()
  inputsParamDict1 = dict ()
  parseDMLFFile (pathFile1, binDict1, limitsParamDict1, inputsParamDict1)

  # Read DMLF files: path2
  binDict2 = dict ()
  limitsParamDict2 = dict ()
  inputsParamDict2 = dict ()
  parseDMLFFile (pathFile2, binDict2, limitsParamDict2, inputsParamDict2)

  # Output file: Write to Log sheet

  LogMessage ('Read ' + pathFile1)
  LogMessage ('Read ' + pathFile2)
  LogMessage ('Create diff report in ' + pathOutputFile)

  # Output file: Write to Diff_Bincodes sheet
  rowDiffBinSheet = doDiff (outputFileDiffBincodesSheet, rowDiffBinSheet, renameTable, file1, binDict1, file2, binDict2, 1)

  # Output file: Write to Diff_Limits sheet
  rowDiffLimitSheet = doDiff (outputFileDiffLimitsSheet, rowDiffLimitSheet, renameTable, file1, limitsParamDict1, file2, limitsParamDict2, 2)

  # Output file: Write to Diff_Inputs sheet
  rowDiffInputSheet = doDiff (outputFileDiffInputsSheet, rowDiffInputSheet, renameTable, file1, inputsParamDict1, file2, inputsParamDict2, 3)


# Print to terminal
print ("path1," + path1)
print ("path2," + path2)
if args.renameFile:
    print ("RenameFile," + pathRenameFile)
print ("OutputFile," + pathOutputFile)

# Output file: hide some compared fields of Limits parameters
if args.hide:
  hideFields = args.hide.split(',')
  hideParamFields (outputFileDiffLimitsSheet, hideFields, True)
  hideParamFields (outputFileDiffInputsSheet, hideFields, False)

outputFile.close () # saving output file
