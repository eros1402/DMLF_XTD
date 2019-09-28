# -*- coding: utf-8 -*-
# Make sure that the xlsxwriter package was installed. If not:
# $cd /home/cph/Python/Softwares/XlsxWriter-1.0.4/
# $sudo python setup.py install

#Example: Cmd with translation file: 
# $python compareDMLF.py 12132AC.PR35.003.19 12132AC.PR35.004.01 --transFile TransFile_12125.csv --hide PROMPT,MIN,MAX,NOM,ERROR,DISPLAY,RES,BIN,ARRAY_SIZE,LOG,ERROR
# $python compareDMLF.py 12125DA.PR35.003.08 12125DA.PR35.004.01 --transFile TransFile_12125.csv --hide PROMPT,MIN,MAX,NOM,ERROR,DISPLAY,RES,BIN,ARRAY_SIZE,LOG,ERROR
 
import argparse
import os
import xlsxwriter
from collections import namedtuple
from collections import OrderedDict
from datetime import datetime

BinStruct = namedtuple ("BinStruct", "BinType BinDesc")
ParamStruct = namedtuple ("ParamStruct", "NUM PROMPT UNIT MEAS_TYPE BIN RES MIN MIN_VALUE MAX MAX_VALUE NOM NOM_VALUE ARRAY_SIZE LOG DISPLAY STATISTICS RANGE ERROR")

def LogMessage (str):
  global LogSheet
  global PrintInTerminal
  global LogRow

  timestamp = "%s" % datetime.now ()
  outputFileLogSheet.write (LogRow, 0, LogRow)
  outputFileLogSheet.write (LogRow, 1, timestamp)
  outputFileLogSheet.write (LogRow, 2, str)
  
  LogRow += 1

def parseDMLFFile (pathFile, binDict, paramDict):
    dmlfFile = open (pathFile, "r")

    paramItem = dict ()
    paramMeas = ''
    readMode = 0    # 1 = Reading bnicodes, 2 = Reading parameters    
    for line in dmlfFile:
        splittedLine = line.strip ('\n').split ('=', 1)
        if splittedLine[0] == 'BINNUM':
            readMode = 1
        elif splittedLine[0] == 'PARAM_NUM':
            readMode = 2
        else:
            if (len (splittedLine) == 2):
                if readMode == 1:
                    binNo = splittedLine[0].strip ('BIN')
                    binType = splittedLine[1].split (',', 1)[0]
                    binDesc = splittedLine[1].split (',', 1)[1]
                    binDict[binNo] = BinStruct (BinType = binType, BinDesc = binDesc)

                elif readMode == 2:                    
                    if splittedLine[0] == 'NUM':
                        if 0 < len (paramItem):
                            paramDict[paramMeas] = ParamStruct (**paramItem)
                        paramItem.clear ()
                        paramMeas = ''
                        
                    if splittedLine[0] == 'MEAS':
                        paramMeas = splittedLine[1]
                    else:
                        if (splittedLine[0] in ('NUM', 'BIN', 'ARRAY_SIZE')):
                            paramItem[splittedLine[0]] = int (splittedLine[1])
                        elif (splittedLine[0] in ('MIN_VALUE', 'MAX_VALUE', 'NOM_VALUE')):
                            if paramItem['MEAS_TYPE'] in ('INTEGER', 'SINGLE'):
                                paramItem[splittedLine[0]] = float (splittedLine[1])
                            else:
                                paramItem[splittedLine[0]] = splittedLine[1]
                        else:                        
                            paramItem[splittedLine[0]] = splittedLine[1]
                  
    dmlfFile.close ()
    
    # print (binDict)
    # print (paramDict)

def parseTransFile (pathFile, transTable):
    transFile = open (pathFile, "r")    
    TransRow = 0
    for line in transFile:
        splittedLine = line.strip ('\n').split (',', 1)
        if (len (splittedLine) >= 2):
            old_param = splittedLine[0]
            new_param = splittedLine[1].strip('\r').strip(' ').replace(',','')
            transTable[old_param] = new_param  
            outputFileTransSheet.write(TransRow, 0, old_param) 
            outputFileTransSheet.write(TransRow, 1, new_param) 
            TransRow += 1             
    
    transFile.close ()
    
def doDiff (diffSheet, nameFile1, nameFile2, binDict1, paramDict1, binDict2, paramDict2, xlsxFormatChanged):
    global args
    
    rowCnt = 0
    
    sortedParamDict1 = OrderedDict (sorted (paramDict1.items (), key = lambda x: x[1].NUM))
    sortedParamDict2 = OrderedDict (sorted (paramDict2.items (), key = lambda x: x[1].NUM))

    #*** Parameters ***
    # Overview
    diffSheet.write (rowCnt, 0, 'Parameters removed ')
    rowParamRemoved = rowCnt
    rowCnt += 1
    diffSheet.write (rowCnt, 0, 'Parameters added')
    rowParamAdded = rowCnt
    rowCnt += 1
    diffSheet.write (rowCnt, 0, 'Parameters changed')
    rowParamChanged = rowCnt
    rowCnt += 1
    diffSheet.write (rowCnt, 0, 'Parameters not changed')
    rowParamNotChanged = rowCnt
    rowCnt += 1

    # Add headers
    diffSheet.write (rowCnt, 1, nameFile1)
    diffSheet.write (rowCnt, 20, nameFile2)
    rowCnt += 1
    diffSheet.write (rowCnt, 0, 'Compare result')
    diffSheet.write (rowCnt, 1, 'Meas')
    diffSheet.write (rowCnt, 2, 'Num')
    diffSheet.write (rowCnt, 3, 'Prompt')
    diffSheet.write (rowCnt, 4, 'Unit')
    diffSheet.write (rowCnt, 5, 'Meas type')
    diffSheet.write (rowCnt, 6, 'Bin')
    diffSheet.write (rowCnt, 7, 'Res')
    diffSheet.write (rowCnt, 8, 'Min')
    diffSheet.write (rowCnt, 9, 'Min value')
    diffSheet.write (rowCnt, 10, 'Max')
    diffSheet.write (rowCnt, 11, 'Max value')
    diffSheet.write (rowCnt, 12, 'Nom')
    diffSheet.write (rowCnt, 13, 'Nom value')
    diffSheet.write (rowCnt, 14, 'Array size')
    diffSheet.write (rowCnt, 15, 'Log')
    diffSheet.write (rowCnt, 16, 'Display')
    diffSheet.write (rowCnt, 17, 'Statistics')
    diffSheet.write (rowCnt, 18, 'Range')
    diffSheet.write (rowCnt, 19, 'Error')
    diffSheet.write (rowCnt, 20, 'Meas')
    diffSheet.write (rowCnt, 21, 'Num')
    diffSheet.write (rowCnt, 22, 'Prompt')
    diffSheet.write (rowCnt, 23, 'Unit')
    diffSheet.write (rowCnt, 24, 'Meas type')
    diffSheet.write (rowCnt, 25, 'Bin')
    diffSheet.write (rowCnt, 26, 'Res')
    diffSheet.write (rowCnt, 27, 'Min')
    diffSheet.write (rowCnt, 28, 'Min value')
    diffSheet.write (rowCnt, 29, 'Max')
    diffSheet.write (rowCnt, 30, 'Max value')
    diffSheet.write (rowCnt, 31, 'Nom')
    diffSheet.write (rowCnt, 32, 'Nom value')
    diffSheet.write (rowCnt, 33, 'Array size')
    diffSheet.write (rowCnt, 34, 'Log')
    diffSheet.write (rowCnt, 35, 'Display')
    diffSheet.write (rowCnt, 36, 'Statistics')
    diffSheet.write (rowCnt, 37, 'Range')
    diffSheet.write (rowCnt, 38, 'Error')
    rowCnt += 1
    
    addedParameters = 0
    changedParameters = 0
    notChangedParameters = 0
    removedParameters = 0

    # Report removed parameters
    for meas1, paramStruct1 in sortedParamDict1.items ():
        if meas1 not in sortedParamDict2:
            diffSheet.write (rowCnt, 0, 'Removed')
            diffSheet.write (rowCnt, 1, meas1, xlsxFormatChanged)
            diffSheet.write (rowCnt, 2, paramStruct1.NUM, xlsxFormatChanged)
            diffSheet.write (rowCnt, 3, paramStruct1.PROMPT, xlsxFormatChanged)
            diffSheet.write (rowCnt, 4, paramStruct1.UNIT, xlsxFormatChanged)
            diffSheet.write (rowCnt, 5, paramStruct1.MEAS_TYPE, xlsxFormatChanged)
            diffSheet.write (rowCnt, 6, paramStruct1.BIN, xlsxFormatChanged)
            diffSheet.write (rowCnt, 7, paramStruct1.RES, xlsxFormatChanged)
            diffSheet.write (rowCnt, 8, paramStruct1.MIN, xlsxFormatChanged)
            diffSheet.write (rowCnt, 9, paramStruct1.MIN_VALUE, xlsxFormatChanged)
            diffSheet.write (rowCnt, 10, paramStruct1.MAX, xlsxFormatChanged)
            diffSheet.write (rowCnt, 11, paramStruct1.MAX_VALUE, xlsxFormatChanged)
            diffSheet.write (rowCnt, 12, paramStruct1.NOM, xlsxFormatChanged)
            diffSheet.write (rowCnt, 13, paramStruct1.NOM_VALUE, xlsxFormatChanged)
            diffSheet.write (rowCnt, 14, paramStruct1.ARRAY_SIZE, xlsxFormatChanged)
            diffSheet.write (rowCnt, 15, paramStruct1.LOG, xlsxFormatChanged)
            diffSheet.write (rowCnt, 16, paramStruct1.DISPLAY, xlsxFormatChanged)
            diffSheet.write (rowCnt, 17, paramStruct1.STATISTICS, xlsxFormatChanged)
            diffSheet.write (rowCnt, 18, paramStruct1.RANGE, xlsxFormatChanged)
            diffSheet.write (rowCnt, 19, paramStruct1.ERROR, xlsxFormatChanged)

            removedParameters += 1
            rowCnt += 1
            
    # Report added parameters
    for meas2, paramStruct2 in sortedParamDict2.items ():
        if meas2 not in sortedParamDict1:
            diffSheet.write (rowCnt, 0, 'Added')
            diffSheet.write (rowCnt, 20, meas2, xlsxFormatChanged)
            diffSheet.write (rowCnt, 21, paramStruct2.NUM, xlsxFormatChanged)
            diffSheet.write (rowCnt, 22, paramStruct2.PROMPT, xlsxFormatChanged)
            diffSheet.write (rowCnt, 23, paramStruct2.UNIT, xlsxFormatChanged)
            diffSheet.write (rowCnt, 24, paramStruct2.MEAS_TYPE, xlsxFormatChanged)
            diffSheet.write (rowCnt, 25, paramStruct2.BIN, xlsxFormatChanged)
            diffSheet.write (rowCnt, 26, paramStruct2.RES, xlsxFormatChanged)
            diffSheet.write (rowCnt, 27, paramStruct2.MIN, xlsxFormatChanged)
            diffSheet.write (rowCnt, 28, paramStruct2.MIN_VALUE, xlsxFormatChanged)
            diffSheet.write (rowCnt, 29, paramStruct2.MAX, xlsxFormatChanged)
            diffSheet.write (rowCnt, 30, paramStruct2.MAX_VALUE, xlsxFormatChanged)
            diffSheet.write (rowCnt, 31, paramStruct2.NOM, xlsxFormatChanged)
            diffSheet.write (rowCnt, 32, paramStruct2.NOM_VALUE, xlsxFormatChanged)
            diffSheet.write (rowCnt, 33, paramStruct2.ARRAY_SIZE, xlsxFormatChanged)
            diffSheet.write (rowCnt, 34, paramStruct2.LOG, xlsxFormatChanged)
            diffSheet.write (rowCnt, 35, paramStruct2.DISPLAY, xlsxFormatChanged)
            diffSheet.write (rowCnt, 36, paramStruct2.STATISTICS, xlsxFormatChanged)
            diffSheet.write (rowCnt, 37, paramStruct2.RANGE, xlsxFormatChanged)
            diffSheet.write (rowCnt, 38, paramStruct2.ERROR, xlsxFormatChanged)
            
            addedParameters += 1
            rowCnt += 1
            
    # Report changed parameters
    for meas1, paramStruct1 in sortedParamDict1.items ():
        if meas1 in sortedParamDict2:
            paramStruct2 = sortedParamDict2[meas1]
           
            paramChanged = False;
            paramNumChanged = False if 'NUM' in args.ignore else paramStruct1.NUM != paramStruct2.NUM
            paramChanged = paramChanged or paramNumChanged
            paramPromptChanged = False if 'PROMPT' in args.ignore else paramStruct1.PROMPT != paramStruct2.PROMPT
            paramChanged = paramChanged or paramPromptChanged
            paramUnitChanged = False if 'UNIT' in args.ignore else paramStruct1.UNIT != paramStruct1.UNIT
            paramChanged = paramChanged or paramUnitChanged
            paramMeasTypeChanged = False if 'MEAS_TYPE' in args.ignore else paramStruct1.MEAS_TYPE != paramStruct2.MEAS_TYPE
            paramChanged = paramChanged or paramMeasTypeChanged
            paramBinChanged = False if 'BIN' in args.ignore else paramStruct1.BIN != paramStruct2.BIN
            paramChanged = paramChanged or paramBinChanged
            paramResChanged = False if 'RES' in args.ignore else paramStruct1.RES != paramStruct2.RES
            paramChanged = paramChanged or paramResChanged
            paramMinChanged = False if 'MIN' in args.ignore else paramStruct1.MIN != paramStruct2.MIN
            paramChanged = paramChanged or paramMinChanged
            paramMinValueChanged = False if 'MIN_VALUE' in args.ignore else paramStruct1.MIN_VALUE != paramStruct2.MIN_VALUE
            paramChanged = paramChanged or paramMinValueChanged
            paramMaxChanged = False if 'MAX' in args.ignore else paramStruct1.MAX != paramStruct2.MAX
            paramChanged = paramChanged or paramMaxChanged
            paramMaxValueChanged = False if 'MAX_VALUE' in args.ignore else paramStruct1.MAX_VALUE != paramStruct2.MAX_VALUE
            paramChanged = paramChanged or paramMaxValueChanged
            paramNomChanged = False if 'NOM' in args.ignore else paramStruct1.NOM != paramStruct2.NOM
            paramChanged = paramChanged or paramNomChanged
            paramNomValueChanged = False if 'NOM_VALUE' in args.ignore else paramStruct1.NOM_VALUE != paramStruct2.NOM_VALUE
            paramChanged = paramChanged or paramNomValueChanged
            paramArraySizeChanged = False if 'ARRAY_SIZE' in args.ignore else paramStruct1.ARRAY_SIZE != paramStruct2.ARRAY_SIZE
            paramChanged = paramChanged or paramArraySizeChanged
            paramLogChanged = False if 'LOG' in args.ignore else paramStruct1.LOG != paramStruct2.LOG
            paramChanged = paramChanged or paramLogChanged
            paramDisplayChanged = False if 'DISPLAY' in args.ignore else paramStruct1.DISPLAY != paramStruct2.DISPLAY
            paramChanged = paramChanged or paramDisplayChanged
            paramStatisticsChanged = False if 'STATISTICS' in args.ignore else paramStruct1.STATISTICS != paramStruct2.STATISTICS
            paramChanged = paramChanged or paramStatisticsChanged
            paramRangeChanged = False if 'RANGE' in args.ignore else paramStruct1.RANGE != paramStruct2.RANGE
            paramChanged = paramChanged or paramRangeChanged
            paramErrorChanged = False if 'ERROR' in args.ignore else paramStruct1.ERROR != paramStruct2.ERROR
            paramChanged = paramChanged or paramErrorChanged

            if paramChanged:
                diffSheet.write (rowCnt, 0, 'Changed')
                diffSheet.write (rowCnt, 1, meas1)
                diffSheet.write (rowCnt, 2, paramStruct1.NUM, xlsxFormatChanged if paramNumChanged else None)
                diffSheet.write (rowCnt, 3, paramStruct1.PROMPT, xlsxFormatChanged if paramPromptChanged else None)
                diffSheet.write (rowCnt, 4, paramStruct1.UNIT, xlsxFormatChanged if paramUnitChanged else None)
                diffSheet.write (rowCnt, 5, paramStruct1.MEAS_TYPE, xlsxFormatChanged if paramMeasTypeChanged else None)
                diffSheet.write (rowCnt, 6, paramStruct1.BIN, xlsxFormatChanged if paramBinChanged else None)
                diffSheet.write (rowCnt, 7, paramStruct1.RES, xlsxFormatChanged if paramResChanged else None)
                diffSheet.write (rowCnt, 8, paramStruct1.MIN, xlsxFormatChanged if paramMinChanged else None)
                diffSheet.write (rowCnt, 9, paramStruct1.MIN_VALUE, xlsxFormatChanged if paramMinValueChanged else None)
                diffSheet.write (rowCnt, 10, paramStruct1.MAX, xlsxFormatChanged if paramMaxChanged else None)
                diffSheet.write (rowCnt, 11, paramStruct1.MAX_VALUE, xlsxFormatChanged if paramMaxValueChanged else None)
                diffSheet.write (rowCnt, 12, paramStruct1.NOM, xlsxFormatChanged if paramNomChanged else None)
                diffSheet.write (rowCnt, 13, paramStruct1.NOM_VALUE, xlsxFormatChanged if paramNomValueChanged else None)
                diffSheet.write (rowCnt, 14, paramStruct1.ARRAY_SIZE, xlsxFormatChanged if paramArraySizeChanged else None)
                diffSheet.write (rowCnt, 15, paramStruct1.LOG, xlsxFormatChanged if paramLogChanged else None)
                diffSheet.write (rowCnt, 16, paramStruct1.DISPLAY, xlsxFormatChanged if paramDisplayChanged else None)
                diffSheet.write (rowCnt, 17, paramStruct1.STATISTICS, xlsxFormatChanged if paramStatisticsChanged else None)
                diffSheet.write (rowCnt, 18, paramStruct1.RANGE, xlsxFormatChanged if paramRangeChanged else None)
                diffSheet.write (rowCnt, 19, paramStruct1.ERROR, xlsxFormatChanged if paramErrorChanged else None)
                diffSheet.write (rowCnt, 20, meas1)
                diffSheet.write (rowCnt, 21, paramStruct2.NUM, xlsxFormatChanged if paramNumChanged else None)
                diffSheet.write (rowCnt, 22, paramStruct2.PROMPT, xlsxFormatChanged if paramPromptChanged else None)
                diffSheet.write (rowCnt, 23, paramStruct2.UNIT, xlsxFormatChanged if paramUnitChanged else None)
                diffSheet.write (rowCnt, 24, paramStruct2.MEAS_TYPE, xlsxFormatChanged if paramMeasTypeChanged else None)
                diffSheet.write (rowCnt, 25, paramStruct2.BIN, xlsxFormatChanged if paramBinChanged else None)
                diffSheet.write (rowCnt, 26, paramStruct2.RES, xlsxFormatChanged if paramResChanged else None)
                diffSheet.write (rowCnt, 27, paramStruct2.MIN, xlsxFormatChanged if paramMinChanged else None)
                diffSheet.write (rowCnt, 28, paramStruct2.MIN_VALUE, xlsxFormatChanged if paramMinValueChanged else None)
                diffSheet.write (rowCnt, 29, paramStruct2.MAX, xlsxFormatChanged if paramMaxChanged else None)
                diffSheet.write (rowCnt, 30, paramStruct2.MAX_VALUE, xlsxFormatChanged if paramMaxValueChanged else None)
                diffSheet.write (rowCnt, 31, paramStruct2.NOM, xlsxFormatChanged if paramNomChanged else None)
                diffSheet.write (rowCnt, 32, paramStruct2.NOM_VALUE, xlsxFormatChanged if paramNomValueChanged else None)
                diffSheet.write (rowCnt, 33, paramStruct2.ARRAY_SIZE, xlsxFormatChanged if paramArraySizeChanged else None)
                diffSheet.write (rowCnt, 34, paramStruct2.LOG, xlsxFormatChanged if paramLogChanged else None)
                diffSheet.write (rowCnt, 35, paramStruct2.DISPLAY, xlsxFormatChanged if paramDisplayChanged else None)
                diffSheet.write (rowCnt, 36, paramStruct2.STATISTICS, xlsxFormatChanged if paramStatisticsChanged else None)
                diffSheet.write (rowCnt, 37, paramStruct2.RANGE, xlsxFormatChanged if paramRangeChanged else None)
                diffSheet.write (rowCnt, 38, paramStruct2.ERROR, xlsxFormatChanged if paramErrorChanged else None)

                changedParameters += 1
                rowCnt += 1
                
    # Report not changed parameters
    for meas1, paramStruct1 in sortedParamDict1.items ():
        if meas1 in sortedParamDict2:
            paramStruct2 = sortedParamDict2[meas1]

            for field, value in paramStruct1._asdict ().items ():
                if field in args.ignore:
                    print ('field:' + str (field) + ',value:' + str (value))
            
            paramChanged = False;
            paramChanged = paramChanged or (False if 'NUM' in args.ignore else paramStruct1.NUM != paramStruct2.NUM)
            paramChanged = paramChanged or (False if 'PROMPT' in args.ignore else paramStruct1.PROMPT != paramStruct2.PROMPT)
            paramChanged = paramChanged or (False if 'UNIT' in args.ignore else paramStruct1.UNIT != paramStruct2.UNIT)
            paramChanged = paramChanged or (False if 'MEAS_TYPE' in args.ignore else paramStruct1.MEAS_TYPE != paramStruct2.MEAS_TYPE)
            paramChanged = paramChanged or (False if 'BIN' in args.ignore else paramStruct1.BIN != paramStruct2.BIN)
            paramChanged = paramChanged or (False if 'RES' in args.ignore else paramStruct1.RES != paramStruct2.RES)
            paramChanged = paramChanged or (False if 'MIN' in args.ignore else paramStruct1.MIN != paramStruct2.MIN)
            paramChanged = paramChanged or (False if 'MIN_VALUE' in args.ignore else paramStruct1.MIN_VALUE != paramStruct2.MIN_VALUE)
            paramChanged = paramChanged or (False if 'MAX' in args.ignore else paramStruct1.MAX != paramStruct2.MAX)
            paramChanged = paramChanged or (False if 'MAX_VALUE' in args.ignore else paramStruct1.MAX_VALUE != paramStruct2.MAX_VALUE)
            paramChanged = paramChanged or (False if 'NOM' in args.ignore else paramStruct1.NOM != paramStruct2.NOM)
            paramChanged = paramChanged or (False if 'NOM_VALUE' in args.ignore else paramStruct1.NOM_VALUE != paramStruct2.NOM_VALUE)
            paramChanged = paramChanged or (False if 'ARRAY_SIZE' in args.ignore else paramStruct1.ARRAY_SIZE != paramStruct2.ARRAY_SIZE)
            paramChanged = paramChanged or (False if 'LOG' in args.ignore else paramStruct1.LOG != paramStruct2.LOG)
            paramChanged = paramChanged or (False if 'DISPLAY' in args.ignore else paramStruct1.DISPLAY != paramStruct2.DISPLAY)
            paramChanged = paramChanged or (False if 'STATISTICS' in args.ignore else paramStruct1.STATISTICS != paramStruct2.STATISTICS)
            paramChanged = paramChanged or (False if 'RANGE' in args.ignore else paramStruct1.RANGE != paramStruct2.RANGE)
            paramChanged = paramChanged or (False if 'ERROR' in args.ignore else paramStruct1.ERROR != paramStruct2.ERROR)

            if not paramChanged:
                diffSheet.write (rowCnt, 0, 'Not changed')
                diffSheet.write (rowCnt, 1, meas1)
                diffSheet.write (rowCnt, 2, paramStruct1.NUM)
                diffSheet.write (rowCnt, 3, paramStruct1.PROMPT)
                diffSheet.write (rowCnt, 4, paramStruct1.UNIT)
                diffSheet.write (rowCnt, 5, paramStruct1.MEAS_TYPE)
                diffSheet.write (rowCnt, 6, paramStruct1.BIN)
                diffSheet.write (rowCnt, 7, paramStruct1.RES)
                diffSheet.write (rowCnt, 8, paramStruct1.MIN)
                diffSheet.write (rowCnt, 9, paramStruct1.MIN_VALUE)
                diffSheet.write (rowCnt, 10, paramStruct1.MAX)
                diffSheet.write (rowCnt, 11, paramStruct1.MAX_VALUE)
                diffSheet.write (rowCnt, 12, paramStruct1.NOM)
                diffSheet.write (rowCnt, 13, paramStruct1.NOM_VALUE)
                diffSheet.write (rowCnt, 14, paramStruct1.ARRAY_SIZE)
                diffSheet.write (rowCnt, 15, paramStruct1.LOG)
                diffSheet.write (rowCnt, 16, paramStruct1.DISPLAY)
                diffSheet.write (rowCnt, 17, paramStruct1.STATISTICS)
                diffSheet.write (rowCnt, 18, paramStruct1.RANGE)
                diffSheet.write (rowCnt, 19, paramStruct1.ERROR)
                diffSheet.write (rowCnt, 20, meas2)
                diffSheet.write (rowCnt, 21, paramStruct2.NUM)
                diffSheet.write (rowCnt, 22, paramStruct2.PROMPT)
                diffSheet.write (rowCnt, 23, paramStruct2.UNIT)
                diffSheet.write (rowCnt, 24, paramStruct2.MEAS_TYPE)
                diffSheet.write (rowCnt, 25, paramStruct2.BIN)
                diffSheet.write (rowCnt, 26, paramStruct2.RES)
                diffSheet.write (rowCnt, 27, paramStruct2.MIN)
                diffSheet.write (rowCnt, 28, paramStruct2.MIN_VALUE)
                diffSheet.write (rowCnt, 29, paramStruct2.MAX)
                diffSheet.write (rowCnt, 30, paramStruct2.MAX_VALUE)
                diffSheet.write (rowCnt, 31, paramStruct2.NOM)
                diffSheet.write (rowCnt, 32, paramStruct2.NOM_VALUE)
                diffSheet.write (rowCnt, 33, paramStruct2.ARRAY_SIZE)
                diffSheet.write (rowCnt, 34, paramStruct2.LOG)
                diffSheet.write (rowCnt, 35, paramStruct2.DISPLAY)
                diffSheet.write (rowCnt, 36, paramStruct2.STATISTICS)
                diffSheet.write (rowCnt, 37, paramStruct2.RANGE)
                diffSheet.write (rowCnt, 38, paramStruct2.ERROR)

                notChangedParameters += 1
                rowCnt += 1
                
    diffSheet.write (rowParamRemoved, 1, removedParameters)
    diffSheet.write (rowParamAdded, 1, addedParameters)
    diffSheet.write (rowParamChanged, 1, changedParameters)
    diffSheet.write (rowParamNotChanged, 1, notChangedParameters)

    #*** Bincodes ***
    rowCnt += 1
    # Overview
    diffSheet.write (rowCnt, 0, 'Bincodes removed ')
    rowBinRemoved = rowCnt
    rowCnt += 1
    diffSheet.write (rowCnt, 0, 'Bincodes added')
    rowBinAdded = rowCnt
    rowCnt += 1
    diffSheet.write (rowCnt, 0, 'Bincodes changed')
    rowBinChanged = rowCnt
    rowCnt += 1
    diffSheet.write (rowCnt, 0, 'Bincodes not changed')
    rowBinNotChanged = rowCnt
    rowCnt += 1

    # Add headers
    diffSheet.write (rowCnt, 1, nameFile1)
    diffSheet.write (rowCnt, 4, nameFile2)
    rowCnt += 1
    diffSheet.write (rowCnt, 0, 'Compare result')
    diffSheet.write (rowCnt, 1, 'Bincode')
    diffSheet.write (rowCnt, 2, 'Bin type')
    diffSheet.write (rowCnt, 3, 'Bin description')
    diffSheet.write (rowCnt, 4, 'Bincode')
    diffSheet.write (rowCnt, 5, 'Bin type')
    diffSheet.write (rowCnt, 6, 'Bin description')
    rowCnt += 1
    
    # Report removed bincodes from DMLF file 2
    addedBincodes = 0
    changedBincodes = 0
    notChangedBincodes = 0
    removedBincodes = 0
        
    for bin1, binStruct1 in binDict1.items ():
        if bin1 not in binDict2:
            diffSheet.write (rowCnt, 0, 'Removed')
            diffSheet.write (rowCnt, 1, bin1, xlsxFormatChanged)
            diffSheet.write (rowCnt, 2, binStruct1.BinType, xlsxFormatChanged)
            diffSheet.write (rowCnt, 3, binStruct1.BinDesc, xlsxFormatChanged)

            removedBincodes += 1
            rowCnt += 1
            
    # Report added bincodes in DMLF file 2
    for bin2, binStruct2 in binDict2.items ():
        if bin2 not in binDict1:
            diffSheet.write (rowCnt, 0, 'Added')
            diffSheet.write (rowCnt, 4, bin2, xlsxFormatChanged)
            diffSheet.write (rowCnt, 5, binStruct2.BinType, xlsxFormatChanged)
            diffSheet.write (rowCnt, 6, binStruct2.BinDesc, xlsxFormatChanged)
            
            addedBincodes += 1
            rowCnt += 1
            
    # Report changed bincodes
    for bin1, binStruct1 in binDict1.items ():
        if bin1 in binDict2:
            binStruct2 = binDict2[bin1]
            binTypeChanged = binStruct1.BinType != binStruct2.BinType
            binDescChanged = binStruct1.BinDesc != binStruct2.BinDesc
            if (binTypeChanged or binDescChanged):
                diffSheet.write (rowCnt, 0, 'Changed')
                diffSheet.write (rowCnt, 1, bin1)
                diffSheet.write (rowCnt, 4, bin1)
                if binTypeChanged:
                    diffSheet.write (rowCnt, 2, binStruct1.BinType, xlsxFormatChanged)
                    diffSheet.write (rowCnt, 5, binStruct2.BinType, xlsxFormatChanged)
                else:
                    diffSheet.write (rowCnt, 2, binStruct1.BinType)
                    diffSheet.write (rowCnt, 5, binStruct2.BinType)
                if binDescChanged:
                    diffSheet.write (rowCnt, 3, binStruct1.BinDesc, xlsxFormatChanged)
                    diffSheet.write (rowCnt, 6, binStruct2.BinDesc, xlsxFormatChanged)
                else:
                    diffSheet.write (rowCnt, 3, binStruct1.BinDesc)
                    diffSheet.write (rowCnt, 6, binStruct2.BinDesc)

                changedBincodes += 1
                rowCnt += 1
                
    # Report not changed bincodes
    for bin1, binStruct1 in binDict1.items ():
        if bin1 in binDict2:
            binStruct2 = binDict2[bin1]
            if ((binStruct1.BinType == binStruct2.BinType) and (binStruct1.BinDesc == binStruct2.BinDesc)):
                diffSheet.write (rowCnt, 0, 'Not changed')
                diffSheet.write (rowCnt, 1, bin1)
                diffSheet.write (rowCnt, 2, binStruct1.BinType)
                diffSheet.write (rowCnt, 3, binStruct1.BinDesc)
                diffSheet.write (rowCnt, 4, bin1)
                diffSheet.write (rowCnt, 5, binStruct2.BinType)
                diffSheet.write (rowCnt, 6, binStruct2.BinDesc)
                
                notChangedBincodes += 1
                rowCnt += 1

    diffSheet.write (rowBinRemoved, 1, removedBincodes)
    diffSheet.write (rowBinAdded, 1, addedBincodes)
    diffSheet.write (rowBinChanged, 1, changedBincodes)
    diffSheet.write (rowBinNotChanged, 1, notChangedBincodes)

def doDiffWithTranslation (diffSheet, nameFile1, nameFile2, binDict1, paramDict1, binDict2, paramDict2, xlsxFormatChanged, xlsxFormatRenamed, transTable):
    global args
    
    rowCnt = 0
    
    sortedParamDict1 = OrderedDict (sorted (paramDict1.items (), key = lambda x: x[1].NUM))
    sortedParamDict2 = OrderedDict (sorted (paramDict2.items (), key = lambda x: x[1].NUM))

    #*** Parameters ***
    # Overview
    diffSheet.write (rowCnt, 0, 'Parameters removed ')
    rowParamRemoved = rowCnt
    rowCnt += 1
    diffSheet.write (rowCnt, 0, 'Parameters added')
    rowParamAdded = rowCnt
    rowCnt += 1
    diffSheet.write (rowCnt, 0, 'Parameters changed')
    rowParamChanged = rowCnt
    rowCnt += 1
    diffSheet.write (rowCnt, 0, 'Parameters not changed')
    rowParamNotChanged = rowCnt
    rowCnt += 1
    diffSheet.write (rowCnt, 0, 'Parameters renamed')
    rowParamRenamed = rowCnt
    rowCnt += 1

    # Add headers
    diffSheet.write (rowCnt, 1, nameFile1)
    diffSheet.write (rowCnt, 20, nameFile2)
    rowCnt += 1
    diffSheet.write (rowCnt, 0, 'Compare result')
    diffSheet.write (rowCnt, 1, 'Meas')
    diffSheet.write (rowCnt, 2, 'Num')
    diffSheet.write (rowCnt, 3, 'Prompt')
    diffSheet.write (rowCnt, 4, 'Unit')
    diffSheet.write (rowCnt, 5, 'Meas type')
    diffSheet.write (rowCnt, 6, 'Bin')
    diffSheet.write (rowCnt, 7, 'Res')
    diffSheet.write (rowCnt, 8, 'Min')
    diffSheet.write (rowCnt, 9, 'Min value')
    diffSheet.write (rowCnt, 10, 'Max')
    diffSheet.write (rowCnt, 11, 'Max value')
    diffSheet.write (rowCnt, 12, 'Nom')
    diffSheet.write (rowCnt, 13, 'Nom value')
    diffSheet.write (rowCnt, 14, 'Array size')
    diffSheet.write (rowCnt, 15, 'Log')
    diffSheet.write (rowCnt, 16, 'Display')
    diffSheet.write (rowCnt, 17, 'Statistics')
    diffSheet.write (rowCnt, 18, 'Range')
    diffSheet.write (rowCnt, 19, 'Error')
    diffSheet.write (rowCnt, 20, 'Meas')
    diffSheet.write (rowCnt, 21, 'Num')
    diffSheet.write (rowCnt, 22, 'Prompt')
    diffSheet.write (rowCnt, 23, 'Unit')
    diffSheet.write (rowCnt, 24, 'Meas type')
    diffSheet.write (rowCnt, 25, 'Bin')
    diffSheet.write (rowCnt, 26, 'Res')
    diffSheet.write (rowCnt, 27, 'Min')
    diffSheet.write (rowCnt, 28, 'Min value')
    diffSheet.write (rowCnt, 29, 'Max')
    diffSheet.write (rowCnt, 30, 'Max value')
    diffSheet.write (rowCnt, 31, 'Nom')
    diffSheet.write (rowCnt, 32, 'Nom value')
    diffSheet.write (rowCnt, 33, 'Array size')
    diffSheet.write (rowCnt, 34, 'Log')
    diffSheet.write (rowCnt, 35, 'Display')
    diffSheet.write (rowCnt, 36, 'Statistics')
    diffSheet.write (rowCnt, 37, 'Range')
    diffSheet.write (rowCnt, 38, 'Error')
    rowCnt += 1
    
    addedParameters = 0
    changedParameters = 0
    notChangedParameters = 0
    removedParameters = 0

    # Report removed parameters
    for meas1, paramStruct1 in sortedParamDict1.items ():
        if meas1 not in sortedParamDict2:
            if meas1 not in transTable: # Check if param name is not in translated table
                diffSheet.write (rowCnt, 0, 'Removed')
                diffSheet.write (rowCnt, 1, meas1, xlsxFormatChanged)
                diffSheet.write (rowCnt, 2, paramStruct1.NUM, xlsxFormatChanged)
                diffSheet.write (rowCnt, 3, paramStruct1.PROMPT, xlsxFormatChanged)
                diffSheet.write (rowCnt, 4, paramStruct1.UNIT, xlsxFormatChanged)
                diffSheet.write (rowCnt, 5, paramStruct1.MEAS_TYPE, xlsxFormatChanged)
                diffSheet.write (rowCnt, 6, paramStruct1.BIN, xlsxFormatChanged)
                diffSheet.write (rowCnt, 7, paramStruct1.RES, xlsxFormatChanged)
                diffSheet.write (rowCnt, 8, paramStruct1.MIN, xlsxFormatChanged)
                diffSheet.write (rowCnt, 9, paramStruct1.MIN_VALUE, xlsxFormatChanged)
                diffSheet.write (rowCnt, 10, paramStruct1.MAX, xlsxFormatChanged)
                diffSheet.write (rowCnt, 11, paramStruct1.MAX_VALUE, xlsxFormatChanged)
                diffSheet.write (rowCnt, 12, paramStruct1.NOM, xlsxFormatChanged)
                diffSheet.write (rowCnt, 13, paramStruct1.NOM_VALUE, xlsxFormatChanged)
                diffSheet.write (rowCnt, 14, paramStruct1.ARRAY_SIZE, xlsxFormatChanged)
                diffSheet.write (rowCnt, 15, paramStruct1.LOG, xlsxFormatChanged)
                diffSheet.write (rowCnt, 16, paramStruct1.DISPLAY, xlsxFormatChanged)
                diffSheet.write (rowCnt, 17, paramStruct1.STATISTICS, xlsxFormatChanged)
                diffSheet.write (rowCnt, 18, paramStruct1.RANGE, xlsxFormatChanged)
                diffSheet.write (rowCnt, 19, paramStruct1.ERROR, xlsxFormatChanged)

                removedParameters += 1
                rowCnt += 1
            
    # Report added parameters
    for meas2, paramStruct2 in sortedParamDict2.items ():
        if meas2 not in sortedParamDict1:
            if meas2 not in transTable.viewvalues(): # Check if param name is not in translated table
                diffSheet.write (rowCnt, 0, 'Added')
                diffSheet.write (rowCnt, 20, meas2, xlsxFormatChanged)
                diffSheet.write (rowCnt, 21, paramStruct2.NUM, xlsxFormatChanged)
                diffSheet.write (rowCnt, 22, paramStruct2.PROMPT, xlsxFormatChanged)
                diffSheet.write (rowCnt, 23, paramStruct2.UNIT, xlsxFormatChanged)
                diffSheet.write (rowCnt, 24, paramStruct2.MEAS_TYPE, xlsxFormatChanged)
                diffSheet.write (rowCnt, 25, paramStruct2.BIN, xlsxFormatChanged)
                diffSheet.write (rowCnt, 26, paramStruct2.RES, xlsxFormatChanged)
                diffSheet.write (rowCnt, 27, paramStruct2.MIN, xlsxFormatChanged)
                diffSheet.write (rowCnt, 28, paramStruct2.MIN_VALUE, xlsxFormatChanged)
                diffSheet.write (rowCnt, 29, paramStruct2.MAX, xlsxFormatChanged)
                diffSheet.write (rowCnt, 30, paramStruct2.MAX_VALUE, xlsxFormatChanged)
                diffSheet.write (rowCnt, 31, paramStruct2.NOM, xlsxFormatChanged)
                diffSheet.write (rowCnt, 32, paramStruct2.NOM_VALUE, xlsxFormatChanged)
                diffSheet.write (rowCnt, 33, paramStruct2.ARRAY_SIZE, xlsxFormatChanged)
                diffSheet.write (rowCnt, 34, paramStruct2.LOG, xlsxFormatChanged)
                diffSheet.write (rowCnt, 35, paramStruct2.DISPLAY, xlsxFormatChanged)
                diffSheet.write (rowCnt, 36, paramStruct2.STATISTICS, xlsxFormatChanged)
                diffSheet.write (rowCnt, 37, paramStruct2.RANGE, xlsxFormatChanged)
                diffSheet.write (rowCnt, 38, paramStruct2.ERROR, xlsxFormatChanged)
                
                addedParameters += 1
                rowCnt += 1
    
    # Report changed parameters
    for meas1, paramStruct1 in sortedParamDict1.items ():
        meas1_trans = meas1
        if meas1 in transTable:
            meas1_trans = transTable[meas1]
    
        if meas1_trans in sortedParamDict2:
            paramStruct2 = sortedParamDict2[meas1_trans]
           
            paramChanged = False;
            paramNumChanged = False if 'NUM' in args.ignore else paramStruct1.NUM != paramStruct2.NUM
            paramChanged = paramChanged or paramNumChanged
            paramPromptChanged = False if 'PROMPT' in args.ignore else paramStruct1.PROMPT != paramStruct2.PROMPT
            paramChanged = paramChanged or paramPromptChanged
            paramUnitChanged = False if 'UNIT' in args.ignore else paramStruct1.UNIT != paramStruct1.UNIT
            paramChanged = paramChanged or paramUnitChanged
            paramMeasTypeChanged = False if 'MEAS_TYPE' in args.ignore else paramStruct1.MEAS_TYPE != paramStruct2.MEAS_TYPE
            paramChanged = paramChanged or paramMeasTypeChanged
            paramBinChanged = False if 'BIN' in args.ignore else paramStruct1.BIN != paramStruct2.BIN
            paramChanged = paramChanged or paramBinChanged
            paramResChanged = False if 'RES' in args.ignore else paramStruct1.RES != paramStruct2.RES
            paramChanged = paramChanged or paramResChanged
            paramMinChanged = False if 'MIN' in args.ignore else paramStruct1.MIN != paramStruct2.MIN
            paramChanged = paramChanged or paramMinChanged
            paramMinValueChanged = False if 'MIN_VALUE' in args.ignore else paramStruct1.MIN_VALUE != paramStruct2.MIN_VALUE
            paramChanged = paramChanged or paramMinValueChanged
            paramMaxChanged = False if 'MAX' in args.ignore else paramStruct1.MAX != paramStruct2.MAX
            paramChanged = paramChanged or paramMaxChanged
            paramMaxValueChanged = False if 'MAX_VALUE' in args.ignore else paramStruct1.MAX_VALUE != paramStruct2.MAX_VALUE
            paramChanged = paramChanged or paramMaxValueChanged
            paramNomChanged = False if 'NOM' in args.ignore else paramStruct1.NOM != paramStruct2.NOM
            paramChanged = paramChanged or paramNomChanged
            paramNomValueChanged = False if 'NOM_VALUE' in args.ignore else paramStruct1.NOM_VALUE != paramStruct2.NOM_VALUE
            paramChanged = paramChanged or paramNomValueChanged
            paramArraySizeChanged = False if 'ARRAY_SIZE' in args.ignore else paramStruct1.ARRAY_SIZE != paramStruct2.ARRAY_SIZE
            paramChanged = paramChanged or paramArraySizeChanged
            paramLogChanged = False if 'LOG' in args.ignore else paramStruct1.LOG != paramStruct2.LOG
            paramChanged = paramChanged or paramLogChanged
            paramDisplayChanged = False if 'DISPLAY' in args.ignore else paramStruct1.DISPLAY != paramStruct2.DISPLAY
            paramChanged = paramChanged or paramDisplayChanged
            paramStatisticsChanged = False if 'STATISTICS' in args.ignore else paramStruct1.STATISTICS != paramStruct2.STATISTICS
            paramChanged = paramChanged or paramStatisticsChanged
            paramRangeChanged = False if 'RANGE' in args.ignore else paramStruct1.RANGE != paramStruct2.RANGE
            paramChanged = paramChanged or paramRangeChanged
            paramErrorChanged = False if 'ERROR' in args.ignore else paramStruct1.ERROR != paramStruct2.ERROR
            paramChanged = paramChanged or paramErrorChanged

            if paramChanged:
                diffSheet.write (rowCnt, 0, 'Changed')

                if meas1 in transTable:                    
                    diffSheet.write (rowCnt, 1, meas1, xlsxFormatRenamed)
                else:
                    diffSheet.write (rowCnt, 1, meas1)
                    
                diffSheet.write (rowCnt, 2, paramStruct1.NUM, xlsxFormatChanged if paramNumChanged else None)
                diffSheet.write (rowCnt, 3, paramStruct1.PROMPT, xlsxFormatChanged if paramPromptChanged else None)
                diffSheet.write (rowCnt, 4, paramStruct1.UNIT, xlsxFormatChanged if paramUnitChanged else None)
                diffSheet.write (rowCnt, 5, paramStruct1.MEAS_TYPE, xlsxFormatChanged if paramMeasTypeChanged else None)
                diffSheet.write (rowCnt, 6, paramStruct1.BIN, xlsxFormatChanged if paramBinChanged else None)
                diffSheet.write (rowCnt, 7, paramStruct1.RES, xlsxFormatChanged if paramResChanged else None)
                diffSheet.write (rowCnt, 8, paramStruct1.MIN, xlsxFormatChanged if paramMinChanged else None)
                diffSheet.write (rowCnt, 9, paramStruct1.MIN_VALUE, xlsxFormatChanged if paramMinValueChanged else None)
                diffSheet.write (rowCnt, 10, paramStruct1.MAX, xlsxFormatChanged if paramMaxChanged else None)
                diffSheet.write (rowCnt, 11, paramStruct1.MAX_VALUE, xlsxFormatChanged if paramMaxValueChanged else None)
                diffSheet.write (rowCnt, 12, paramStruct1.NOM, xlsxFormatChanged if paramNomChanged else None)
                diffSheet.write (rowCnt, 13, paramStruct1.NOM_VALUE, xlsxFormatChanged if paramNomValueChanged else None)
                diffSheet.write (rowCnt, 14, paramStruct1.ARRAY_SIZE, xlsxFormatChanged if paramArraySizeChanged else None)
                diffSheet.write (rowCnt, 15, paramStruct1.LOG, xlsxFormatChanged if paramLogChanged else None)
                diffSheet.write (rowCnt, 16, paramStruct1.DISPLAY, xlsxFormatChanged if paramDisplayChanged else None)
                diffSheet.write (rowCnt, 17, paramStruct1.STATISTICS, xlsxFormatChanged if paramStatisticsChanged else None)
                diffSheet.write (rowCnt, 18, paramStruct1.RANGE, xlsxFormatChanged if paramRangeChanged else None)
                diffSheet.write (rowCnt, 19, paramStruct1.ERROR, xlsxFormatChanged if paramErrorChanged else None)
                
                if meas1 in transTable:                    
                    diffSheet.write (rowCnt, 20, meas1_trans, xlsxFormatRenamed)
                else:
                    diffSheet.write (rowCnt, 20, meas1)
                
                diffSheet.write (rowCnt, 21, paramStruct2.NUM, xlsxFormatChanged if paramNumChanged else None)
                diffSheet.write (rowCnt, 22, paramStruct2.PROMPT, xlsxFormatChanged if paramPromptChanged else None)
                diffSheet.write (rowCnt, 23, paramStruct2.UNIT, xlsxFormatChanged if paramUnitChanged else None)
                diffSheet.write (rowCnt, 24, paramStruct2.MEAS_TYPE, xlsxFormatChanged if paramMeasTypeChanged else None)
                diffSheet.write (rowCnt, 25, paramStruct2.BIN, xlsxFormatChanged if paramBinChanged else None)
                diffSheet.write (rowCnt, 26, paramStruct2.RES, xlsxFormatChanged if paramResChanged else None)
                diffSheet.write (rowCnt, 27, paramStruct2.MIN, xlsxFormatChanged if paramMinChanged else None)
                diffSheet.write (rowCnt, 28, paramStruct2.MIN_VALUE, xlsxFormatChanged if paramMinValueChanged else None)
                diffSheet.write (rowCnt, 29, paramStruct2.MAX, xlsxFormatChanged if paramMaxChanged else None)
                diffSheet.write (rowCnt, 30, paramStruct2.MAX_VALUE, xlsxFormatChanged if paramMaxValueChanged else None)
                diffSheet.write (rowCnt, 31, paramStruct2.NOM, xlsxFormatChanged if paramNomChanged else None)
                diffSheet.write (rowCnt, 32, paramStruct2.NOM_VALUE, xlsxFormatChanged if paramNomValueChanged else None)
                diffSheet.write (rowCnt, 33, paramStruct2.ARRAY_SIZE, xlsxFormatChanged if paramArraySizeChanged else None)
                diffSheet.write (rowCnt, 34, paramStruct2.LOG, xlsxFormatChanged if paramLogChanged else None)
                diffSheet.write (rowCnt, 35, paramStruct2.DISPLAY, xlsxFormatChanged if paramDisplayChanged else None)
                diffSheet.write (rowCnt, 36, paramStruct2.STATISTICS, xlsxFormatChanged if paramStatisticsChanged else None)
                diffSheet.write (rowCnt, 37, paramStruct2.RANGE, xlsxFormatChanged if paramRangeChanged else None)
                diffSheet.write (rowCnt, 38, paramStruct2.ERROR, xlsxFormatChanged if paramErrorChanged else None)

                changedParameters += 1
                rowCnt += 1
                
    # Report not changed parameters
    for meas1, paramStruct1 in sortedParamDict1.items ():
        if meas1 in sortedParamDict2:
            paramStruct2 = sortedParamDict2[meas1]

            for field, value in paramStruct1._asdict ().items ():
                if field in args.ignore:
                    print ('field:' + str (field) + ',value:' + str (value))
            
            paramChanged = False;
            paramChanged = paramChanged or (False if 'NUM' in args.ignore else paramStruct1.NUM != paramStruct2.NUM)
            paramChanged = paramChanged or (False if 'PROMPT' in args.ignore else paramStruct1.PROMPT != paramStruct2.PROMPT)
            paramChanged = paramChanged or (False if 'UNIT' in args.ignore else paramStruct1.UNIT != paramStruct2.UNIT)
            paramChanged = paramChanged or (False if 'MEAS_TYPE' in args.ignore else paramStruct1.MEAS_TYPE != paramStruct2.MEAS_TYPE)
            paramChanged = paramChanged or (False if 'BIN' in args.ignore else paramStruct1.BIN != paramStruct2.BIN)
            paramChanged = paramChanged or (False if 'RES' in args.ignore else paramStruct1.RES != paramStruct2.RES)
            paramChanged = paramChanged or (False if 'MIN' in args.ignore else paramStruct1.MIN != paramStruct2.MIN)
            paramChanged = paramChanged or (False if 'MIN_VALUE' in args.ignore else paramStruct1.MIN_VALUE != paramStruct2.MIN_VALUE)
            paramChanged = paramChanged or (False if 'MAX' in args.ignore else paramStruct1.MAX != paramStruct2.MAX)
            paramChanged = paramChanged or (False if 'MAX_VALUE' in args.ignore else paramStruct1.MAX_VALUE != paramStruct2.MAX_VALUE)
            paramChanged = paramChanged or (False if 'NOM' in args.ignore else paramStruct1.NOM != paramStruct2.NOM)
            paramChanged = paramChanged or (False if 'NOM_VALUE' in args.ignore else paramStruct1.NOM_VALUE != paramStruct2.NOM_VALUE)
            paramChanged = paramChanged or (False if 'ARRAY_SIZE' in args.ignore else paramStruct1.ARRAY_SIZE != paramStruct2.ARRAY_SIZE)
            paramChanged = paramChanged or (False if 'LOG' in args.ignore else paramStruct1.LOG != paramStruct2.LOG)
            paramChanged = paramChanged or (False if 'DISPLAY' in args.ignore else paramStruct1.DISPLAY != paramStruct2.DISPLAY)
            paramChanged = paramChanged or (False if 'STATISTICS' in args.ignore else paramStruct1.STATISTICS != paramStruct2.STATISTICS)
            paramChanged = paramChanged or (False if 'RANGE' in args.ignore else paramStruct1.RANGE != paramStruct2.RANGE)
            paramChanged = paramChanged or (False if 'ERROR' in args.ignore else paramStruct1.ERROR != paramStruct2.ERROR)

            if not paramChanged:
                diffSheet.write (rowCnt, 0, 'Not changed')
                diffSheet.write (rowCnt, 1, meas1)
                diffSheet.write (rowCnt, 2, paramStruct1.NUM)
                diffSheet.write (rowCnt, 3, paramStruct1.PROMPT)
                diffSheet.write (rowCnt, 4, paramStruct1.UNIT)
                diffSheet.write (rowCnt, 5, paramStruct1.MEAS_TYPE)
                diffSheet.write (rowCnt, 6, paramStruct1.BIN)
                diffSheet.write (rowCnt, 7, paramStruct1.RES)
                diffSheet.write (rowCnt, 8, paramStruct1.MIN)
                diffSheet.write (rowCnt, 9, paramStruct1.MIN_VALUE)
                diffSheet.write (rowCnt, 10, paramStruct1.MAX)
                diffSheet.write (rowCnt, 11, paramStruct1.MAX_VALUE)
                diffSheet.write (rowCnt, 12, paramStruct1.NOM)
                diffSheet.write (rowCnt, 13, paramStruct1.NOM_VALUE)
                diffSheet.write (rowCnt, 14, paramStruct1.ARRAY_SIZE)
                diffSheet.write (rowCnt, 15, paramStruct1.LOG)
                diffSheet.write (rowCnt, 16, paramStruct1.DISPLAY)
                diffSheet.write (rowCnt, 17, paramStruct1.STATISTICS)
                diffSheet.write (rowCnt, 18, paramStruct1.RANGE)
                diffSheet.write (rowCnt, 19, paramStruct1.ERROR)
                diffSheet.write (rowCnt, 20, meas1)
                diffSheet.write (rowCnt, 21, paramStruct2.NUM)
                diffSheet.write (rowCnt, 22, paramStruct2.PROMPT)
                diffSheet.write (rowCnt, 23, paramStruct2.UNIT)
                diffSheet.write (rowCnt, 24, paramStruct2.MEAS_TYPE)
                diffSheet.write (rowCnt, 25, paramStruct2.BIN)
                diffSheet.write (rowCnt, 26, paramStruct2.RES)
                diffSheet.write (rowCnt, 27, paramStruct2.MIN)
                diffSheet.write (rowCnt, 28, paramStruct2.MIN_VALUE)
                diffSheet.write (rowCnt, 29, paramStruct2.MAX)
                diffSheet.write (rowCnt, 30, paramStruct2.MAX_VALUE)
                diffSheet.write (rowCnt, 31, paramStruct2.NOM)
                diffSheet.write (rowCnt, 32, paramStruct2.NOM_VALUE)
                diffSheet.write (rowCnt, 33, paramStruct2.ARRAY_SIZE)
                diffSheet.write (rowCnt, 34, paramStruct2.LOG)
                diffSheet.write (rowCnt, 35, paramStruct2.DISPLAY)
                diffSheet.write (rowCnt, 36, paramStruct2.STATISTICS)
                diffSheet.write (rowCnt, 37, paramStruct2.RANGE)
                diffSheet.write (rowCnt, 38, paramStruct2.ERROR)

                notChangedParameters += 1
                rowCnt += 1
                
    diffSheet.write (rowParamRemoved, 1, removedParameters)
    diffSheet.write (rowParamAdded, 1, addedParameters)
    diffSheet.write (rowParamChanged, 1, changedParameters, xlsxFormatChanged)
    diffSheet.write (rowParamNotChanged, 1, notChangedParameters)
    
    renamedParameters = 0
    for param in transTable:
        if ( (param != transTable[param]) & (param != 'Old_Param') ):
            renamedParameters += 1
    diffSheet.write (rowParamRenamed, 1, renamedParameters, xlsxFormatRenamed)
    
    #*** Bincodes ***
    rowCnt += 1
    # Overview
    diffSheet.write (rowCnt, 0, 'Bincodes removed ')
    rowBinRemoved = rowCnt
    rowCnt += 1
    diffSheet.write (rowCnt, 0, 'Bincodes added')
    rowBinAdded = rowCnt
    rowCnt += 1
    diffSheet.write (rowCnt, 0, 'Bincodes changed')
    rowBinChanged = rowCnt
    rowCnt += 1
    diffSheet.write (rowCnt, 0, 'Bincodes not changed')
    rowBinNotChanged = rowCnt
    rowCnt += 1

    # Add headers
    diffSheet.write (rowCnt, 1, nameFile1)
    diffSheet.write (rowCnt, 4, nameFile2)
    rowCnt += 1
    diffSheet.write (rowCnt, 0, 'Compare result')
    diffSheet.write (rowCnt, 1, 'Bincode')
    diffSheet.write (rowCnt, 2, 'Bin type')
    diffSheet.write (rowCnt, 3, 'Bin description')
    diffSheet.write (rowCnt, 4, 'Bincode')
    diffSheet.write (rowCnt, 5, 'Bin type')
    diffSheet.write (rowCnt, 6, 'Bin description')
    rowCnt += 1
    
    # Report removed bincodes from DMLF file 2
    addedBincodes = 0
    changedBincodes = 0
    notChangedBincodes = 0
    removedBincodes = 0
        
    for bin1, binStruct1 in binDict1.items ():
        if bin1 not in binDict2:
            diffSheet.write (rowCnt, 0, 'Removed')
            diffSheet.write (rowCnt, 1, bin1, xlsxFormatChanged)
            diffSheet.write (rowCnt, 2, binStruct1.BinType, xlsxFormatChanged)
            diffSheet.write (rowCnt, 3, binStruct1.BinDesc, xlsxFormatChanged)

            removedBincodes += 1
            rowCnt += 1
            
    # Report added bincodes in DMLF file 2
    for bin2, binStruct2 in binDict2.items ():
        if bin2 not in binDict1:
            diffSheet.write (rowCnt, 0, 'Added')
            diffSheet.write (rowCnt, 4, bin2, xlsxFormatChanged)
            diffSheet.write (rowCnt, 5, binStruct2.BinType, xlsxFormatChanged)
            diffSheet.write (rowCnt, 6, binStruct2.BinDesc, xlsxFormatChanged)
            
            addedBincodes += 1
            rowCnt += 1
            
    # Report changed bincodes
    for bin1, binStruct1 in binDict1.items ():
        if bin1 in binDict2:
            binStruct2 = binDict2[bin1]
            binTypeChanged = binStruct1.BinType != binStruct2.BinType
            binDescChanged = binStruct1.BinDesc != binStruct2.BinDesc
            if (binTypeChanged or binDescChanged):
                diffSheet.write (rowCnt, 0, 'Changed')
                diffSheet.write (rowCnt, 1, bin1)
                diffSheet.write (rowCnt, 4, bin1)
                if binTypeChanged:
                    diffSheet.write (rowCnt, 2, binStruct1.BinType, xlsxFormatChanged)
                    diffSheet.write (rowCnt, 5, binStruct2.BinType, xlsxFormatChanged)
                else:
                    diffSheet.write (rowCnt, 2, binStruct1.BinType)
                    diffSheet.write (rowCnt, 5, binStruct2.BinType)
                if binDescChanged:
                    diffSheet.write (rowCnt, 3, binStruct1.BinDesc, xlsxFormatChanged)
                    diffSheet.write (rowCnt, 6, binStruct2.BinDesc, xlsxFormatChanged)
                else:
                    diffSheet.write (rowCnt, 3, binStruct1.BinDesc)
                    diffSheet.write (rowCnt, 6, binStruct2.BinDesc)

                changedBincodes += 1
                rowCnt += 1
                
    # Report not changed bincodes
    for bin1, binStruct1 in binDict1.items ():
        if bin1 in binDict2:
            binStruct2 = binDict2[bin1]
            if ((binStruct1.BinType == binStruct2.BinType) and (binStruct1.BinDesc == binStruct2.BinDesc)):
                diffSheet.write (rowCnt, 0, 'Not changed')
                diffSheet.write (rowCnt, 1, bin1)
                diffSheet.write (rowCnt, 2, binStruct1.BinType)
                diffSheet.write (rowCnt, 3, binStruct1.BinDesc)
                diffSheet.write (rowCnt, 4, bin1)
                diffSheet.write (rowCnt, 5, binStruct2.BinType)
                diffSheet.write (rowCnt, 6, binStruct2.BinDesc)
                
                notChangedBincodes += 1
                rowCnt += 1

    diffSheet.write (rowBinRemoved, 1, removedBincodes)
    diffSheet.write (rowBinAdded, 1, addedBincodes)
    diffSheet.write (rowBinChanged, 1, changedBincodes)
    diffSheet.write (rowBinNotChanged, 1, notChangedBincodes)

#### __main__ ####    
parser = argparse.ArgumentParser ()
parser.add_argument ('file1', help='the link of first DMLF file')
parser.add_argument ('file2', help='the link of second DMLF file')
parser.add_argument ('--transFile', default = '', help='Option to use translate table file for adjusted names of parameters. E.g --transFile transFile_12125.csv')
parser.add_argument ('--ignore', default='', choices=ParamStruct._fields, help='Option to Ignore the listed fields in the parameter comparison')
parser.add_argument ('--hide', default='', help='Option to hide columns in report. E.g --hide PROMPT,MIN,MAX,NOM,ERROR,DISPLAY,RES,BIN,ARRAY_SIZE,LOG,ERROR')
args = parser.parse_args ()

if args.transFile:
    print ('args.transFile: ' + args.transFile)
    
if args.ignore:
    print ('args.ignore: ' + args.ignore)

if args.hide:
    print('args.hide: ' + args.hide)   
    
pathFile1 = os.path.abspath (args.file1)
pathFile2 = os.path.abspath (args.file2)
baseNameFile1 = os.path.basename (args.file1)
baseNameFile2 = os.path.basename (args.file2)
pathOutputFile = os.path.abspath (os.path.join (os.getcwd (),
                                                'diff_' + baseNameFile1 + '_' + baseNameFile2 + '.xlsx'))                                                   
    
outputFile = xlsxwriter.Workbook (pathOutputFile)
outputFileLogSheet = outputFile.add_worksheet ('Log')    
outputFileDiffSheet = outputFile.add_worksheet ('Diff')

xlsxFormatChanged = outputFile.add_format ()
xlsxFormatChanged.set_bg_color ('#FFCC66')
xlsxFormatRenamed = outputFile.add_format ()
xlsxFormatRenamed.set_bg_color ('#00FF00')

LogRow = 0

binDict1 = dict ()
paramDict1 = dict ()
binDict2 = dict ()
paramDict2 = dict ()

LogMessage ('Read ' + pathFile1)
parseDMLFFile (pathFile1, binDict1, paramDict1)

LogMessage ('Read ' + pathFile2)
parseDMLFFile (pathFile2, binDict2, paramDict2)

transTable = dict ()
if args.transFile:    
    pathTransFile = os.path.abspath (args.transFile)
    print ("pathTransFile," + pathTransFile)
    outputFileTransSheet = outputFile.add_worksheet ('TransTable')
    LogMessage ('Read ' + pathTransFile)
    parseTransFile(pathTransFile, transTable)
    

LogMessage ('Create diff report in ' + pathOutputFile)

if args.transFile: 
    doDiffWithTranslation (outputFileDiffSheet, baseNameFile1, baseNameFile2, binDict1, paramDict1, binDict2, paramDict2, xlsxFormatChanged, xlsxFormatRenamed,transTable)
else:
    doDiff (outputFileDiffSheet, baseNameFile1, baseNameFile2, binDict1, paramDict1, binDict2, paramDict2, xlsxFormatChanged)

hideColumns = args.hide.split(',')
if( (args.hide != '') & (len(hideColumns) >= 1) ):
    for col in hideColumns:
        print("hidden column:" + col)
        if col == 'NUM': 
            outputFileDiffSheet.set_column(2,2,None, None, {'hidden': True})
            outputFileDiffSheet.set_column(21,21,None, None, {'hidden': True})
        if col == 'PROMPT':            
            outputFileDiffSheet.set_column(3,3,None, None, {'hidden': 1})
            outputFileDiffSheet.set_column(22,22, None, None,{'hidden': 1})
        if col == 'UNIT': 
            outputFileDiffSheet.set_column(4,4,None, None, {'hidden': True})
            outputFileDiffSheet.set_column(23,23,None, None, {'hidden': True})
        if col == 'MEAS_TYPE': 
            outputFileDiffSheet.set_column(5,5,None, None, {'hidden': True})
            outputFileDiffSheet.set_column(24,24,None, None, {'hidden': True})            
        if col == 'BIN': 
            outputFileDiffSheet.set_column(6,6,None, None, {'hidden': True})
            outputFileDiffSheet.set_column(25,25,None, None, {'hidden': True})
        if col == 'RES': 
            outputFileDiffSheet.set_column(7,7,None, None, {'hidden': True})
            outputFileDiffSheet.set_column(26,26,None, None, {'hidden': True})    
        if col == 'MIN': 
            outputFileDiffSheet.set_column(8,8,None, None, {'hidden': True})
            outputFileDiffSheet.set_column(27,27,None, None, {'hidden': True})
        if col == 'MIN_VALUE': 
            outputFileDiffSheet.set_column(9,9,None, None, {'hidden': True})
            outputFileDiffSheet.set_column(28,28,None, None, {'hidden': True})
        if col == 'MAX': 
            outputFileDiffSheet.set_column(10,10,None, None, {'hidden': True})
            outputFileDiffSheet.set_column(29,29,None, None, {'hidden': True})
        if col == 'MAX_VALUE': 
            outputFileDiffSheet.set_column(11,11,None, None, {'hidden': True})
            outputFileDiffSheet.set_column(30,30,None, None, {'hidden': True})
        if col == 'NOM': 
            outputFileDiffSheet.set_column(12,12,None, None, {'hidden': True})
            outputFileDiffSheet.set_column(31,31,None, None, {'hidden': True})
        if col == 'NOM_VALUE': 
            outputFileDiffSheet.set_column(13,13,None, None, {'hidden': True})
            outputFileDiffSheet.set_column(32,32,None, None, {'hidden': True})
        if col == 'ARRAY_SIZE': 
            outputFileDiffSheet.set_column(14,14,None, None, {'hidden': True})
            outputFileDiffSheet.set_column(33,33,None, None, {'hidden': True})
        if col == 'LOG': 
            outputFileDiffSheet.set_column(15,15,None, None, {'hidden': True})
            outputFileDiffSheet.set_column(34,34,None, None, {'hidden': True})
        if col == 'DISPLAY': 
            outputFileDiffSheet.set_column(16,16,None, None, {'hidden': True})
            outputFileDiffSheet.set_column(35,35,None, None, {'hidden': True})
        if col == 'STATISTICS': 
            outputFileDiffSheet.set_column(17,17,None, None, {'hidden': True})
            outputFileDiffSheet.set_column(36,36,None, None, {'hidden': True})            
        if col == 'RANGE': 
            outputFileDiffSheet.set_column(18,18,None, None, {'hidden': True})
            outputFileDiffSheet.set_column(37,37,None, None, {'hidden': True})
        if col == 'ERROR': 
            outputFileDiffSheet.set_column(19,19,None, None, {'hidden': True})
            outputFileDiffSheet.set_column(38,38,None, None, {'hidden': True})            
outputFile.close () # saving output file

print ("pathFile1," + pathFile1)
print ("pathFile2," + pathFile2)    
print ("pathOutputFile," + pathOutputFile)

#print (binDict1)
#print (paramDict1)