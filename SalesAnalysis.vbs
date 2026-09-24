KillProcess("excel.exe")

Dim colOrder, colLine, colStatus, colItem1, colItem2, colQtyOrdered, colUM
Dim colUnitPrice, colOrderDate, colDueDate, colInvoiced, colName, colNetPrice, colCurrency
Dim fileText, lines, headerFields, headerIndex, columnCount, requiredCols
Dim orderLinesDict, rowFields, newRow, keyValue
Dim idxOrder, idxLine, idxStatus, idxItem1, idxItem2, idxQtyOrdered, idxUM
Dim idxUnitPrice, idxOrderDate, idxDueDate, idxInvoiced, idxName, idxNetPrice, idxCurrency
Dim yearFrom, yearTo, dueYear, dueMonth, dateValue, statusValue, monthYear
Dim qtyOrdered, invoiced, netPrice, converted, lookupValue
Dim fxDict, orderTypeDict, aopDict, productLineDict, unifiedDict, rowsNameUnified
Dim rowsKept, rowsOutOfWindow, rowsBadStatus, rowsPlannedZero, rowsDuplicate
Dim rowsNoFY, rowsNoProductLine, rowsNoRate, j, masterHeader, masterDict
'Dim masterSourceCols, mstrIdx, mstrLines, mstrFields, mstrHeaderIdx
'Dim mstrRow, mstrKey, mstrCsv
'Dim mstrDate, mstrYear, mstrMonth, mstrMY, mstrQty, mstrInvoiced, mstrNetPrice, mstrCurrency
'Dim mstrBlankKey, mstrDupe
Dim sheet_map, inputBook, dvsXl, dvsWb, dvsWs
Dim sheetNames, inputData, sheetData, missingSheets, sheetName, i
Dim shell, rootFolder, masterCsvFile, in_Xlsx_file, out_Csv_file
Dim fso, inputFolder, archiveFolder, outputFolder, delimiter, rowsSourceConflict
Dim inputSubFolders, oneSub, inFolderCustomerOrderLinesPath, inFolderBjornPath, inFolderEcNotInSLPath
Dim renameCols, renameFields, rn, unifiedName
Dim bjornHeaderIndex, colCalendarYear, idxCalendarYear
Dim forecastSource, forecastDict, forecastKeys, sourceRow, forecastKey
Dim fcNetPrice, fcCurrency, fcMonth, fcYear
Dim aopSource, aopMasterDict, aopSheetKeys, aopRowIn, aopKey
Dim aopMonth, aopYear, rowsAop
Dim sourceDicts, sourceNames, sourceDict, sourceKeys, sourceRowIn, existingRow
Dim outputBuffer, outputKeys, changedCols, valueOld, valueNew
Dim rowsInserted, rowsUpdated, rowsUnchanged, changeLogged
Dim d, k, n, c
Dim preserveMasterOnBlank
Dim usdValue, rowsNoUsd, rowsZeroUsd
Dim statusRules
Dim ruleKeys, ruleRow, oneRule, ruleMatches, cutoffDate, rowsOverridden, r
Dim xlsExcel, xlsBook, xlsSheet, xlsData, xlsFile
Dim xlsKeys, xlsRow, xlsCol, xlsValues

Const map_SHEET = 0
Const map_HDRROW = 1
Const map_FIRSTROW = 2
Const map_LASTCOL = 3
Const map_KEYCOLS = 4
Const map_VALUECOL = 5

colOrder = "Order"
colLine = "Line"
colStatus = "Status"
colItem1 = "Item"
colItem2 = "Item_2"
colQtyOrdered = "Qty Ordered"
colUM = "U/M"
colUnitPrice = "Unit Price"
colOrderDate = "Order Date"
colDueDate = "Due Date"
colInvoiced = "Invoiced"
colName = "Name"
colNetPrice = "Net Price"
colCurrency = "Currency"
masterHeader = Array("Order", "Line", "Status (SL)", "Status SIOP", "Order + Pos", "Item", "Item_2", "Qty Ordered", "U/M", "Unit Price", "Order Date", "Due Date",              "Month", "Year", "M+Y", "FY", "Name", "Order type", "Net Price", "Currency", "Net Price EUR", "Net Price USD", "Product Line", "Qty Difference")

Const M_ORDER = 0      : Const M_LINE = 1       : Const M_STATUS = 2
Const M_STATUSSIOP = 3 : Const M_ORDERPOS = 4   : Const M_ITEM1 = 5
Const M_ITEM2 = 6      : Const M_QTY = 7        : Const M_UM = 8
Const M_UNITPRICE = 9  : Const M_ORDERDATE = 10 : Const M_DUEDATE = 11
Const M_MONTH = 12     : Const M_YEAR = 13      : Const M_MY = 14
Const M_FY = 15        : Const M_NAME = 16      : Const M_ORDERTYPE = 17
Const M_NETPRICE = 18  : Const M_CURRENCY = 19  : Const M_NPEUR = 20
Const M_NPUSD = 21     : Const M_PRODLINE = 22  : Const M_QTYDIFF = 23

Set shell = CreateObject("WScript.Shell")
rootFolder = shell.CurrentDirectory & "\"
inputFolder = rootFolder & "Input"
archiveFolder = rootFolder & "Archive"
outputFolder = rootFolder & "Output"
masterCsvFile = rootFolder & "MasterData.csv"
masterCsvFileName = "MasterData"
masterCsvFileExt = ".csv"
'masterSrcFilePath = inputFolder & "\" & masterCsvFileName & masterCsvFileExt ' - Change for master file in future if needed CSV-moved from Src
destArchFileLoc = archiveFolder & "\" & masterCsvFileName & "_" & TimeStamp() & masterCsvFileExt
delimiter = "|"
inFolderCustomerOrderLinesPath = inputFolder & "\CustomerOrderLines"
inFolderBjornPath = inputFolder & "\Bjorn"
inFolderEcNotInSLPath = inputFolder & "\ExportControlNotInSL"
inputSubFolders = Array("Bjorn", "CustomerOrderLines", "ExportControlNotInSL")

WriteLog "Process started."
Set fso = CreateObject("Scripting.FileSystemObject")
If Not fso.FolderExists(inputFolder) Then
    fso.CreateFolder inputFolder
    WriteLog "Created folder " & inputFolder
    WScript.Echo "Check if input files are placed in 'Input' folder"
    WScript.Quit
End If

If Not fso.FolderExists(archiveFolder) Then
    fso.CreateFolder archiveFolder
    WriteLog "Created folder " & archiveFolder
End If

If Not fso.FolderExists(outputFolder) Then
    fso.CreateFolder outputFolder
    WriteLog "Created folder " & outputFolder
End If

For Each oneSub In inputSubFolders
    If Not fso.FolderExists(inputFolder & "\" & oneSub) Then
        fso.CreateFolder inputFolder & "\" & oneSub
        WriteLog "Created folder " & inputFolder & "\" & oneSub
    End If
Next

If fso.FileExists(masterCsvFile) Then
    On Error Resume Next
    fso.CopyFile masterCsvFile, destArchFileLoc, False
    If Err.Number <> 0 Then
        WriteLog "Backup failed: " & Err.Description
        Err.Clear
        On Error GoTo 0
        MsgBox "Could not back up the master file. Process stopped.", vbCritical, "Sales Analysis"
        WScript.Quit 1
    End If
    On Error GoTo 0
    WriteLog "Backup: " & destArchFileLoc
End If
CleanArchive archiveFolder, "MasterData_", 5

'############################################# Load all Input.xlsx sheets into Dictionary #############################################

sheet_map = Array(Array("FX", 1, 2, 3, "From Currency|To Currency", "Value"),_
                  Array("OrderType", 1, 2, 2, "Spares customers", "Value"),_
                  Array("AOP", 1, 2, 5, "Concate", ""),_
                  Array("OtherForecast&NRC", 1, 2, 10, "", ""),_
                  Array("UnifiedNames", 1, 2, 2, "Customer Names", "Unified Names"),_
                  Array("ProductLine", 1, 2, 3, "Desc", "Clasification"))

inputBook = inputFolder & "\Input.xlsx"

If Not fso.FileExists(inputBook) Then
    WriteLog "Input workbook not found: " & inputBook
    MsgBox "Input.xlsx not found in the Input folder.", vbCritical, "Sales Analysis"
    WScript.Quit 1
End If

On Error Resume Next
Set dvsXl = CreateObject("Excel.Application")
dvsXl.Visible = False
dvsXl.DisplayAlerts = False
dvsXl.AskToUpdateLinks = False
dvsXl.EnableEvents = False
dvsXl.ScreenUpdating = False
Err.Clear

Set dvsWb = dvsXl.Workbooks.Open(inputBook, 0, True)
If Err.Number <> 0 Then
    WriteLog "Could not open " & inputBook & " - " & Err.Description
    CleanUpExcel dvsXl, dvsWb
    MsgBox "Could not open Input.xlsx. See log.", vbCritical, "Sales Analysis"
    WScript.Quit 1
End If

Set sheetNames = CreateObject("Scripting.Dictionary")
sheetNames.CompareMode = vbTextCompare
For Each dvsWs In dvsWb.Worksheets
    sheetNames(dvsWs.Name) = True
Next

missingSheets = ""
For i = 0 To UBound(sheet_map)
    If Not sheetNames.Exists(sheet_map(i)(map_SHEET)) Then
        missingSheets = missingSheets & " '" & sheet_map(i)(map_SHEET) & "'"
    End If
Next
If missingSheets <> "" Then
    WriteLog "Input.xlsx is missing sheet(s):" & missingSheets
    CleanUpExcel dvsXl, dvsWb
    MsgBox "Input.xlsx is missing sheet(s):" & missingSheets, vbCritical, "Sales Analysis"
    WScript.Quit 1
End If

Set inputData = CreateObject("Scripting.Dictionary")
inputData.CompareMode = vbTextCompare

For i = 0 To UBound(sheet_map)
    sheetName = sheet_map(i)(map_SHEET)
    Set sheetData = ReadSheet(dvsWb.Worksheets(sheetName), sheet_map(i)(map_HDRROW), sheet_map(i)(map_FIRSTROW), sheet_map(i)(map_LASTCOL), sheet_map(i)(map_KEYCOLS), sheet_map(i)(map_VALUECOL), sheetName)
     If sheetData Is Nothing Then
        CleanUpExcel dvsXl, dvsWb
        MsgBox "Sheet '" & sheetName & "' could not be read. See log.", vbCritical, "Sales Analysis"
        WScript.Quit 1
     End If

    Set inputData(sheetName) = sheetData
    WriteLog "Sheet '" & sheetName & "': " & sheetData.Count & " entries"

Next

On Error GoTo 0
CleanUpExcel dvsXl, dvsWb

'############################# Load master (Orderbook) -> dictionary #############################
Set masterDict = CreateObject("Scripting.Dictionary")
masterDict.CompareMode = vbTextCompare

'masterSourceCols = Array("Order","Line","Status (SL)","Status SIOP","Order + Pos","Item","Item_2","Qty Ordered","U/M", _
'    "Unit Price","Order Date","Due Date","Month","Year","M+Y","FY","Name","Order type","Net Price", _
'    "Currency","Net Price EUR","Net Price USD","Product Line","")

'mstrCsv = outputFolder & "\MasterCurrent.csv"
'If XlsxToCsv(masterSrcFilePath, "SIOP", 4, 1, 0, delimiter, mstrCsv) < 0 Then
'    MsgBox "Could not read the master file. See log.", vbCritical, "Sales Analysis"
'    WScript.Quit 1
'End If

'fileText = ReadUtf8(mstrCsv)
'If Trim(fileText) = "" Then
'    WriteLog "Master CSV is empty: " & mstrCsv
'    MsgBox "The master file produced no rows.", vbCritical, "Sales Analysis"
'    WScript.Quit 1
'End If

'fileText  = Replace(Replace(fileText, vbCrLf, vbLf), vbCr, vbLf)
'mstrLines = Split(fileText, vbLf)
'mstrFields = Split(Trim(mstrLines(0)), delimiter)

'Set mstrHeaderIdx = CreateObject("Scripting.Dictionary")
'mstrHeaderIdx.CompareMode = vbTextCompare
'For i = 0 To UBound(mstrFields)
'    If Not mstrHeaderIdx.Exists(Trim(mstrFields(i))) Then
'        mstrHeaderIdx.Add Trim(mstrFields(i)), i
'    End If
'Next
'ReDim mstrIdx(UBound(masterHeader))

'For j = 0 To UBound(masterHeader)
'    If masterSourceCols(j) = "" Then
'        mstrIdx(j) = -1
'    ElseIf Not mstrHeaderIdx.Exists(masterSourceCols(j)) Then
'        WriteLog "Master: column '" & masterSourceCols(j) & "' not found in Orderbook."
'        MsgBox "Master is missing column '" & masterSourceCols(j) & "'.", vbCritical, "Sales Analysis"
'        WScript.Quit 1
'    Else
'        mstrIdx(j) = mstrHeaderIdx(masterSourceCols(j))
'    End If
'Next

'Set masterDict = CreateObject("Scripting.Dictionary")
'masterDict.CompareMode = vbTextCompare
'mstrBlankKey = 0 : mstrDupe = 0 : rowsNoFY = 0 : rowsNoProductLine = 0 : rowsNoRate = 0

'For i = 1 To UBound(mstrLines)
' If Trim(mstrLines(i)) <> "" Then

'    rowFields = Split(mstrLines(i), delimiter)
'    ReDim masterRow(UBound(masterHeader))

'    For j = 0 To UBound(masterHeader)
'        If mstrIdx(j) = -1 Then
'            masterRow(j) = ""
'        ElseIf mstrIdx(j) <= UBound(rowFields) Then
'            masterRow(j) = Trim(rowFields(mstrIdx(j)))
'        Else
'            masterRow(j) = ""
'        End If
'    Next

'    If masterRow(M_ORDER) = "" And masterRow(M_LINE) = "" Then
'        mstrBlankKey = mstrBlankKey + 1
'    Else

'    mstrKey = UCase(masterRow(M_ORDER)) & Chr(1) & UCase(masterRow(M_LINE))
'    If masterDict.Exists(mstrKey) Then mstrDupe = mstrDupe + 1
'       masterDict(mstrKey) = masterRow
'    End If
' End If
'Next

'WriteLog "Master loaded: " & masterDict.Count & " keys | blank keys " & mstrBlankKey &          " | duplicate keys " & mstrDupe & " | no FY " & rowsNoFY &          " | no product line " & rowsNoProductLine & " | no FX rate " & rowsNoRate

'############################################# Convert XLSX to CSV - Customer Order Lines Export file #############################################
in_Xlsx_file = FindSingleXlsx(inFolderCustomerOrderLinesPath, "CustomerOrderLines")
If in_Xlsx_file = "" Then
    MsgBox "Put CustomerOrderLines input file in Input\CustomerOrderLines.", vbCritical, "Sales Analysis"
    WScript.Quit 1
End If
out_Csv_file = outputFolder & "\CustomerOrderLines.csv"

'XlsxToCsv in_Xlsx_file, 1, 1, 1, 0, delimiter, out_Csv_file
If XlsxToCsv(in_Xlsx_file, 1, 1, 1, 0, delimiter, out_Csv_file) < 0 Then
    MsgBox "Could not convert CustomerOrderLinesExport. See log.", vbCritical, "Sales Analysis"
    WScript.Quit 1
End If

Set fxDict = inputData("FX")
Set orderTypeDict = inputData("OrderType")
Set aopDict = inputData("AOP")
Set productLineDict = inputData("ProductLine")
Set unifiedDict = inputData("UnifiedNames")

fileText = ReadUtf8(out_Csv_file)
If Trim(fileText) = "" Then
    WriteLog "CustOrderLines CSV is empty or unreadable: " & out_Csv_file
    MsgBox "The converted CustomerOrderLines CSV is empty.", vbCritical, "Sales Analysis"
    WScript.Quit 1
End If

fileText = Replace(Replace(fileText, vbCrLf, vbLf), vbCr, vbLf)
lines = Split(fileText, vbLf)
 
    'Array( 4, "Item"),        _   ' D
    'Array(69, "Currency")     _   ' BQ
renameCols = Array(Array( 4, "Item"), Array( 5, "Item_2"),Array(22, "Currency_2"),Array(69, "Currency"))
renameFields = Split(Trim(lines(0)), delimiter)
For rn = 0 To UBound(renameCols)
    If renameCols(rn)(0) - 1 <= UBound(renameFields) Then
        WriteLog "CustomerOrderLines: column " & renameCols(rn)(0) & " '" & renameFields(renameCols(rn)(0) - 1) & "' -> '" & renameCols(rn)(1) & "'"
        renameFields(renameCols(rn)(0) - 1) = renameCols(rn)(1)
    Else
        WriteLog "CustomerOrderLines: column " & renameCols(rn)(0) & _
                 " is past the end of the header - export layout changed"
    End If
Next

lines(0)     = Join(renameFields, delimiter)
headerFields = Split(Trim(lines(0)), delimiter)
columnCount  = UBound(headerFields) + 1

headerFields = Split(Trim(lines(0)), delimiter)
columnCount = UBound(headerFields) + 1
Set headerIndex = CreateObject("Scripting.Dictionary")
headerIndex.CompareMode = vbTextCompare

For i = 0 To UBound(headerFields)
    If Not headerIndex.Exists(Trim(headerFields(i))) Then
        headerIndex.Add Trim(headerFields(i)), i
    End If
Next

requiredCols = Array(colOrder, colLine, colStatus, colItem1, colItem2, colQtyOrdered, colUM, colUnitPrice, colOrderDate, colDueDate, colInvoiced, colName, colNetPrice, colCurrency)

For i = 0 To UBound(requiredCols)
    If Not headerIndex.Exists(requiredCols(i)) Then
        WriteLog "CustOrderLines: column '" & requiredCols(i) & "' not found. Header: " & lines(0)
        MsgBox "CustomerOrderLines is missing column '" & requiredCols(i) & "'.", vbCritical, "Sales Analysis"
        WScript.Quit 1
    End If
Next

idxOrder = headerIndex(colOrder)
idxLine = headerIndex(colLine)
idxStatus = headerIndex(colStatus)
idxItem1 = headerIndex(colItem1)
idxItem2 = headerIndex(colItem2)
idxQtyOrdered = headerIndex(colQtyOrdered)
idxUM = headerIndex(colUM)
idxUnitPrice = headerIndex(colUnitPrice)
idxOrderDate = headerIndex(colOrderDate)
idxDueDate = headerIndex(colDueDate)
idxInvoiced = headerIndex(colInvoiced)
idxName = headerIndex(colName)
idxNetPrice = headerIndex(colNetPrice)
idxCurrency = headerIndex(colCurrency)
yearFrom = Year(Date) - 1
yearTo   = Year(Date) + 1

Set orderLinesDict = CreateObject("Scripting.Dictionary")
orderLinesDict.CompareMode = vbTextCompare
rowsKept = 0 : rowsOutOfWindow = 0 : rowsBadStatus = 0 : rowsPlannedZero = 0 : rowsDuplicate = 0 : rowsNoFY = 0 : rowsNoProductLine = 0 : rowsNoRate = 0

For i = 1 To UBound(lines)
 If Trim(lines(i)) <> "" Then
    rowFields = Split(lines(i), delimiter)
    ReDim newRow(columnCount - 1)

    For j = 0 To columnCount - 1
        If j <= UBound(rowFields) Then 
            newRow(j) = Trim(rowFields(j)) 
        Else 
            newRow(j) = ""
        End If 
    Next

    statusValue = UCase(newRow(idxStatus))
    dateValue   = newRow(idxDueDate)
    dueYear = 0
     If Len(dateValue) >= 7 Then
        If IsNumeric(Left(dateValue, 4)) Then dueYear = CLng(Left(dateValue, 4))
     End If

    netPrice = ToNumber(newRow(idxNetPrice))
    qtyOrdered = ToNumber(newRow(idxQtyOrdered))
    invoiced = ToNumber(newRow(idxInvoiced))

    If statusValue <> "ORDERED" And statusValue <> "PLANNED" Then
        rowsBadStatus = rowsBadStatus + 1
    ElseIf dueYear < yearFrom Or dueYear > yearTo Then
        rowsOutOfWindow = rowsOutOfWindow + 1
    ElseIf statusValue = "PLANNED" And Not IsNull(netPrice) And netPrice = 0 Then
        rowsPlannedZero = rowsPlannedZero + 1
    Else
        ReDim masterRow(UBound(masterHeader))
        For j = 0 To UBound(masterRow)
            masterRow(j) = ""
        Next
        If statusValue = "ORDERED" then
            outStatus = "Orderbook"
        Else
            outStatus = statusValue
        End If

        masterRow(M_ORDER) = newRow(idxOrder)
        masterRow(M_LINE) = newRow(idxLine)
        masterRow(M_STATUS) = newRow(idxStatus)
        masterRow(M_STATUSSIOP) = outStatus
        masterRow(M_ORDERPOS) = newRow(idxOrder) & newRow(idxLine)
        masterRow(M_ITEM1) = newRow(idxItem1)
        masterRow(M_ITEM2) = newRow(idxItem2)
        masterRow(M_QTY) = newRow(idxQtyOrdered)
        masterRow(M_UM) = newRow(idxUM)
        masterRow(M_UNITPRICE) = newRow(idxUnitPrice)
        masterRow(M_ORDERDATE) = newRow(idxOrderDate)
        masterRow(M_DUEDATE) = dateValue
        If unifiedDict.Exists(Trim(newRow(idxName))) Then
            newRow(idxName) = unifiedDict(Trim(newRow(idxName)))
            rowsNameUnified = rowsNameUnified + 1
        End If
        masterRow(M_NAME) = newRow(idxName)
        masterRow(M_NETPRICE) = newRow(idxNetPrice)
        masterRow(M_CURRENCY) = newRow(idxCurrency)
        dueMonth  = CLng(Mid(dateValue, 6, 2))
        monthYear = CStr(dueMonth) & CStr(dueYear)
        masterRow(M_MONTH) = CStr(dueMonth)
        masterRow(M_YEAR) = CStr(dueYear)
        masterRow(M_MY) = monthYear
        If aopDict.Exists(monthYear) Then
            masterRow(M_FY) = aopDict(monthYear)("FY")
        Else
            rowsNoFY = rowsNoFY + 1
        End If

        lookupValue = Trim(newRow(idxName))
        If orderTypeDict.Exists(lookupValue) Then
            masterRow(M_ORDERTYPE) = orderTypeDict(lookupValue)
        Else
            masterRow(M_ORDERTYPE) = "OEM"
        End If

        lookupValue = Trim(newRow(idxItem2))
        If productLineDict.Exists(lookupValue) Then
            masterRow(M_PRODLINE) = productLineDict(lookupValue)
        Else
            rowsNoProductLine = rowsNoProductLine + 1
        End If

        If Not IsNull(netPrice) Then
            converted = ConvertCurrency(newRow(idxCurrency), "EUR", netPrice, fxDict)
            If IsNull(converted) Then
                rowsNoRate = rowsNoRate + 1
            Else
                masterRow(M_NPEUR) = Replace(CStr(converted), ",", ".")
            End If

            converted = ConvertCurrency(newRow(idxCurrency), "USD", netPrice, fxDict)
            If Not IsNull(converted) Then
                masterRow(M_NPUSD) = Replace(CStr(converted), ",", ".")
            End If
        End If

        If Not IsNull(qtyOrdered) Then
            If IsNull(invoiced) Then invoiced = 0
            masterRow(M_QTYDIFF) = Replace(CStr(qtyOrdered - invoiced), ",", ".")
        End If

        keyValue = UCase(newRow(idxOrder)) & Chr(1) & UCase(newRow(idxLine))
        If orderLinesDict.Exists(keyValue) Then rowsDuplicate = rowsDuplicate + 1

        orderLinesDict(keyValue) = masterRow
        rowsKept = rowsKept + 1
    End If
 End If
Next

WriteLog "CustOrderLines: kept " & orderLinesDict.Count & " | out of window " & rowsOutOfWindow & " | status excluded " & rowsBadStatus & " | planned zero price " & rowsPlannedZero & " | duplicate keys " & rowsDuplicate & " | no FY " & rowsNoFY & " | no product line " & rowsNoProductLine & " | no FX rate " & rowsNoRate

'############################################# Loading - Bjorn file #######################################################################################################################################
fileText = ""
lines = ""
headerFields = ""
columnCount = ""

in_Xlsx_file = FindSingleXlsx(inFolderBjornPath, "Bjorn")
If in_Xlsx_file = "" Then
    MsgBox "Put exactly one xlsx file in Input\Bjorn File.", vbCritical, "Sales Analysis"
    WScript.Quit 1
End If
out_Csv_file = outputFolder & "\Bjorn.csv"

If XlsxToCsv(in_Xlsx_file, "Sales Analysis", 1, 1, 0, delimiter, out_Csv_file) < 0 Then
    MsgBox "Could not convert Bjorn file -add filename. See log.", vbCritical, "Sales Analysis"
    WScript.Quit 1
End If

colOrder = "Order"
colLine = "Order Line"
colItem1 = "Item"
colItem2 = "Description"
colQtyOrdered = "Qty"
colUnitPrice = "Unit Price"
colDueDate = "Invoice Date"
colName = "Name"
colNetPrice = "DOM Ext Price"
colNetPriceUSD = "USD Sales"
colCalendarYear = "Calendar Year"

fileText = ReadUtf8(out_Csv_file)
If Trim(fileText) = "" Then
    WriteLog "Bjorn CSV is empty or unreadable: " & out_Csv_file
    MsgBox "The converted Bjorn CSV is empty.", vbCritical, "Sales Analysis"
    WScript.Quit 1
End If

fileText = Replace(Replace(fileText, vbCrLf, vbLf), vbCr, vbLf)
lines = Split(fileText, vbLf)
headerFields = Split(Trim(lines(0)), delimiter)
columnCount = UBound(headerFields) + 1
Set bjornHeaderIndex = CreateObject("Scripting.Dictionary")
bjornHeaderIndex.CompareMode = vbTextCompare

For i = 0 To UBound(headerFields)
    If Not bjornHeaderIndex.Exists(Trim(headerFields(i))) Then
        bjornHeaderIndex.Add Trim(headerFields(i)), i
    End If
Next

requiredCols = Array(colOrder, colLine, colItem1, colItem2, colQtyOrdered, colUnitPrice, colDueDate, colName, colNetPrice, colCalendarYear)

For i = 0 To UBound(requiredCols)
    If Not bjornHeaderIndex.Exists(requiredCols(i)) Then
        WriteLog "Bjorn: column '" & requiredCols(i) & "' not found. Header: " & lines(0)
        MsgBox "Bjorn is missing column '" & requiredCols(i) & "'.", vbCritical, "Sales Analysis"
        WScript.Quit 1
    End If
Next

idxOrder = bjornHeaderIndex(colOrder)
idxLine = bjornHeaderIndex(colLine)
idxItem1 = bjornHeaderIndex(colItem1)
idxItem2 = bjornHeaderIndex(colItem2)
idxQtyOrdered = bjornHeaderIndex(colQtyOrdered)
idxUnitPrice = bjornHeaderIndex(colUnitPrice)
idxDueDate = bjornHeaderIndex(colDueDate)
idxName = bjornHeaderIndex(colName)
idxNetPrice = bjornHeaderIndex(colNetPrice)
idxNetPriceUSD = bjornHeaderIndex(colNetPriceUSD)
idxCalendarYear = bjornHeaderIndex(colCalendarYear)
yearFrom = Year(Date) - 1
yearTo   = Year(Date)

Set bjornDict = CreateObject("Scripting.Dictionary")
bjornDict.CompareMode = vbTextCompare
rowsKept = 0 : rowsOutOfWindow = 0 : rowsBadStatus = 0 : rowsPlannedZero = 0 : rowsDuplicate = 0 : rowsNoFY = 0 : rowsNoProductLine = 0 : rowsNoRate = 0

For i = 1 To UBound(lines)
 If Trim(lines(i)) <> "" Then
    rowFields = Split(lines(i), delimiter)
    ReDim newRow(columnCount - 1)

    For j = 0 To columnCount - 1
        If j <= UBound(rowFields) Then 
            newRow(j) = Trim(rowFields(j)) 
        Else 
            newRow(j) = ""
        End If 
    Next

    dateValue = newRow(idxDueDate)
    dueYear = 0
     If Len(dateValue) >= 7 Then
        If IsNumeric(Left(dateValue, 4)) Then dueYear = CLng(Left(dateValue, 4))
     End If

    netPrice = ToNumber(newRow(idxNetPrice))
    qtyOrdered = ToNumber(newRow(idxQtyOrdered))

    ReDim masterRow(UBound(masterHeader))
    For j = 0 To UBound(masterRow)
        masterRow(j) = ""
    Next

    masterRow(M_ORDER) = newRow(idxOrder)
    masterRow(M_LINE) = newRow(idxLine)
    masterRow(M_STATUS) = ""
    masterRow(M_STATUSSIOP) = "Invoiced"
    masterRow(M_ORDERPOS) = ""
    masterRow(M_ITEM1) = newRow(idxItem1)
    masterRow(M_ITEM2) = newRow(idxItem2)
    masterRow(M_QTY) = newRow(idxQtyOrdered)
    masterRow(M_UM) =  ""
    masterRow(M_UNITPRICE) = newRow(idxUnitPrice)
    masterRow(M_ORDERDATE) =  ""
    masterRow(M_DUEDATE) = dateValue
    If unifiedDict.Exists(Trim(newRow(idxName))) Then
        newRow(idxName) = unifiedDict(Trim(newRow(idxName)))
        rowsNameUnified = rowsNameUnified + 1
    End If
    masterRow(M_NAME) = newRow(idxName)
    masterRow(M_NETPRICE) = newRow(idxNetPrice)
    masterRow(M_CURRENCY) = "EUR"
    dueMonth  = CLng(Mid(dateValue, 6, 2))
    monthYear = CStr(dueMonth) & CStr(dueYear)
    masterRow(M_MONTH) = CStr(dueMonth)
    masterRow(M_YEAR) = CStr(dueYear)
    masterRow(M_MY) = monthYear
    If aopDict.Exists(monthYear) Then
        masterRow(M_FY) = aopDict(monthYear)("FY")
    Else
        rowsNoFY = rowsNoFY + 1
    End If

    lookupValue = Trim(newRow(idxName))
    If orderTypeDict.Exists(lookupValue) Then
        masterRow(M_ORDERTYPE) = orderTypeDict(lookupValue)
    Else
        masterRow(M_ORDERTYPE) = "OEM"
    End If
  
      lookupValue = Trim(newRow(idxItem2))
    If productLineDict.Exists(lookupValue) Then
        masterRow(M_PRODLINE) = productLineDict(lookupValue)
    Else
        rowsNoProductLine = rowsNoProductLine + 1
    End If

    masterRow(M_NPEUR) = newRow(idxNetPrice)
    If Not IsNull(netPrice) Then
        'USD Sales kolona je DOM Ext * FX rate Eur to USD
        converted = ConvertCurrency("EUR", "USD", netPrice, fxDict)
        If Not IsNull(converted) Then
            masterRow(M_NPUSD) = Replace(CStr(converted), ",", ".")
        End If
    End If

    masterRow(M_QTYDIFF) = ""

    keyValue = UCase(newRow(idxOrder)) & Chr(1) & UCase(newRow(idxLine))
    If bjornDict.Exists(keyValue) Then rowsDuplicate = rowsDuplicate + 1

    bjornDict(keyValue) = masterRow
    rowsKept = rowsKept + 1
 End If
Next

WriteLog "Bjorn: kept " & orderLinesDict.Count & " | out of window " & rowsOutOfWindow & " | status excluded " & rowsBadStatus & " | planned zero price " & rowsPlannedZero & " | duplicate keys " & rowsDuplicate & " | no FY " & rowsNoFY & " | no product line " & rowsNoProductLine & " | no FX rate " & rowsNoRate

'############################################# Loading - Export Control not in SL file #######################################################################################################################################
fileText = ""
lines = ""
headerFields = ""
columnCount = ""

in_Xlsx_file = FindSingleXlsx(inFolderEcNotInSLPath, "EC not in SL")
If in_Xlsx_file = "" Then
    MsgBox "Put exactly one xlsx file in Input\ECNotInSL File.", vbCritical, "Sales Analysis"
    WScript.Quit 1
End If
out_Csv_file = outputFolder & "\ExportControlNotInSL.csv"

If XlsxToCsv(in_Xlsx_file, 1, 4, 1, 0, delimiter, out_Csv_file) < 0 Then
    MsgBox "Could not convert EC not in SL file -add filename. See log.", vbCritical, "Sales Analysis"
    WScript.Quit 1
End If

colItem2 = "Customer PO"
colName = "Name"
colNetPrice = "Total Price"
colCurrency = "Currency"
colDueDate = "Due Date"

fileText = ReadUtf8(out_Csv_file)
 If Trim(fileText) = "" Then
    WriteLog "EC not in SL CSV is empty or unreadable: " & out_Csv_file
    MsgBox "The converted EC not in SL CSV is empty.", vbCritical, "Sales Analysis"
    WScript.Quit 1
 End If
fileText = Replace(Replace(fileText, vbCrLf, vbLf), vbCr, vbLf)
lines = Split(fileText, vbLf)
headerFields = Split(Trim(lines(0)), delimiter)
columnCount = UBound(headerFields) + 1
Set ecnislHeaderIndex = CreateObject("Scripting.Dictionary")
ecnislHeaderIndex.CompareMode = vbTextCompare

For i = 0 To UBound(headerFields)
    If Not ecnislHeaderIndex.Exists(Trim(headerFields(i))) Then
        ecnislHeaderIndex.Add Trim(headerFields(i)), i
    End If
Next

requiredCols = Array(colItem2, colDueDate, colName, colNetPrice, colCurrency)

For i = 0 To UBound(requiredCols)
    If Not ecnislHeaderIndex.Exists(requiredCols(i)) Then
        WriteLog "EC not in SL: column '" & requiredCols(i) & "' not found. Header: " & lines(0)
        MsgBox "EC not in SL is missing column '" & requiredCols(i) & "'.", vbCritical, "Sales Analysis"
        WScript.Quit 1
    End If
Next

idxItem2 = ecnislHeaderIndex(colItem2)
idxName = ecnislHeaderIndex(colName)
idxDueDate = ecnislHeaderIndex(colDueDate)
idxNetPrice = ecnislHeaderIndex(colNetPrice)
idxColCurrency = ecnislHeaderIndex(colCurrency)
Set ecnisDict = CreateObject("Scripting.Dictionary")
ecnisDict.CompareMode = vbTextCompare
rowsKept = 0 : rowsOutOfWindow = 0 : rowsBadStatus = 0 : rowsPlannedZero = 0 : rowsDuplicate = 0 : rowsNoFY = 0 : rowsNoProductLine = 0 : rowsNoRate = 0

For i = 1 To UBound(lines)
 If Trim(lines(i)) <> "" Then
    rowFields = Split(lines(i), delimiter)
    ReDim newRow(columnCount - 1)

    For j = 0 To columnCount - 1
        If j <= UBound(rowFields) Then 
            newRow(j) = Trim(rowFields(j)) 
        Else 
            newRow(j) = ""
        End If 
    Next

    dateValue = newRow(idxDueDate)
    dueYear = 0
     If Len(dateValue) >= 7 Then
        If IsNumeric(Left(dateValue, 4)) Then dueYear = CLng(Left(dateValue, 4))
     End If

    netPrice = ToNumber(newRow(idxNetPrice))

    ReDim masterRow(UBound(masterHeader))
    For j = 0 To UBound(masterRow)
        masterRow(j) = ""
    Next

    masterRow(M_ORDER) = ""
    masterRow(M_LINE) = ""
    masterRow(M_STATUS) = ""
    masterRow(M_STATUSSIOP) = "Other Forecast"
    masterRow(M_ORDERPOS) = ""
    masterRow(M_ITEM1) = ""
    masterRow(M_ITEM2) = newRow(idxItem2)
    masterRow(M_QTY) = ""
    masterRow(M_UM) =  ""
    masterRow(M_UNITPRICE) = ""
    masterRow(M_ORDERDATE) =  ""
    masterRow(M_DUEDATE) = newRow(idxDueDate)
    If unifiedDict.Exists(Trim(newRow(idxName))) Then
        newRow(idxName) = unifiedDict(Trim(newRow(idxName)))
        rowsNameUnified = rowsNameUnified + 1
    End If
    masterRow(M_NAME) = newRow(idxName)
    masterRow(M_NETPRICE) = newRow(idxNetPrice)
    masterRow(M_CURRENCY) = newRow(idxColCurrency)
    dueMonth  = CLng(Mid(dateValue, 6, 2))
    monthYear = CStr(dueMonth) & CStr(dueYear)
    masterRow(M_MONTH) = CStr(dueMonth)
    masterRow(M_YEAR) = CStr(dueYear)
    masterRow(M_MY) = monthYear
    If aopDict.Exists(monthYear) Then
        masterRow(M_FY) = aopDict(monthYear)("FY")
    Else
        rowsNoFY = rowsNoFY + 1
    End If

    lookupValue = Trim(newRow(idxName))
    If orderTypeDict.Exists(lookupValue) Then
        masterRow(M_ORDERTYPE) = orderTypeDict(lookupValue)
    Else
        masterRow(M_ORDERTYPE) = "OEM"
    End If
  
      lookupValue = Trim(newRow(idxItem2))
    If productLineDict.Exists(lookupValue) Then
        masterRow(M_PRODLINE) = productLineDict(lookupValue)
    Else
        rowsNoProductLine = rowsNoProductLine + 1
    End If

    If Not IsNull(netPrice) Then
      converted = ConvertCurrency(newRow(idxColCurrency), "EUR", netPrice, fxDict)
        If Not IsNull(converted) Then
            masterRow(M_NPEUR) = Replace(CStr(converted), ",", ".")
        End If
        
      converted = ConvertCurrency(newRow(idxColCurrency), "USD", netPrice, fxDict)
        If Not IsNull(converted) Then
            masterRow(M_NPUSD) = Replace(CStr(converted), ",", ".")
        End If
    End If

    masterRow(M_QTYDIFF) = ""

    ecnisKey = "EC" & Chr(1) & UCase(Trim(newRow(idxItem2))) & Chr(1) & monthYear & Chr(1) & UCase(Trim(newRow(idxName)))
    If ecnisDict.Exists(ecnisKey) Then rowsDuplicate = rowsDuplicate + 1
    ecnisDict(UniqueKey(ecnisDict, ecnisKey)) = masterRow

    rowsKept = rowsKept + 1
 End If
Next

WriteLog "EC not in SL: kept " & ecnisDict.Count & " | out of window " & rowsOutOfWindow & " | status excluded " & rowsBadStatus & " | planned zero price " & rowsPlannedZero & " | duplicate keys " & rowsDuplicate & " | no FY " & rowsNoFY & " | no product line " & rowsNoProductLine & " | no FX rate " & rowsNoRate

'############################# OtherForecast&NRC -> master rows #############################

Set forecastSource = inputData("OtherForecast&NRC")
Set forecastDict = CreateObject("Scripting.Dictionary")
forecastDict.CompareMode = vbTextCompare
rowsKept = 0 : rowsNoFY = 0 : rowsNoProductLine = 0 : rowsNoRate = 0
forecastKeys = forecastSource.Keys

For n = 0 To UBound(forecastKeys)
    Set sourceRow = forecastSource(forecastKeys(n))

    ReDim masterRow(UBound(masterHeader))
    For j = 0 To UBound(masterRow)
        masterRow(j) = ""
    Next

    masterRow(M_STATUSSIOP) = "Other Forecast"
    masterRow(M_ITEM1) = sourceRow("Item 1")
    masterRow(M_ITEM2) = sourceRow("Item 2")
    unifiedName = Trim(sourceRow("Name (customer)"))
    If unifiedDict.Exists(unifiedName) Then
        unifiedName = unifiedDict(unifiedName)
        rowsNameUnified = rowsNameUnified + 1
    End If
    masterRow(M_NAME) = unifiedName
    masterRow(M_ORDERTYPE) = sourceRow("Order Type")
    masterRow(M_NETPRICE) = sourceRow("Net Price")
    masterRow(M_CURRENCY) = sourceRow("Currency")
    fcMonth = Trim(sourceRow("Month"))
    fcYear = Trim(sourceRow("Year"))
    masterRow(M_MONTH) = fcMonth
    masterRow(M_YEAR) = fcYear
    monthYear = fcMonth & fcYear
    masterRow(M_MY) = monthYear
     If aopDict.Exists(monthYear) Then
        masterRow(M_FY) = aopDict(monthYear)("FY")
     Else
        rowsNoFY = rowsNoFY + 1
     End If
    lookupValue = Trim(sourceRow("Item 2"))
    If productLineDict.Exists(lookupValue) Then
        masterRow(M_PRODLINE) = productLineDict(lookupValue)
    Else
        rowsNoProductLine = rowsNoProductLine + 1
    End If

    fcNetPrice = ToNumber(sourceRow("Net Price"))
    fcCurrency = Trim(sourceRow("Currency"))
    If Not IsNull(fcNetPrice) Then
        converted = ConvertCurrency(fcCurrency, "EUR", fcNetPrice, fxDict)
        If IsNull(converted) Then
            rowsNoRate = rowsNoRate + 1
        Else
            masterRow(M_NPEUR) = Replace(CStr(converted), ",", ".")
        End If

        converted = ConvertCurrency(fcCurrency, "USD", fcNetPrice, fxDict)
        If Not IsNull(converted) Then
            masterRow(M_NPUSD) = Replace(CStr(converted), ",", ".")
        End If
    End If

    forecastKey = UniqueKey(forecastDict, "FC" & Chr(1) & UCase(Trim(sourceRow("Item 1"))) & Chr(1) & monthYear & Chr(1) & UCase(unifiedName))
    forecastDict(forecastKey) = masterRow
    rowsKept = rowsKept + 1
Next

WriteLog "OtherForecast&NRC: kept " & forecastDict.Count & " | no FY " & rowsNoFY & " | no product line " & rowsNoProductLine & " | no FX rate " & rowsNoRate

'############################# AOP -> master rows #############################


Set aopSource = inputData("AOP")
Set aopMasterDict = CreateObject("Scripting.Dictionary")
aopMasterDict.CompareMode = vbTextCompare
rowsAop = 0
aopSheetKeys = aopSource.Keys

For n = 0 To UBound(aopSheetKeys)
    Set aopRowIn = aopSource(aopSheetKeys(n))

    ReDim masterRow(UBound(masterHeader))
    For j = 0 To UBound(masterRow)
        masterRow(j) = ""
    Next
    aopMonth = Trim(aopRowIn("Month"))
    aopYear  = Trim(aopRowIn("Year"))
    masterRow(M_STATUSSIOP) = "Target"
    masterRow(M_FY) = Trim(aopRowIn("FY"))
    masterRow(M_MONTH) = aopMonth
    masterRow(M_YEAR) = aopYear
    masterRow(M_MY) = Trim(aopRowIn("Concate"))
    masterRow(M_NPUSD) = Replace(Trim(aopRowIn("AOP USD")), ",", ".")
    aopKey = UniqueKey(aopMasterDict, "AOP" & Chr(1) & aopMonth & Chr(1) & aopYear)
    aopMasterDict(aopKey) = masterRow
    rowsAop = rowsAop + 1
Next

WriteLog "AOP: " & aopMasterDict.Count & " rows prepared for master"

'############################# Merge sources into master #############################
preserveMasterOnBlank = False
sourceDicts = Array(orderLinesDict, bjornDict, ecnisDict, forecastDict, aopMasterDict)
sourceNames = Array("CustomerOrderLines", "Bjorn", "EC not in SL", "OtherForecast&NRC", "AOP")
rowsInserted = 0 : rowsUpdated = 0 : rowsUnchanged = 0 : rowsSourceConflict = 0 : rowsNoUsd = 0 : rowsZeroUsd = 0

For d = 0 To UBound(sourceDicts)
    Set sourceDict = sourceDicts(d)
    sourceKeys = sourceDict.Keys

    For n = 0 To UBound(sourceKeys)
        k = sourceKeys(n)
        sourceRowIn = sourceDict(k)
        usdValue = ToNumber(sourceRowIn(M_NPUSD))

        If IsNull(usdValue) Then
            rowsNoUsd = rowsNoUsd + 1
        ElseIf usdValue = 0 Then
            rowsZeroUsd = rowsZeroUsd + 1
        ElseIf Not masterDict.Exists(k) Then
            masterDict(k) = sourceRowIn
            rowsInserted = rowsInserted + 1
        Else
            existingRow = masterDict(k)
            changedCols = ""

            If Trim(sourceRowIn(M_STATUSSIOP)) <> "" Then
                If StrComp(Trim(existingRow(M_STATUSSIOP)), Trim(sourceRowIn(M_STATUSSIOP)), vbTextCompare) <> 0 Then
                    rowsSourceConflict = rowsSourceConflict + 1
                    WriteLog "CONFLICT " & Replace(k, Chr(1), " + ") & ": already '" & existingRow(M_STATUSSIOP) & "', " & sourceNames(d) & " reports '" & sourceRowIn(M_STATUSSIOP) & "'"
                End If
            End If

            For c = 0 To UBound(masterHeader)
                valueOld = existingRow(c)
                valueNew = sourceRowIn(c)

                If preserveMasterOnBlank And Trim(valueNew) = "" Then
                ElseIf StrComp(Trim(valueOld), Trim(valueNew), vbTextCompare) <> 0 Then
                    changedCols = changedCols & masterHeader(c) & " '" & valueOld & "' -> '" & valueNew & "'; "
                    existingRow(c) = valueNew
                End If
            Next

            If changedCols = "" Then
                rowsUnchanged = rowsUnchanged + 1
            Else
                masterDict(k) = existingRow
                rowsUpdated = rowsUpdated + 1

                If changeLogged < 200 Then
                    WriteLog "Update " & Replace(k, Chr(1), " + ") & " from " & sourceNames(d) & ": " & changedCols
                    changeLogged = changeLogged + 1
                End If
            End If
        End If
    Next

    WriteLog "Merged " & sourceNames(d) & ": " & sourceDict.Count & " rows processed, " & (rowsSourceConflict - conflictsBefore) & " conflicts"
Next

WriteLog "Merge result: " & rowsInserted & " inserted, " & rowsUpdated & " updated, " & rowsUnchanged & " unchanged, " & rowsSourceConflict & " conflicts, " & rowsNoUsd & " no USD, " & rowsZeroUsd & " zero USD | master now " & masterDict.Count & " rows"

'############################# Status SIOP overrides #############################
statusRules = Array(Array("Airbus Operations GMBH", "", "P", 0, "SL Forecast"), _
                    Array("GOODRICH ACTUATION", "", "A", 0, "SL Forecast"), _
                    Array("ROLLS ROYCE", "", "",  3, "SL Forecast"), _
                    Array("", "L32A320N-70A", "A", 0, "SL Forecast"), _
                    Array("DIEHL AVIATION GILCHING GMBH", "", "P", 0, "SL Forecast"))
Const RULE_NAME     = 0
Const RULE_ITEM     = 1
Const RULE_ORDER    = 2
Const RULE_MONTHS   = 3
Const RULE_STATUS   = 4
cutoffDate = ""
rowsOverridden = 0
ruleKeys = masterDict.Keys

For n = 0 To UBound(ruleKeys)
    ruleRow = masterDict(ruleKeys(n))

    For r = 0 To UBound(statusRules)
        oneRule = statusRules(r)
        ruleMatches = True

        If oneRule(RULE_NAME) <> "" Then
            If StrComp(Trim(ruleRow(M_NAME)), oneRule(RULE_NAME), vbTextCompare) <> 0 Then
                ruleMatches = False
            End If
        End If

        If ruleMatches And oneRule(RULE_ITEM) <> "" Then
            If StrComp(Trim(ruleRow(M_ITEM1)), oneRule(RULE_ITEM), vbTextCompare) <> 0 Then
                ruleMatches = False
            End If
        End If

        If ruleMatches And oneRule(RULE_ORDER) <> "" Then
            If InStr(1, ruleRow(M_ORDER), oneRule(RULE_ORDER), vbTextCompare) = 0 Then
                ruleMatches = False
            End If
        End If

        If ruleMatches And oneRule(RULE_MONTHS) > 0 Then
            cutoffDate = IsoDate(DateAdd("m", oneRule(RULE_MONTHS), Date))
            If Len(ruleRow(M_DUEDATE)) < 10 Then
                ruleMatches = False
            ElseIf ruleRow(M_DUEDATE) <= cutoffDate Then
                ruleMatches = False
            End If
        End If

        If ruleMatches Then
            If StrComp(ruleRow(M_STATUSSIOP), oneRule(RULE_STATUS), vbTextCompare) <> 0 Then
                WriteLog "Override " & Replace(ruleKeys(n), Chr(1), " + ") & ": '" & ruleRow(M_STATUSSIOP) & "' -> '" & oneRule(RULE_STATUS) & "' (rule " & (r + 1) & ")"
                ruleRow(M_STATUSSIOP) = oneRule(RULE_STATUS)
                masterDict(ruleKeys(n)) = ruleRow
                rowsOverridden = rowsOverridden + 1
            End If
            Exit For
        End If
    Next
Next

WriteLog "Status SIOP overrides applied: " & rowsOverridden & " rows"

'############################# Write MasterData.csv #############################

outputKeys = masterDict.Keys
ReDim outputBuffer(masterDict.Count)
outputBuffer(0) = Join(masterHeader, delimiter)

For n = 0 To UBound(outputKeys)
    outputBuffer(n + 1) = Join(masterDict(outputKeys(n)), delimiter)
Next

If Not WriteUtf8(masterCsvFile, Join(outputBuffer, vbCrLf) & vbCrLf) Then
    MsgBox "Could not write MasterData.csv. See log.", vbCritical, "Sales Analysis"
    WScript.Quit 1
End If

WriteLog "MasterData.csv written: " & masterDict.Count & " rows, " & (UBound(masterHeader) + 1) & " columns"

'############################# Export MasterData.xlsx #############################
xlsFile = rootFolder & "MasterData.xlsx"
On Error Resume Next
Set xlsExcel = CreateObject("Excel.Application")
xlsExcel.Visible = False
xlsExcel.DisplayAlerts = False
xlsExcel.EnableEvents = False
xlsExcel.ScreenUpdating = False
Err.Clear

Set xlsBook  = xlsExcel.Workbooks.Add
Set xlsSheet = xlsBook.Worksheets(1)
xlsSheet.Name = "MasterData"

If Err.Number <> 0 Then
    WriteLog "xlsx export: could not start Excel - " & Err.Description
    CleanUpExcel xlsExcel, xlsBook
    Err.Clear
Else
    xlsKeys = masterDict.Keys
    ReDim xlsData(masterDict.Count, UBound(masterHeader))

    For xlsCol = 0 To UBound(masterHeader)
        xlsData(0, xlsCol) = masterHeader(xlsCol)
    Next
    For n = 0 To UBound(xlsKeys)
        xlsValues = masterDict(xlsKeys(n))
        For xlsCol = 0 To UBound(masterHeader)
            xlsData(n + 1, xlsCol) = xlsValues(xlsCol)
        Next
    Next

    xlsSheet.Range(xlsSheet.Cells(1, 1), xlsSheet.Cells(masterDict.Count + 1, UBound(masterHeader) + 1)).Value = xlsData
    xlsSheet.Rows(1).Font.Bold = True
    xlsSheet.Range(xlsSheet.Cells(1, 1), xlsSheet.Cells(1, UBound(masterHeader) + 1)).AutoFilter
    xlsExcel.ActiveWindow.FreezePanes = False
    xlsSheet.Range("A2").Select
    xlsExcel.ActiveWindow.FreezePanes = True
    xlsSheet.Columns.AutoFit

    If fso.FileExists(xlsFile) Then fso.DeleteFile xlsFile, True
    xlsBook.SaveAs xlsFile, 51

    If Err.Number <> 0 Then
        WriteLog "xlsx export failed: " & Err.Description
        Err.Clear
    Else
        WriteLog "MasterData.xlsx written: " & masterDict.Count & " rows"
    End If
    CleanUpExcel xlsExcel, xlsBook
End If

On Error GoTo 0

MsgBox "Done." & vbCrLf & vbCrLf & "Inserted:   " & rowsInserted & vbCrLf & "Updated:    " & rowsUpdated & vbCrLf & _
       "Unchanged:  " & rowsUnchanged & vbCrLf & "Conflicts:  " & rowsSourceConflict & vbCrLf & vbCrLf & "Skipped - no USD value: " & rowsNoUsd & vbCrLf & _
       "Skipped - zero USD:     " & rowsZeroUsd & vbCrLf & vbCrLf & "Status SIOP overrides:  " & rowsOverridden & vbCrLf & vbCrLf & "Total rows in MasterData.csv: " & masterDict.Count, vbInformation, "Sales Analysis"

'#####################################################################################################################################
'                                                               FUNCTIONS
'#####################################################################################################################################

Function XlsxToCsv(srcBook, sheetName, headerRow, firstCol, lastCol, sDelim, dstCsv)

    Dim dvsXl, dvsWb, dvsWs, lastRow, arrHdr, arrData, sOut, nRows, nCols
    XlsxToCsv = -1
    Set dvsXl = Nothing
    Set dvsWb = Nothing

    If Not fso.FileExists(srcBook) Then
        WriteLog("Workbook not found: " & srcBook)
        Exit Function
    End If

    On Error Resume Next

    Set dvsXl = CreateObject("Excel.Application")
    dvsXl.Visible = False
    dvsXl.DisplayAlerts = False
    dvsXl.AskToUpdateLinks = False
    dvsXl.EnableEvents = False
    dvsXl.ScreenUpdating = False

    Set dvsWb = dvsXl.Workbooks.Open(srcBook, 0, True)
        If Err.Number <> 0 Then
         WriteLog("Could not open workbook: " & srcBook & " - " & Err.Description)
         CleanUpExcel dvsXl, dvsWb
        Exit Function
    End If
    Err.Clear
    Set dvsWs = dvsWb.Worksheets(sheetName)
        If Err.Number <> 0 Or dvsWs Is Nothing Then
            WriteLog "Sheet '" & sheetName & "' not found in " & fso.GetFileName(srcBook)
            CleanUpExcel dvsXl, dvsWb
        Exit Function
    End If
    Err.Clear

    If lastCol = 0 Then
        lastCol = dvsWs.Cells(headerRow, dvsWs.Columns.Count).End(-4159).Column   'XlToLeft
        If Err.Number <> 0 Or lastCol < firstCol Then
            WriteLog("Could not determine last column on row " & headerRow)
            CleanUpExcel dvsXl, dvsWb
            Exit Function
        End If
    End If
    nCols = lastCol - firstCol + 1

    lastRow = dvsWs.UsedRange.Row + dvsWs.UsedRange.Rows.Count - 1
    If lastRow <= headerRow Then
        WriteLog("Sheet '" & sheetName & "' has no data rows below row " & headerRow)
        CleanUpExcel dvsXl, dvsWb
        Exit Function
    End If

    arrHdr  = dvsWs.Range(dvsWs.Cells(headerRow, firstCol), dvsWs.Cells(headerRow, lastCol)).Value
    arrData = dvsWs.Range(dvsWs.Cells(headerRow + 1, firstCol), dvsWs.Cells(lastRow, lastCol)).Value

    If Err.Number <> 0 Then
        WriteLog("Failed reading '" & sheetName & "': " & Err.Description)
        CleanUpExcel dvsXl, dvsWb
        Exit Function
    End If
    On Error GoTo 0

    CleanUpExcel dvsXl, dvsWb        ' release Excel before the write

    If Not IsArray(arrHdr)  Then arrHdr  = OneCellArray(arrHdr)
    If Not IsArray(arrData) Then arrData = OneCellArray(arrData)

    sOut = BuildHeaderLine(arrHdr, nCols, sDelim) & vbCrLf &     BuildDataLines(arrData, nCols, sDelim, nRows)

    If Not WriteUtf8(dstCsv, sOut) Then Exit Function

   XlsxToCsv = nRows

End Function

Function IsoDate(inputDate)
    IsoDate = Year(inputDate) & "-" & P2(Month(inputDate)) & "-" & P2(Day(inputDate))
End Function

Function BuildHeaderLine(arrHdr, nCols, sDelim)

    Dim d, c, sName, sBase, n, parts
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    ReDim parts(nCols - 1)

    For c = 1 To nCols
        sBase = Clean(Trim(CStr(arrHdr(1, c) & "")), sDelim)
        If sBase = "" Then sBase = "Column" & c

        sName = sBase
        n = 1
        Do While d.Exists(sName)
            n = n + 1
            sName = sBase & "_" & n
        Loop
        d.Add sName, True
        parts(c - 1) = sName
    Next

    BuildHeaderLine = Join(parts, sDelim)

End Function


Function BuildDataLines(arrData, nCols, sDelim, ByRef nRows)

    Dim r, c, parts, buf, nBuf, bEmpty
    ReDim parts(nCols - 1)
    ReDim buf(999)
    nBuf = 0
    nRows = 0

    For r = 1 To UBound(arrData, 1)
        bEmpty = True
        For c = 1 To nCols
            parts(c - 1) = FormatCell(arrData(r, c), sDelim)
            If bEmpty And parts(c - 1) <> "" Then bEmpty = False
        Next

        If Not bEmpty Then
            If nBuf > UBound(buf) Then ReDim Preserve buf(nBuf * 2)
            buf(nBuf) = Join(parts, sDelim)
            nBuf = nBuf + 1
            nRows = nRows + 1
        End If
    Next

    If nBuf = 0 Then
        BuildDataLines = ""
    Else
        ReDim Preserve buf(nBuf - 1)
        BuildDataLines = Join(buf, vbCrLf) & vbCrLf
    End If

End Function

Function FormatCell(v, sDelim)

    If VarType(v) = vbError Then
        FormatCell = "#ERR"
    ElseIf IsNull(v) Or IsEmpty(v) Then
        FormatCell = ""
    ElseIf VarType(v) = vbBoolean Then
        If v Then FormatCell = "TRUE" Else FormatCell = "FALSE"
    ElseIf IsDate(v) Then
        FormatCell = Year(v) & "-" & P2(Month(v)) & "-" & P2(Day(v))
    ElseIf IsNumeric(v) And VarType(v) <> vbString Then
        FormatCell = Replace(CStr(v), ",", ".")
    Else
        FormatCell = Clean(CStr(v), sDelim)
    End If

End Function

Function Clean(s, sDelim)

    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    s = Replace(s, vbTab, " ")
    If sDelim <> "" Then s = Replace(s, sDelim, "/")
    Clean = Trim(s)

End Function

Function OneCellArray(v)

    Dim a
    ReDim a(1, 1)
    a(1, 1) = v
    OneCellArray = a

End Function


Function ReadUtf8(filePath)

    Dim stream
    ReadUtf8 = ""

    If Not fso.FileExists(filePath) Then
        WriteLog "ReadUtf8: file not found - " & filePath
        Exit Function
    End If

    On Error Resume Next

    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.LoadFromFile filePath
    ReadUtf8 = stream.ReadText
    stream.Close

    If Err.Number <> 0 Then
        WriteLog "ReadUtf8: failed reading " & filePath & " - " & Err.Description
        ReadUtf8 = ""
        Err.Clear
    End If

    On Error GoTo 0
    Set stream = Nothing

End Function


Function WriteUtf8(filePath, fileText)

    Dim stream
    WriteUtf8 = False

    On Error Resume Next
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText fileText
    stream.SaveToFile filePath, 2
    stream.Close

     If Err.Number <> 0 Then
        WriteLog "WriteUtf8: failed writing " & filePath & " - " & Err.Description
        Err.Clear
        On Error GoTo 0
        Set stream = Nothing
        Exit Function
     End If
    On Error GoTo 0
    
    Set stream = Nothing
    WriteUtf8 = True

End Function
 
Sub CleanUpExcel(ByRef dvsXl, ByRef dvsWb)

    On Error Resume Next
    If IsObject(dvsWb) Then
        If Not dvsWb Is Nothing Then dvsWb.Close False
    End If
    If IsObject(dvsXl) Then
        If Not dvsXl Is Nothing Then
           dvsXl.EnableEvents = True
           dvsXl.ScreenUpdating = True
           dvsXl.Quit
        End If
    End If
    Set dvsWb = Nothing
    Set dvsXl = Nothing
    Err.Clear
    On Error GoTo 0
End Sub

Sub KillProcess(processName)

    Dim objShell, command

    Set objShell = CreateObject("WScript.Shell")
    command = "taskkill /F /IM " & processName
    objShell.Run command, 0, True

    Set objShell = Nothing

End Sub

Public Sub WriteLog(logText)

    Const ForAppending = 8
    Dim strLogLevel
    Dim objFSO, objLogFile
    Dim scriptName, logFileName, fullLogPath, timeFormat, logLine, LogFolderPath

    Set localShell = CreateObject("WScript.Shell")
    LogFolderPath = localShell.CurrentDirectory & "\"

      If IsEmpty(logLevel) Or Trim(logLevel) = "" Then
          strLogLevel = "INFO" ' Default
      Else
          strLogLevel = UCase(Trim(logLevel))
      End If

    On Error Resume Next
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    scriptName = objFSO.GetBaseName(WScript.ScriptName)
    logFileName = scriptName & "_logs_" & Year(Date) & "-" & Right("0" & Month(Date), 2) & "-" & Right("0" & Day(Date), 2) & ".log"
    fullLogPath = objFSO.BuildPath(LogFolderPath, logFileName)
    timeFormat = TimeStamp()
    logLine = timeFormat & " [" & strLogLevel & "] " & logText

    Set objLogFile = objFSO.OpenTextFile(fullLogPath, ForAppending, True)
      If Err.Number <> 0 Then
          WScript.Echo "ERROR: Could not open or create log file: " & fullLogPath
      Else
          objLogFile.WriteLine logLine
          objLogFile.Close
      End If
    On Error GoTo 0

    Set objLogFile = Nothing
    Set objFSO = Nothing
    Set localShell = nothing

End Sub

Function TimeStamp()

    Dim d
    d = Now
    TimeStamp = Year(d) & P2(Month(d)) & P2(Day(d)) & "_" & P2(Hour(d)) & P2(Minute(d)) & P2(Second(d))

End Function

Function P2(in_number)

    P2 = Right("0" & in_number, 2)

End Function

Function ConvertCurrency(inputCurrency, targetCurrency, amount, fxDict)

    Dim lookupKey, rate, value
    ConvertCurrency = Null
    value = ToNumber(amount)
     If IsNull(value) Then
        WriteLog "ConvertCurrency: unusable amount '" & amount & "'"
        Exit Function
     End If

    If UCase(Trim(inputCurrency)) = UCase(Trim(targetCurrency)) Then
        ConvertCurrency = value
        Exit Function
    End If

    lookupKey = UCase(Trim(inputCurrency)) & UCase(Trim(targetCurrency))
     If Not fxDict.Exists(lookupKey) Then
        WriteLog "ConvertCurrency: no rate for " & Trim(inputCurrency) & " -> " & Trim(targetCurrency)
        Exit Function
     End If

    rate = ToNumber(fxDict(lookupKey))
     If IsNull(rate) Then
        WriteLog "ConvertCurrency: unusable rate for " & lookupKey
        Exit Function
     End If

    ConvertCurrency = Round(value * rate, 2)

End Function

Function ToNumber(textValue)

    Dim cleanValue, localeDecimal
    ToNumber = Null

    cleanValue = Trim(textValue)
    If cleanValue = "" Then Exit Function

    localeDecimal = Mid(CStr(1.5), 2, 1)          ' this machine's separator
    If localeDecimal <> "." Then cleanValue = Replace(cleanValue, ".", localeDecimal)
    If Not IsNumeric(cleanValue) Then Exit Function

    ToNumber = CDbl(cleanValue)

End Function

Function ReadSheet(dvsWs, headerRow, firstDataRow, lastColumn, keyColumns, valueColumn, sheetLabel)
 
    Dim result, rowDict, arrHdr, arrData, headerNames, usedNames, keyParts
    Dim lastRow, baseName, keyValue, rowText, valueIndex
    Dim n, r, c, i
    Set ReadSheet = Nothing
 
    arrHdr = dvsWs.Range(dvsWs.Cells(headerRow, 1), dvsWs.Cells(headerRow, lastColumn)).Value
     If Err.Number <> 0 Then
        WriteLog "Sheet '" & sheetLabel & "': could not read header row " & headerRow & " - " & Err.Description
        Err.Clear
        Exit Function
     End If
    If Not IsArray(arrHdr) Then arrHdr = OneCellArray(arrHdr)
    Set usedNames = CreateObject("Scripting.Dictionary")
    usedNames.CompareMode = vbTextCompare
    ReDim headerNames(lastColumn - 1)
    
     For c = 1 To lastColumn
        baseName = Trim(FormatCell(arrHdr(1, c), "|"))
        If baseName = "" Then baseName = "Column" & c
        keyValue = baseName
        n = 1

        Do While usedNames.Exists(keyValue)
            n = n + 1
            keyValue = baseName & "_" & n
        Loop

        usedNames.Add keyValue, True
        headerNames(c - 1) = keyValue
     Next
 
    If keyColumns = "" Then
        keyParts = Array()
    Else
        keyParts = Split(keyColumns, "|")
        For i = 0 To UBound(keyParts)
            If Not usedNames.Exists(keyParts(i)) Then
                WriteLog "Sheet '" & sheetLabel & "': key column '" & keyParts(i) & "' not found. Columns: " & Join(headerNames, ", ")
                Exit Function
            End If
        Next
    End If
 
    valueIndex = -1
    If valueColumn <> "" Then
        For c = 0 To lastColumn - 1
            If StrComp(headerNames(c), valueColumn, vbTextCompare) = 0 Then valueIndex = c
        Next
        If valueIndex = -1 Then
            WriteLog "Sheet '" & sheetLabel & "': value column '" & valueColumn & "' not found. Columns: " & Join(headerNames, ", ")
            Exit Function
        End If
    End If
 
    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare
    Set ReadSheet = result
 
    lastRow = dvsWs.Cells(dvsWs.Rows.Count, 1).End(-4162).Row
    If Err.Number <> 0 Then
        WriteLog "Sheet '" & sheetLabel & "': could not find the last row - " & Err.Description
        Err.Clear
        Exit Function
    End If
    If lastRow < firstDataRow Then
        WriteLog "Sheet '" & sheetLabel & "': no data below row " & (firstDataRow - 1)
        Exit Function
    End If
 
    arrData = dvsWs.Range(dvsWs.Cells(firstDataRow, 1), dvsWs.Cells(lastRow, lastColumn)).Value
    If Err.Number <> 0 Then
        WriteLog "Sheet '" & sheetLabel & "': read failed - " & Err.Description
        Err.Clear
        Exit Function
    End If
    If Not IsArray(arrData) Then arrData = OneCellArray(arrData)
 
    ' rows
    n = 0
    For r = 1 To UBound(arrData, 1)
 
        Set rowDict = CreateObject("Scripting.Dictionary")
        rowDict.CompareMode = vbTextCompare
 
        rowText = ""
        For c = 1 To lastColumn
            rowDict(headerNames(c - 1)) = FormatCell(arrData(r, c), "|")
            rowText = rowText & rowDict(headerNames(c - 1))
        Next
 
        If Trim(rowText) <> "" Then
 
            If UBound(keyParts) < 0 Then
                n = n + 1
                keyValue = CStr(n)
            Else
                keyValue = ""
                For i = 0 To UBound(keyParts)
                    keyValue = keyValue & UCase(Trim(rowDict(keyParts(i))))
                Next
            End If
 
            If keyValue = "" Then
                WriteLog "Sheet '" & sheetLabel & "' row " & (r + firstDataRow - 1) & ": blank key, row skipped"
            Else
                If result.Exists(keyValue) Then
                    WriteLog "Sheet '" & sheetLabel & "' row " & (r + firstDataRow - 1) & ": duplicate key '" & keyValue & "'"
                End If
 
                If valueIndex = -1 Then
                    Set result(keyValue) = rowDict 
                Else
                    result(keyValue) = rowDict(headerNames(valueIndex))
                End If
            End If
        End If
    Next
 
End Function

Function UniqueKey(targetDict, baseKey)
    Dim candidate, n
    candidate = baseKey
    n = 1
    Do While targetDict.Exists(candidate)
        n = n + 1
        candidate = baseKey & Chr(1) & "#" & n
    Loop
    UniqueKey = candidate
End Function

Function FindSingleXlsx(folderPath, folderLabel)

    Dim oneFile, foundPath, foundCount, ext
    FindSingleXlsx = ""
    foundPath  = ""
    foundCount = 0
     If Not fso.FolderExists(folderPath) Then
        WriteLog "Folder not found: " & folderPath
        Exit Function
     End If

    For Each oneFile In fso.GetFolder(folderPath).Files
        ext = LCase(fso.GetExtensionName(oneFile.Name))
        If (ext = "xlsx" Or ext = "xlsm") And Left(oneFile.Name, 2) <> "~$" Then
            foundCount = foundCount + 1
            foundPath  = oneFile.Path
        End If
    Next

    If foundCount = 0 Then
        WriteLog folderLabel & ": no xlsx file found in " & folderPath
    ElseIf foundCount > 1 Then
        WriteLog folderLabel & ": " & foundCount & " xlsx files found in " & folderPath &  " - exactly one is required"
        foundPath = ""
    Else
        WriteLog folderLabel & ": using " & fso.GetFileName(foundPath)
    End If

    FindSingleXlsx = foundPath

End Function

Sub CleanArchive(folderPath, filePrefix, keepCount)

    Dim oneFile, names, count, i, j, swap, deleted

    If Not fso.FolderExists(folderPath) Then Exit Sub

    count = 0
    ReDim names(255)

    For Each oneFile In fso.GetFolder(folderPath).Files
        If StrComp(Left(oneFile.Name, Len(filePrefix)), filePrefix, vbTextCompare) = 0 Then
            If count > UBound(names) Then ReDim Preserve names(count * 2)
            names(count) = oneFile.Name
            count = count + 1
        End If
    Next

    If count <= keepCount Then Exit Sub
    ReDim Preserve names(count - 1)

    ' descending, so the newest land at the front
    For i = 0 To count - 2
        For j = 0 To count - 2 - i
            If names(j) < names(j + 1) Then
                swap = names(j) : names(j) = names(j + 1) : names(j + 1) = swap
            End If
        Next
    Next

    deleted = 0
    On Error Resume Next
    For i = keepCount To count - 1
        fso.DeleteFile fso.BuildPath(folderPath, names(i)), True
        If Err.Number = 0 Then
            deleted = deleted + 1
        Else
            WriteLog "Could not delete " & names(i) & ": " & Err.Description
            Err.Clear
        End If
    Next
    On Error GoTo 0

    WriteLog "Archive cleaned: kept " & keepCount & ", deleted " & deleted

End Sub

'#####################################################################################################################################
