KillProcess("excel.exe")

Dim shell, rootFolder, masterCsvFile, in_Xlsx_file, out_Csv_file
Dim fso, inputFolder, archiveFolder, outputFolder, delimiter
'shareFolder = "P:\!04_Sales\"                                   '  Share folder location
Set shell = CreateObject("WScript.Shell")
rootFolder = shell.CurrentDirectory & "\"
inputFolder = rootFolder & "Input"
archiveFolder = rootFolder & "Archive"
outputFolder = rootFolder & "Output"
masterCsvFile = rootFolder & "MasterData.csv"
masterSrcFileName = "Base SIOP test"
masterSrcFileExt = ".xlsx"
masterSrcFilePath = inputFolder & "\" & masterSrcFileName & masterSrcFileExt
destArchFileLoc = archiveFolder & "\" & masterSrcFileName & "_" & TimeStamp() & ".xlsx"

WriteLog "Process started."
Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(inputFolder) then 
        fso.CreateFolder(inputFolder)
        WScript.Echo "Check if input files are placed in 'Input' folder"
        WScript.Quit
    End if

    If Not fso.FolderExists(archiveFolder) then 
        fso.CreateFolder(archiveFolder)
        WriteLog "Created folder " & archiveFolder
    End if

    If Not fso.FolderExists(outputFolder) then 
        fso.CreateFolder(outputFolder)
        WriteLog "Created folder " & outputFolder
    End if

    On Error Resume Next
    fso.CopyFile masterSrcFilePath, destArchFileLoc, False
     If Err.Number <> 0 Then
        WriteLog "Backup failed: " & Err.Description
        Err.Clear
        On Error GoTo 0
        MsgBox "Could not back up the master file. Process stopped.", vbCritical, "Sales Analysis"
        WScript.Quit 1
     End If
    On Error GoTo 0
'############################################# Load all Input.xlsx sheets into Dictionary #############################################

Dim sheet_map
Dim inputBook, dvsXl, dvsWb, dvsWs
Dim sheetNames, inputData, sheetData, missingSheets, sheetName, i
Const map_SHEET = 0
Const map_HDRROW = 1
Const map_FIRSTROW = 2
Const map_LASTCOL = 3
Const map_KEYCOLS = 4
Const map_VALUECOL = 5

sheet_map = Array(Array("FX", 1, 2,  2, "Currency", "Value"), _
                  Array("OrderType", 1, 2,  2, "Spares customers", "Column2"), _
                  Array("AOP", 1, 2,  4, "Month|Year", ""), _
                  Array("NRCs", 1, 2,  7, "", ""), _ 
                  Array("OtherForcast", 1, 2, 10, "", ""), _
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


'############################################# Convert XLSX to CSV - Customer Order Lines Export file #############################################
in_Xlsx_file = inputFolder & "\CustomerOrderLinesExport33.xlsx"
out_Csv_file = outputFolder & "\CustomerOrderLinesExport.csv"
delimiter = "|"
'XlsxToCsv in_Xlsx_file, 1, 1, 1, 0, delimiter, out_Csv_file
If XlsxToCsv(in_Xlsx_file, 1, 1, 1, 0, delimiter, out_Csv_file) < 0 Then
    MsgBox "Could not convert CustomerOrderLinesExport. See log.", vbCritical, "Sales Analysis"
    WScript.Quit 1
End If

Dim colOrder, colLine, colStatus, colItem1, colItem2, colQtyOrdered, colUM
Dim colUnitPrice, colOrderDate, colDueDate, colInvoiced, colName, colNetPrice, colCurrency
Dim fileText, lines, headerFields, headerIndex, columnCount, requiredCols
Dim orderLinesDict, rowFields, newRow, masterRow, keyValue
Dim idxOrder, idxLine, idxStatus, idxItem1, idxItem2, idxQtyOrdered, idxUM
Dim idxUnitPrice, idxOrderDate, idxDueDate, idxInvoiced, idxName, idxNetPrice, idxCurrency
Dim yearFrom, yearTo, dueYear, dueMonth, dateValue, statusValue, monthYear
Dim qtyOrdered, invoiced, netPrice, converted, lookupValue
Dim fxDict, orderTypeDict, aopDict, productLineDict
Dim rowsKept, rowsOutOfWindow, rowsBadStatus, rowsPlannedZero, rowsDuplicate
Dim rowsNoFY, rowsNoProductLine, rowsNoRate, j, masterHeader

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
masterHeader = Array("Order", "Line", "Status (SL)", "Status SIOP", "Order + Pos", "Item", "Item_2", "Qty Ordered", "U/M", "Unit Price", "Order Date", "Due Date", _
                    "Month", "Year", "M+Y", "FY", "Name", "Order type", "Net Price", "Currency", "Net Price EUR", "Net Price USD", "Product Line", "Qty Difference")

Const M_ORDER = 0      : Const M_LINE = 1       : Const M_STATUS = 2
Const M_STATUSSIOP = 3 : Const M_ORDERPOS = 4   : Const M_ITEM1 = 5
Const M_ITEM2 = 6      : Const M_QTY = 7        : Const M_UM = 8
Const M_UNITPRICE = 9  : Const M_ORDERDATE = 10 : Const M_DUEDATE = 11
Const M_MONTH = 12     : Const M_YEAR = 13      : Const M_MY = 14
Const M_FY = 15        : Const M_NAME = 16      : Const M_ORDERTYPE = 17
Const M_NETPRICE = 18  : Const M_CURRENCY = 19  : Const M_NPEUR = 20
Const M_NPUSD = 21     : Const M_PRODLINE = 22  : Const M_QTYDIFF = 23

Set fxDict = inputData("FX")
Set orderTypeDict = inputData("OrderType")
Set aopDict = inputData("AOP")
Set productLineDict = inputData("ProductLine")

fileText = ReadUtf8(out_Csv_file)
If Trim(fileText) = "" Then
    WriteLog "CustOrderLines CSV is empty or unreadable: " & out_Csv_file
    MsgBox "The converted CustomerOrderLines CSV is empty.", vbCritical, "Sales Analysis"
    WScript.Quit 1
End If

fileText = Replace(Replace(fileText, vbCrLf, vbLf), vbCr, vbLf)
lines = Split(fileText, vbLf)
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
        MsgBox "CustomerOrderLines is missing column '" & requiredCols(i) & "'.", _
               vbCritical, "Sales Analysis"
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
rowsKept = 0 : rowsOutOfWindow = 0 : rowsBadStatus = 0 : rowsPlannedZero = 0
rowsDuplicate = 0 : rowsNoFY = 0 : rowsNoProductLine = 0 : rowsNoRate = 0

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
            masterRow(j) = ""              ' unmapped columns stay blank, never 0
        Next

        masterRow(M_ORDER) = newRow(idxOrder)
        masterRow(M_LINE) = newRow(idxLine)
        masterRow(M_STATUS) = newRow(idxStatus)
        masterRow(M_STATUSSIOP) = "Orderbook"
        masterRow(M_ORDERPOS) = newRow(idxOrder) & newRow(idxLine)
        masterRow(M_ITEM1) = newRow(idxItem1)
        masterRow(M_ITEM2) = newRow(idxItem2)
        masterRow(M_QTY) = newRow(idxQtyOrdered)
        masterRow(M_UM) = newRow(idxUM)
        masterRow(M_UNITPRICE) = newRow(idxUnitPrice)
        masterRow(M_ORDERDATE) = newRow(idxOrderDate)
        masterRow(M_DUEDATE) = dateValue
        masterRow(M_NAME) = newRow(idxName)
        masterRow(M_NETPRICE) = newRow(idxNetPrice)
        masterRow(M_CURRENCY) = newRow(idxCurrency)
        dueMonth  = CLng(Mid(dateValue, 6, 2))        ' "07" -> 7
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


bjorn file now!































































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

    ' One array read. Cell-by-cell COM over 150k cells takes minutes.
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

    sOut = BuildHeaderLine(arrHdr, nCols, sDelim) & vbCrLf & _
           BuildDataLines(arrData, nCols, sDelim, nRows)

    If Not WriteUtf8(dstCsv, sOut) Then Exit Function

   XlsxToCsv = nRows

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

        If Not bEmpty Then                    ' drop UsedRange's trailing blanks
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

    If VarType(v) = vbError Then          ' MUST be first
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
    stream.Type = 2                    ' adTypeText - treat contents as text, not bytes
    stream.Charset = "utf-8"
    stream.Open                        ' must Open before any read
    stream.LoadFromFile filePath       ' pull the file into the stream buffer
    ReadUtf8 = stream.ReadText         ' whole buffer as one string
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
    stream.WriteText fileText          ' whole string into the buffer
    stream.SaveToFile filePath, 2      ' 2 = adSaveCreateOverWrite
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
    TimeStamp = Year(d) & P2(Month(d)) & P2(Day(d)) & "_" & _
                P2(Hour(d)) & P2(Minute(d)) & P2(Second(d))

End Function

Function P2(in_number)

    P2 = Right("0" & in_number, 2)

End Function

'------------------------------------------------------------------------------
' Converts an amount between currencies using the FX rates already loaded
' from Input.xlsx. No Excel, no file access.
'
'   fxDict - pass inputData("FX")
'
' Returns a Double, or Null if the amount or the rate is unusable.
'------------------------------------------------------------------------------
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

    lookupKey = UCase(Trim(inputCurrency)) & " TO " & UCase(Trim(targetCurrency))
     If Not fxDict.Exists(lookupKey) Then
        WriteLog "ConvertCurrency: no rate for '" & lookupKey & "'"
        Exit Function
     End If

    rate = ToNumber(fxDict(lookupKey))
     If IsNull(rate) Then
        WriteLog "ConvertCurrency: unusable rate for '" & lookupKey & "'"
        Exit Function
     End If

    ConvertCurrency = Round(value * rate, 2)

End Function

'------------------------------------------------------------------------------
' Converts a text field to a number.
' CDbl follows the machine's regional decimal separator: on a Croatian Windows
' CDbl("1005.97") throws or returns 100597. Values always use a period, so
' swap it for whatever this machine expects before converting.
' Returns Null for unusable values ("", "#ERR", text).
'------------------------------------------------------------------------------
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

'------------------------------------------------------------------------------
' Reads one sheet into a dictionary.
'
'   valueColumn = ""   ->  key -> Dictionary(columnName -> value)
'   valueColumn = name ->  key -> single value  (flat lookup)
'   keyColumns  = ""   ->  key is the row's position, "1", "2", "3"...
'
' Returns Nothing on a structural problem so the caller can abort.
'------------------------------------------------------------------------------
Function ReadSheet(dvsWs, headerRow, firstDataRow, lastColumn, keyColumns, valueColumn, sheetLabel)
 
    Dim result, rowDict, arrHdr, arrData, headerNames, usedNames, keyParts
    Dim lastRow, baseName, keyValue, rowText, valueIndex
    Dim n, r, c, i
    Set ReadSheet = Nothing
 
    arrHdr = dvsWs.Range(dvsWs.Cells(headerRow, 1), dvsWs.Cells(headerRow, lastColumn)).Value
     If Err.Number <> 0 Then
        WriteLog "Sheet '" & sheetLabel & "': could not read header row " & headerRow & _
                 " - " & Err.Description
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
                WriteLog "Sheet '" & sheetLabel & "': key column '" & keyParts(i) & _
                         "' not found. Columns: " & Join(headerNames, ", ")
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
            WriteLog "Sheet '" & sheetLabel & "': value column '" & valueColumn & _
                     "' not found. Columns: " & Join(headerNames, ", ")
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

'#####################################################################################################################################
