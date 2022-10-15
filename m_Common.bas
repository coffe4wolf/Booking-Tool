Attribute VB_Name = "m_Common"
Option Compare Text
Option Explicit

Const Debug_Mode                As Boolean = False

Function ConvertToLetter(iCol As Long) As String                                                                'converting values in current columnt
                                                                                                                'into letters
    If iCol > 0 And iCol <= Columns.Count Then ConvertToLetter = Replace(Cells(1, iCol).Address(0, 0), 1, "")   'if
    
End Function

Function SheetExists(ByVal shtName As String, Optional wbObj As Workbook) As Boolean                            'testing sheets in book for existing
                                                                                                                '
    Dim sht As Worksheet                                                                                        '
    
    If wbObj Is Nothing Then Set wbObj = ThisWorkbook                                                           'if we havn't active book - take current as active
    
    On Error Resume Next
        Set sht = wbObj.Sheets(shtName)                                                                         'taking sheet from func argument
    On Error GoTo 0
    
    SheetExists = Not sht Is Nothing
    
End Function
Function SheetsCountVisible(Optional wbObj As Workbook) As Long
    
    Dim sht As Worksheet
    
    If wbObj Is Nothing Then Set wbObj = ThisWorkbook
    
    SheetsCountVisible = 0
    For Each sht In wbObj.Sheets
        If sht.Visible Then SheetsCountVisible = SheetsCountVisible + 1
    Next sht

End Function

Sub RecreateSheet(ByVal shtNameToRecreate As String, Optional shtNameInsertAfter As String, Optional wbObj As Workbook)     'recreate sheet
                                                                                                                            
    If wbObj Is Nothing Then Set wbObj = ThisWorkbook
    
    Select Case SheetExists(shtNameToRecreate, wbObj)
        Case True
        
            wbObj.Worksheets.Add(After:=wbObj.Sheets(wbObj.Sheets.Count)).Name = "TmpSheetName"
            wbObj.Sheets(shtNameToRecreate).Delete
            wbObj.Sheets("TmpSheetName").Name = shtNameToRecreate
                        
        Case False
        
            If shtNameInsertAfter = "" Then
                wbObj.Worksheets.Add(After:=wbObj.Sheets(wbObj.Sheets.Count)).Name = shtNameToRecreate
            Else
                If SheetExists(shtNameInsertAfter, wbObj) = True Then
                    wbObj.Worksheets.Add(After:=wbObj.Sheets(shtNameInsertAfter)).Name = shtNameToRecreate
                Else
                    wbObj.Worksheets.Add(After:=wbObj.Sheets(wbObj.Sheets.Count)).Name = shtNameToRecreate
                End If
            End If
    
    End Select

End Sub

Function FileExists(FilePath As String) As Boolean                                                                          'checking file for existing

FileExists = False

If FilePath <> "" Then
    On Error Resume Next
        If Dir(FilePath) <> "" Then FileExists = True
    On Error GoTo 0
End If

End Function

Function GetBorders(LRorLC As String, shtName As String, Optional wbObj As Workbook) As Long

    Dim sht As Worksheet
    
    
    Dim rLastColumnCell As Range
    Dim rLastRowCell As Range
    
    Dim LastColumn As Long
    Dim lastRow As Long
    
    Dim counter As Long
    
    If wbObj Is Nothing Then Set wbObj = ThisWorkbook
    
    On Error Resume Next
        Set sht = wbObj.Sheets(shtName)
    On Error GoTo 0
    
    Set rLastColumnCell = sht.Cells.Find(What:="*", After:=sht.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)
    Set rLastRowCell = sht.Cells.Find(What:="*", After:=sht.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False)
    
    If rLastColumnCell Is Nothing Then LastColumn = 1 Else LastColumn = rLastColumnCell.Column
    If rLastRowCell Is Nothing Then lastRow = 1 Else lastRow = rLastRowCell.Row
    
    If lastRow = 1 Then
        For counter = 1 To 3
            If Trim(sht.Cells(counter, 1).value) <> "" Then
                lastRow = counter
            End If
        Next counter
    End If
    
    Select Case LRorLC
        Case "LR"
            GetBorders = lastRow
        Case "LC"
            GetBorders = LastColumn
        Case Else
            GetBorders = 0
    End Select
    
End Function

Function ImprovePerformance(TrueFalse As Boolean)

    Select Case TrueFalse
        Case True
            Application.ScreenUpdating = False
            Application.DisplayAlerts = False
            Application.AskToUpdateLinks = False
            Application.Calculation = xlCalculationManual
        Case False
            Application.ScreenUpdating = True
            Application.DisplayAlerts = True
            Application.AskToUpdateLinks = True
            Application.Calculation = xlCalculationAutomatic
    End Select

End Function

Function NumberToMSSQLFormat(ByVal NumberToConvert) As String

NumberToConvert = LTrim(RTrim(NumberToConvert))

If RxTest(NumberToConvert, "\d+([\.,]\d+)?") = False Then
    NumberToMSSQLFormat = "null"
Else
    If InStr(1, CStr(NumberToConvert), ".", vbTextCompare) > 0 Then
        NumberToMSSQLFormat = NumberToConvert
    Else
        If InStr(1, CStr(NumberToConvert), ",", vbTextCompare) > 0 Then
            NumberToMSSQLFormat = Replace(CStr(NumberToConvert), ",", ".")
        Else
            NumberToMSSQLFormat = CStr(NumberToConvert)
        End If
    End If
End If

End Function

Function StringToMSSQLFormat(ByVal StringToConvert, Optional ToUnicode As Boolean = False) As String

StringToConvert = LTrim(RTrim(StringToConvert))

If StringToConvert = vbNullString Then
    StringToMSSQLFormat = "null"
Else
    StringToMSSQLFormat = StringToConvert
    If InStr(1, StringToMSSQLFormat, "'", vbTextCompare) > 0 Then StringToMSSQLFormat = ReplaceSingleQuote(StringToMSSQLFormat)
    
    If ToUnicode = True Then
        StringToMSSQLFormat = "N" & "'" & StringToMSSQLFormat & "'"
    Else
        StringToMSSQLFormat = "'" & StringToMSSQLFormat & "'"
    End If
End If

End Function

Function ReplaceSingleQuote(ByVal StringToReplaceSingleQuote As String) As String

If Len(StringToReplaceSingleQuote) > 0 Then ReplaceSingleQuote = Replace(StringToReplaceSingleQuote, "'", "''") Else ReplaceSingleQuote = vbNullString

End Function

Function ToSingleQuotes(ByVal StringToSingleQuote As String) As String

If Len(StringToSingleQuote) > 0 Then ToSingleQuotes = "'" & StringToSingleQuote & "'" Else ToSingleQuotes = "''"

End Function

Function ToSquareBracket(ByVal StringToBracket As String) As String

QuoteName = "[" & NameToQuote & "]"

End Function

Function RangeBorders(rngObj As Range, FirstOrLast As String) As String

Dim rngStr As String

If IsMissing(rngObj) = False Then

    rngStr = rngObj.Address(RowAbsolute:=True, ColumnAbsolute:=True)
    rngStr = Replace(rngStr, "$", vbNullString)
    
    Select Case FirstOrLast
           Case "First"
                RangeBorders = RxMatch(rngStr, "[A-Z0-9]+(?=\:)", False, False)
           Case "Last"
                RangeBorders = Replace(RxMatch(rngStr, "\:[A-Z0-9]+", False, False), ":", "")
    End Select
    
End If

End Function

Function ReplaceForbiddenChars(ByVal StringToReplace As String) As String

Dim CharsArray
Dim X As Byte

ReplaceForbiddenChars = StringToReplace

CharsArray = Array("<", ">", "|", "/", "*", "\", ":", "?", """")
For X = LBound(CharsArray) To UBound(CharsArray)
    If CharsArray(X) = ":" Then
        ReplaceForbiddenChars = Replace(ReplaceForbiddenChars, CharsArray(X), "-", 1)
    Else
        ReplaceForbiddenChars = Replace(ReplaceForbiddenChars, CharsArray(X), "_", 1)
    End If
Next X

End Function

Function NullToEmptyString(ObjectToCheck As Variant) As String

NullToEmptyString = vbNullString

On Error Resume Next
    NullToEmptyString = CStr(ObjectToCheck)
On Error GoTo 0

End Function

Function GetLocalDateFormat() As String

Dim sht As Worksheet

On Error Resume Next

    ThisWorkbook.Sheets(shtName).Delete
    ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)).Name = "TempSheetToCheckDateFormat"
    Set sht = ThisWorkbook.Sheets("TempSheetToCheckDateFormat")
    sht.Range("A1").value = "2015-01-01"
    
    If InStr(1, UCase(sht.Range("A1").NumberFormatLocal), "AA", vbTextCompare) > 0 Or InStr(1, UCase(sht.Range("A1").NumberFormat), "AA", vbTextCompare) > 0 Then GetLocalDateFormat = "RU"
    If InStr(1, UCase(sht.Range("A1").NumberFormatLocal), "YY", vbTextCompare) > 0 And InStr(1, UCase(sht.Range("A1").NumberFormat), "YY", vbTextCompare) > 0 Then GetLocalDateFormat = "EN"
    
    sht.Delete
    
On Error GoTo 0

End Function

Function Pivot_Table(shtName_Source As String, shtName_Pivot As String, DestinationCell As String, Optional wbObj As Workbook, Optional ColumnField As String, _
                                                                                                   Optional RowField As String, _
                                                                                                   Optional FilterField As String, _
                                                                                                   Optional ValuesField As String)

Dim WS_Pivot As Worksheet
Dim WS_Source As Worksheet
Dim WS_Pivot_Name As String

Dim i As Long, J As Long
Dim LR_Source As Long
Dim LC_Source As Long
Dim r As Range

Dim SrcData As String
Dim StartPvt As String
Dim pvtCache As PivotCache
Dim pvt As PivotTable

If wbObj Is Nothing Then Set wbObj = ThisWorkbook

On Error Resume Next
    Set WS_Source = wbObj.Sheets(shtName_Source)
On Error GoTo 0

WS_Pivot_Name = shtName_Pivot

If SheetExists(WS_Pivot_Name, wbObj) = False Then wbObj.Worksheets.Add(After:=WS_Source).Name = WS_Pivot_Name
Set WS_Pivot = wbObj.Sheets(WS_Pivot_Name)

LR_Source = GetBorders("LR", WS_Source.Name, wbObj)
LC_Source = GetBorders("LC", WS_Source.Name, wbObj)

'Pivot Table creation

'Determine the data range you want to pivot
SrcData = WS_Source.Name & "!" & WS_Source.Range(WS_Source.Cells(1, 1), WS_Source.Cells(LR_Source, LC_Source)).Address(ReferenceStyle:=xlR1C1)

'Where do you want Pivot Table to start?
StartPvt = WS_Pivot.Name & "!" & WS_Pivot.Range(DestinationCell).Address(ReferenceStyle:=xlR1C1)

'Create Pivot Cache from Source Data
Set pvtCache = wbObj.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=SrcData)

'Create Pivot table from Pivot Cache
Set pvt = pvtCache.CreatePivotTable(TableDestination:=StartPvt, TableName:="PivotTable1")

If ColumnField <> "" Then pvt.PivotFields(ColumnField).Orientation = xlColumnField
If RowField <> "" Then pvt.PivotFields(RowField).Orientation = xlRowField
If FilterField <> "" Then pvt.PivotFields(FilterField).Orientation = xlPageField
If ValuesField <> "" Then
    pvt.AddDataField pvt.PivotFields(ValuesField), "Sum of " & ValuesField, xlSum
    pvt.DataFields("Sum of " & ValuesField).NumberFormat = "#,#0.00"
End If

pvt.ManualUpdate = True

End Function

Sub MakeBeautiful(ByVal shtName As String, Optional wbName As Workbook, Optional AutoFitMode As String)

Dim sht As Worksheet
Dim tmpLR As Long
Dim tmpLC As Long
Dim i As Long, J As Long

If wbName Is Nothing Then Set wbName = ThisWorkbook

On Error Resume Next
    Set sht = wbName.Sheets(shtName)
On Error GoTo 0
    
With sht

    tmpLC = GetBorders("LC", .Name, wbName)
    tmpLR = GetBorders("LR", .Name, wbName)
    
    .Rows(1).HorizontalAlignment = xlLeft
    .Activate
    .Rows(1).Font.Bold = True ' Making headers bold
    
    If tmpLR > 1 Then
        If Not .AutoFilterMode Then .Range("A1").AutoFilter ' Enabling Autofilter
    End If
    
    If AutoFitMode = "Full" Then
        .Range(.Cells(1, 1), .Cells(tmpLR, tmpLC)).Columns.AutoFit
    Else
        .Range(.Cells(1, 1), .Cells(1, tmpLC)).Columns.AutoFit
    End If
    
    ' Format dates
    If tmpLR > 1 Then
        i = 1: Do While i <= tmpLC
            If RxTest(.Cells(2, i).value, "\d{4}\-\d{2}\-\d{2}", True, False) = True Then
                If tmpLR > 1 Then
                    .Range(.Cells(2, i), .Cells(tmpLR, i)).value = .Range(.Cells(2, i), .Cells(tmpLR, i)).value
                    .Range(.Cells(2, i), .Cells(tmpLR, i)).NumberFormat = "yyyy-mm-dd"
                End If
            End If
        i = i + 1: Loop
   End If
   
End With

End Sub

Function TimeDiffInSeconds(TimeStart As Date, TimeEnd As Date) As String

Dim DurationInSeconds As Long

DurationInSeconds = DateDiff("s", TimeStart, TimeEnd)

If DurationInSeconds = 1 Then
    TimeDiffInSeconds = "1 second"
ElseIf DurationInSeconds <= 60 Then
    TimeDiffInSeconds = CStr(DurationInSeconds) & " seconds"
ElseIf DurationInSeconds > 60 Then
    TimeDiffInSeconds = CStr(Fix(DurationInSeconds / 60)) & " minutes " & CStr(DurationInSeconds - Fix(DurationInSeconds / 60) * 60) & " seconds"
Else
    TimeDiffInSeconds = "<Unknown error>"
End If

End Function

Function BrowseForFolder(Optional Comments As String, Optional OpenAt As Variant) As Variant
     'Function purpose:  To Browser for a user selected folder.
     'If the "OpenAt" path is provided, open the browser at that directory
     'NOTE:  If invalid, it will open at the Desktop level

    Dim ShellApp As Object
    
    If Comments = vbNullString Then Comments = "Please choose a folder"
    
     'Create a file browser window at the default folder
    Set ShellApp = CreateObject("Shell.Application"). _
    BrowseForFolder(0, Comments, 0, OpenAt)

     'Set the folder to that selected.  (On error in case cancelled)
    On Error Resume Next
    BrowseForFolder = ShellApp.self.Path
    On Error GoTo 0

     'Destroy the Shell Application
    Set ShellApp = Nothing

     'Check for invalid or non-entries and send to the Invalid error
     'handler if found
     'Valid selections can begin L: (where L is a letter) or
     '\\ (as in \\servername\sharename.  All others are invalid
    Select Case Mid(BrowseForFolder, 2, 1)
    Case Is = ":"
        If Left(BrowseForFolder, 1) = ":" Then GoTo Invalid
    Case Is = "\"
        If Not Left(BrowseForFolder, 1) = "\" Then GoTo Invalid
    Case Else
        GoTo Invalid
    End Select

    Exit Function

Invalid:
     'If it was determined that the selection was invalid, set to False
    BrowseForFolder = vbNullString
End Function

Sub GetUsername()

    If DbUsername = "" Then
        DbUsername = InputBox("Please enter username for accessing the database server. ", "Enter username")
    End If
    
End Sub

Sub GetPwd(DbUsernameLocal As String)
    
    If DbPassword = "" Then
        DbPassword = InputBox("Please enter password for the database user: " & vbCrLf & DbUsernameLocal, "Enter password")
    End If
   
End Sub

Sub RunSQLSelect(shtName As String, sql As String, DbServerAddressLocal As String, _
                                                                             Optional DbNameLocal As String, _
                                                                             Optional wbObj As Workbook, _
                                                                             Optional rangeToInsert As String, _
                                                                             Optional DrawHeader As Boolean = True)
                                                   
Dim sht As Worksheet
Dim rangeToInsertData As String
Dim i As Long, J As Long, RowsCount As Long
Dim iCols As Variant
Dim conn  As ADODB.Connection
Dim rs  As ADODB.Recordset



If wbObj Is Nothing Then Set wbObj = ThisWorkbook

If rangeToInsert = vbNullString Then
    rangeToInsert = "A1": rangeToInsertData = "A2"
Else
    i = RxMatch(rangeToInsert, "\d+", True, False)
    rangeToInsertData = Replace(rangeToInsert, i, i + 1)
End If

If SheetExists(shtName, wbObj) = False Then Call RecreateSheet(shtName, , wbObj)
Set sht = wbObj.Sheets(shtName)

Set conn = New ADODB.Connection
conn.CursorLocation = adUseClient


conn.ConnectionString = "Provider=SQLOLEDB;Data Source=" & DbServerAddressLocal & ";Trusted_connection=yes;Database=" & DbNameLocal

conn.Open
conn.CommandTimeout = 0

Set rs = New ADODB.Recordset
rs.Open sql, conn
RowsCount = rs.RecordCount

If DrawHeader = True Then

    ' Making row with headers
    For iCols = 0 To rs.Fields.Count - 1
        i = sht.Range(rangeToInsert).Row
        J = sht.Range(rangeToInsert).Column
        sht.Cells(i, J + iCols).value = rs.Fields(iCols).Name
    Next
    
    ' Inserting data
    sht.Range(rangeToInsertData).CopyFromRecordset rs
Else
    ' Inserting data
    sht.Range(rangeToInsertData).Offset(-1, 0).CopyFromRecordset rs
End If

' Inserting data
sht.Range(rangeToInsertData).CopyFromRecordset rs

rs.Close
conn.Close

If (RowsCount > 1048576) Then
    MsgBox ("Too much rows for import to Excel. Not all data was imported.")
End If

Oops:
    Select Case Err
    
        Case -2147467259:
            If (InStr(1, Err.Description, "ORA-01017") > 0) Then
                MsgBox ("Неверное имя пользователя или пароль.")
                Exit Sub
            ElseIf (InStr(1, Err.Description, "ORA-12154") > 0) Then
                MsgBox ("Не удаётся установить соединение по указанному адресу.")
                Exit Sub
            ElseIf (InStr(1, Err.Description, "ORA-00904") > 0) Then
                MsgBox ("Одно или более из запрашиваемых полей отсутствует в таблице.")
                Exit Sub
            ElseIf (InStr(1, Err.Description, "ORA-00942") > 0) Then
                MsgBox ("Запрашиваемая таблица или представление не существует.")
                Exit Sub
            ElseIf (InStr(1, Err.Description, "ORA-00936") > 0) Then
                MsgBox ("Не задано условие предиката.")
                Exit Sub
            ElseIf (InStr(1, Err.Description, "ORA-00920") > 0) Then
                MsgBox ("Условие предиката не корректно.")
                Exit Sub
            ElseIf (InStr(1, Err.Description, "Cannot open database") > 0) Then
                MsgBox ("Невозможно открыть базу данных. Проверьте правильность ввода названия базы данных.")
                Exit Sub
            Else
                MsgBox ("Непредвиденная ошибка. " & Err & ", " & Err.Description)
                Exit Sub
            End If
            
        Case -2147217873:
            If (InStr(1, Err.Description, "ORA-00001") > 0) Then
                MsgBox ("Нарушение уникального ключа. Возможно вы вставляете запись туда, где уже существует значение.")
            Else
                MsgBox ("Чёта нехорошее случилось." & Err.Number & ", " & Err.Description)
            End If
            Exit Sub
            
        Case -2147217900:
            MsgBox ("Некорректный SQL-запрос.")
            Exit Sub
            
        Case -2147217865:
            If (InStr(1, Err.Description, "ORA-00942") > 0) Then
                MsgBox ("Запрашиваемая таблица не существует.")
            End If
            Exit Sub
            
        Case -2147217843:
            MsgBox ("Неверное имя пользователя или пароль.")
            Exit Sub
            
        Case 0:
        
        Case Else:
            MsgBox (Err & ", " & Err.Description)
            
    End Select

End Sub

Sub RunSQLInsert(sql As String, DbServerAddressLocal As String, Optional DbNameLocal As String)
                                                   
Dim sht As Worksheet
Dim conn  As ADODB.Connection


Set conn = New ADODB.Connection
conn.CursorLocation = adUseClient

conn.ConnectionString = "Provider=SQLOLEDB;Data Source=" & DbServerAddressLocal & ";Trusted_connection=yes;Database=" & DbNameLocal

conn.Open
conn.CommandTimeout = 0

conn.Execute (sql)

conn.Close

Oops:
    Select Case Err
    
        Case -2147467259:
            If (InStr(1, Err.Description, "ORA-01017") > 0) Then
                MsgBox ("Неверное имя пользователя или пароль.")
                Exit Sub
            ElseIf (InStr(1, Err.Description, "ORA-12154") > 0) Then
                MsgBox ("Не удаётся установить соединение по указанному адресу.")
                Exit Sub
            ElseIf (InStr(1, Err.Description, "ORA-00904") > 0) Then
                MsgBox ("Одно или более из запрашиваемых полей отсутствует в таблице.")
                Exit Sub
            ElseIf (InStr(1, Err.Description, "ORA-00942") > 0) Then
                MsgBox ("Запрашиваемая таблица или представление не существует.")
                Exit Sub
            ElseIf (InStr(1, Err.Description, "ORA-00936") > 0) Then
                MsgBox ("Не задано условие предиката.")
                Exit Sub
            ElseIf (InStr(1, Err.Description, "ORA-00920") > 0) Then
                MsgBox ("Условие предиката не корректно.")
                Exit Sub
            ElseIf (InStr(1, Err.Description, "Cannot open database") > 0) Then
                MsgBox ("Невозможно открыть базу данных. Проверьте правильность ввода названия базы данных.")
                Exit Sub
            Else
                MsgBox ("Непредвиденная ошибка. " & Err & ", " & Err.Description)
                Exit Sub
            End If
            
        Case -2147217873:
            If (InStr(1, Err.Description, "ORA-00001") > 0) Then
                MsgBox ("Нарушение уникального ключа. Возможно вы вставляете запись туда, где уже существует значение.")
            Else
                MsgBox ("Чёта нехорошее случилось." & Err.Number & ", " & Err.Description)
            End If
            Exit Sub
            
        Case -2147217900:
            MsgBox ("Некорректный SQL-запрос.")
            Exit Sub
            
        Case -2147217865:
            If (InStr(1, Err.Description, "ORA-00942") > 0) Then
                MsgBox ("Запрашиваемая таблица не существует.")
            End If
            Exit Sub
            
        Case -2147217843:
            MsgBox ("Неверное имя пользователя или пароль.")
            Exit Sub
            
        Case 0:
        
        Case Else:
            MsgBox (Err & ", " & Err.Description)
            
    End Select

End Sub



Function PingIP(MyIP As String) As Boolean

' Returns True if IP is accessible, False if bot

Dim strCommand As String
Dim strPing As String

strCommand = "%ComSpec% /C %SystemRoot%\system32\ping.exe -n 1 -w 500 " & MyIP & " | " & "%SystemRoot%\system32\find.exe /i " & Chr(34) & "TTL=" & Chr(34)
strPing = fShellRun(strCommand)

If strPing = "" Then
    PingIP = False
Else
    PingIP = True
End If

End Function

Function fShellRun(sCommandStringToExecute)

' This function will accept a string as a DOS command to execute.
' It will then execute the command in a shell, and capture the output into a file.
' That file is then read in and its contents are returned as the value the function returns.

Dim oShellObject, oFileSystemObject, sShellRndTmpFile
Dim oShellOutputFileToRead, iErr

Set oShellObject = CreateObject("Wscript.Shell")
Set oFileSystemObject = CreateObject("Scripting.FileSystemObject")

    sShellRndTmpFile = oShellObject.ExpandEnvironmentStrings("%temp%") & oFileSystemObject.GetTempName
    On Error Resume Next
    oShellObject.Run sCommandStringToExecute & " > " & sShellRndTmpFile, 0, True
    iErr = Err.Number

    On Error GoTo 0
    If iErr <> 0 Then
        fShellRun = ""
        Exit Function
    End If

    On Error GoTo err_skip
    fShellRun = oFileSystemObject.OpenTextFile(sShellRndTmpFile, 1).ReadAll
    oFileSystemObject.DeleteFile sShellRndTmpFile, True

Exit Function

err_skip:
    fShellRun = ""
    oFileSystemObject.DeleteFile sShellRndTmpFile, True

End Function


Sub SaveBookToSeparateFile()

Dim targetWb  As Workbook
Dim currentWb As Workbook

Dim choiceFileDialog As Integer
Dim resultExtension  As Integer
Dim shtIndex         As Integer

Dim chosenExtension As String
Dim pathToFile      As String
Dim fullPathToSave  As Variant

Dim defaultSheetName As String


    Call ImprovePerformance(True)
        
    Set currentWb = ThisWorkbook
    
    ' Call FileDialog to choose location for file saving.
    fullPathToSave = Application.GetSaveAsFilename(InitialFileName:="", _
                                                    FileFilter:="Excel Workbook (*.xlsx),*.xlsx," + _
                                                                "Excel Binary Workbook (*.xlsb),*.xlsb,")
    'Interrupt sub if user pressed Cancel or X in FileDialog.
    If fullPathToSave = False Then
        MsgBox "Path is not chosen."
        Exit Sub
    End If

    chosenExtension = RxMatch(fullPathToSave, "\.[\w]+$")       'Get chosen excel workbook's extension.
    fullPathToSave = RxReplace(fullPathToSave, "\.[\w]+$", "")  'Cut extension from full path to saving workbook.
    
    
    Select Case chosenExtension
        Case ".xlsx"
            'You want to save Excel 2007-2016 file
            resultExtension = xlOpenXMLStrictWorkbook
        Case ".xlsb"
            'You want ta save Excel 2007-2016 BINARY file
            resultExtension = xlExcel12
    End Select
    
    ' Save new Workbook to specified folder with specified in FilDialog name.
    Workbooks.Add
    Set targetWb = ActiveWorkbook
    targetWb.SaveAs FileName:=fullPathToSave, FileFormat:=resultExtension
    
    ' Adding temporary list to have a possibility delete default
    ' workbook sheet.
    With targetWb
        .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = "EmptyList"
    End With
    
    ' Detect defualt sheet
    If SheetExists("Sheet1", targetWb) = True Then
        defaultSheetName = "Sheet1"
    ElseIf SheetExists("Лист1", targetWb) = True Then
        defaultSheetName = "Лист1"
    End If
    
    ' Delete default sheet
    targetWb.Sheets(defaultSheetName).Delete
    
    ' Copy sheets from cource workbook to target.
    shtIndex = 1
    Dim ws As Worksheet
    For Each ws In currentWb.Worksheets
        If ws.Visible = True Then
            currentWb.Activate
            ws.Copy Before:=targetWb.Sheets(targetWb.Sheets.Count)
        End If
        shtIndex = shtIndex + 1
    Next ws
    
    ' Handle cases when no sheets was cpoied.
    If targetWb.Sheets.Count <= 1 Then
        MsgBox ("Copy error. There are no sheets to copy.")
        Exit Sub
    End If
    
    If SheetExists("EmptyList", targetWb) = True Then: targetWb.Sheets("EmptyList").Delete
    
    targetWb.Save
    targetWb.Close
    
    Call ImprovePerformance(False)
    
    MsgBox ("Success.")

End Sub


Function SheetProtected(TargetSheet As Worksheet) As Boolean
     'Function purpose:  To evaluate if a worksheet is protected
     
    If TargetSheet.ProtectContents = True Then
        SheetProtected = True
    Else
        SheetProtected = False
    End If
     
End Function
 
Sub SetZoom(ZoomRate As Integer, WorksheetToZoom As Worksheet)

        WorksheetToZoom.Select
        ActiveWindow.Zoom = ZoomRate ' change as per your requirements

End Sub

Public Sub Populate2DimListBox(shtName As String, Optional wbObj As Workbook)

    Dim lr As Long, i As Long
    
    If wbObj Is Nothing Then Set wbObj = ThisWorkbook
    
    lr = GetBorders("LR", shtName, wbObj)
    
    With MultiSelectionForm.ListBox1 ' Change the form name here!
    .Clear
    .ColumnWidths = "70;200"
    '.RowSource = shtName & "!A2:B" & LR
    
    For i = 2 To lr
    .AddItem 'Populate listbox with items
    .List(i - 2, 0) = wbObj.Worksheets(shtName).Cells(i, 1).value
    .List(i - 2, 1) = wbObj.Worksheets(shtName).Cells(i, 2).value
    Next i
    
    End With

End Sub
Function GetADUsername()

    Dim objAD As Variant
    Dim objUser As Variant
    Dim strDisplayName As Variant
    
    Set objAD = CreateObject("ADSystemInfo")
    Set objUser = GetObject("LDAP://" & objAD.username)
    GetADUsername = objUser.samAccountName ''DisplayName
    
End Function
Function ColumnNumberToLetter(lngCol As Long) As String

    Dim vArr
    vArr = Split(Cells(1, lngCol).Address(True, False), "$")
    ColumnNumberToLetter = vArr(0)
    
End Function

Private Function IsInArray(valToBeFound As Variant, arr As Variant) As Boolean
'DEVELOPER: Ryan Wells (wellsr.com)
'DESCRIPTION: Function to check if a value is in an array of values
'INPUT: Pass the function a value to search for and an array of values of any data type.
'OUTPUT: True if is in array, false otherwise

Dim element As Variant
On Error GoTo IsInArrayError: 'array is empty
    For Each element In arr
        If element = valToBeFound Then
            IsInArray = True
            Exit Function
        End If
    Next element
Exit Function
IsInArrayError:
On Error GoTo 0
IsInArray = False
End Function

Function GetSheetBySelection(Optional wb As Workbook) As String
    
    If wb Is Nothing Then
        Set wb = ThisWorkbook
    End If
    
    If wb.Windows(1).SelectedSheets.Count Then
        GetSheetBySelection = wb.Windows(1).SelectedSheets(1).Name
    Else
        MsgBox ("Choose only one sheet!")
    End If
    

End Function

Function StringToArray(InputString As String, Optional Delimiter As String = ",") As Variant

    Dim tempArr As Variant
    Dim value   As Variant
    
    StringToArray = Split(InputString, Delimiter)
    
End Function

Function GetColumnsValues(ColumnNames As String, Optional HeaderRowNumber As Integer = 1, Optional DataStartRowNumber As Long, Optional NumberOfRows As Long, Optional ws As Worksheet, Optional wb As Workbook) As Variant

    If Not Debug_Mode Then On Error GoTo ErrHandler

    Dim LastColumnLetter    As String
    Dim LastColumnNumber    As Long
    Dim colsCounter         As Long
    Dim rowsCounter         As Long
    Dim arrayColsCounter    As Long
    Dim arrayRowsCounter    As Long
    Dim LastRowNumber       As Long
    Dim RowNumberDataStart  As Integer
    Dim RowNumberDataEnd    As Long
    Dim HeaderValue         As String
    Dim ColumnNamesArr      As Variant
    Dim value               As Variant
    Dim ColumnsToSave       As New Scripting.Dictionary
    Dim ResultValues()      As Variant

    If wb Is Nothing Then
        Set wb = ThisWorkbook
    End If
    
    If ws Is Nothing Then
        Set ws = wb.Sheets(GetSheetBySelection)
    End If
    
    LastRowNumber = GetBorders("LR", ws.Name, wb)
    
    If NumberOfRows = 0 Then
        NumberOfRows = LastRowNumber - HeaderRowNumber
    End If
    
    If DataStartRowNumber = 0 Then
        RowNumberDataStart = HeaderRowNumber + 1
    Else
        RowNumberDataStart = DataStartRowNumber
    End If
    
    RowNumberDataEnd = RowNumberDataStart + NumberOfRows - 1
    
    LastColumnNumber = GetBorders("LC", ws.Name, wb)
    LastColumnLetter = ColumnNumberToLetter(GetBorders("LC", ws.Name, wb))
    
    ColumnNamesArr = StringToArray(ColumnNames)
    
    
    With ws
    
        'Search for column and save its addresses by names.
        For colsCounter = 1 To LastColumnNumber
                HeaderValue = .Range(ColumnNumberToLetter(colsCounter) & HeaderRowNumber).value
                If IsInArray(HeaderValue, ColumnNamesArr) = True Then
                    ColumnsToSave.Add HeaderValue, ColumnNumberToLetter(colsCounter)
                End If
        Next colsCounter
        
        ' Saving data from list.
        arrayRowsCounter = 1
        For rowsCounter = RowNumberDataStart To RowNumberDataEnd
            arrayColsCounter = 1
            If .Range("A" & rowsCounter).EntireRow.Hidden = False Then
                ReDim Preserve ResultValues(1 To ColumnsToSave.Count, 1 To arrayRowsCounter)
                For Each value In ColumnsToSave
                    ResultValues(arrayColsCounter, arrayRowsCounter) = .Range(ColumnsToSave(value) & rowsCounter).value
                    arrayColsCounter = arrayColsCounter + 1
                Next value
            arrayRowsCounter = arrayRowsCounter + 1
            End If
        Next rowsCounter
                      
    End With

    GetColumnsValues = ResultValues
    Erase ResultValues
    
Done:
    Exit Function
ErrHandler:
    MsgBox ("Error in func GetColumnsValues: " & Err.Description)
    
End Function

Sub Count_Selection()
    Dim cell As Object
    Dim Count As Integer
    Count = 0
    For Each cell In Selection
        Count = Count + 1
    Next cell
    MsgBox Count & " item(s) selected"
End Sub

Function CheckFileExists(FileName As String) As Boolean

Dim strFileName As String
Dim strFileExists As String

    strFileName = FileName
    strFileExists = Dir(strFileName)

   If strFileExists = "" Then
        CheckFileExists = False
    Else
        CheckFileExists = True
    End If

End Function


Sub FixDateFormatInRange(rng As String, wsName As String, Optional wb As Workbook)

    If wb Is Nothing Then
        Set wb = ThisWorkbook
    End If
    
    With wb.Sheets(wsName).Range(rng)
    
        .Range(rng).value = .Range(rng).value
    
    End With

End Sub

Function CreateDictFromColumns(sheet As String, keyCol As String, valCol As String) As Dictionary

    Set CreateDictFromColumns = New Dictionary
    Dim rng As Range: Set rng = Sheets(sheet).Range(keyCol & ":" & valCol)
    Dim i As Long
    Dim lastCol As Long '// for non-adjacent ("A:ZZ")
    lastCol = rng.Columns.Count
    For i = 1 To rng.Rows.Count
        If (rng(i, 1).value = "") Then Exit Function
        CreateDictFromColumns.Add rng(i, 1).value, rng(i, lastCol).value
    Next
    
End Function

Function SelectionToDictionary(rng As Range, Optional wb As Workbook) As Dictionary
    ' NOTE: Works only for one column.
    
    If wb Is Nothing Then Set wb = ThisWorkbook
    
    Dim tmpDict As Dictionary
    Set tmpDict = New Dictionary
    
    Dim cll As Range
    
    For Each cll In rng
        If cll.EntireRow.Hidden = False And Not tmpDict.Exists(cll.value) Then
            tmpDict.Add cll.value, Nothing
        End If
    Next cll
    
    Set SelectionToDictionary = tmpDict

End Function


Function ConcatToString(input_object As Variant, Delimiter As String)

    Dim tmpElement  As Variant
    Dim tmpString   As String
    
    For Each tmpElement In input_object
        tmpString = tmpString & CStr(tmpElement) & Delimiter
    Next tmpElement
    
    ConcatToString = Left(tmpString, Len(tmpString) - 1)

End Function

Function CreateConnection(DbServerAddressLocal As String, Optional DbNameLocal As String, Optional OpenConnection As Boolean = True, Optional ConnectionTimeOut As Integer = 0) As ADODB.Connection
                                                   
    If Not Debug_Mode Then On Error GoTo ErrHandler
                                                   
    Dim conn As New ADODB.Connection
    
    conn.CursorLocation = adUseClient
    conn.ConnectionString = "Provider=SQLOLEDB;Data Source=" & DbServerAddressLocal & ";Trusted_connection=yes;Database=" & DbNameLocal
    
    conn.CursorLocation = adUseClient
    
    Set CreateConnection = conn
    
    CreateConnection.CommandTimeout = ConnectionTimeOut
    
    If conn.State = adStateOpen Then: CreateConnection.Close
    If OpenConnection = True Then: CreateConnection.Open

Done:
    Exit Function
ErrHandler:
    MsgBox ("Error in func CreateConnection: " & Err.Description)
    
End Function

Function InsertArrayToServer(arr As Variant, TargetTableName As String, TargetColumns As String, conn As ADODB.Connection)

    If Not Debug_Mode Then On Error GoTo ErrHandler

    Dim ColsCounterArr  As Long
    Dim RowsCounterArr  As Long
    Dim ColumnsString   As String
    Dim ValuesRow       As String
    Dim Row             As String
    Dim InsertQuery     As String
    Dim MaxLenValueRow  As Integer
    Dim CmdQry          As New ADODB.Command
    
    Dim rs              As New ADODB.Recordset

    CmdQry.CommandType = adCmdText
    Set CmdQry.ActiveConnection = conn
    
    ColumnsString = ColumnNamesToSquareBrackets(TargetColumns)

    InsertQuery = "INSERT INTO " & TargetTableName & " (" & ColumnsString & ")" & "VALUES "
    MaxLenValueRow = 2048
    ValuesRow = ""
    Row = ""

    Dim temp As Variant
    RowsCounterArr = 1
    Do While RowsCounterArr <= UBound(arr, 2)
    
    
        If RowsCounterArr = UBound(arr, 2) Then
            
            ValuesRow = ""
        
            For ColsCounterArr = 1 To UBound(arr, 1)
                ValuesRow = ValuesRow & StringToMSSQLFormat(arr(ColsCounterArr, RowsCounterArr), True) & ","
            Next ColsCounterArr
            
            Row = Row & "(" & Left(ValuesRow, Len(ValuesRow) - 1) & "),"
            temp = InsertQuery & Left(Row, Len(Row) - 1)
            CmdQry.CommandText = InsertQuery & Left(Row, Len(Row) - 1)
            CmdQry.Execute

        ElseIf Len(ValuesRow) <= MaxLenValueRow Then
            
            ValuesRow = ""
        
            For ColsCounterArr = 1 To UBound(arr, 1)
                ValuesRow = ValuesRow & StringToMSSQLFormat(arr(ColsCounterArr, RowsCounterArr), True) & ","
            Next ColsCounterArr
            
            Row = Row & "(" & Left(ValuesRow, Len(ValuesRow) - 1) & "),"
            
        Else
        
            CmdQry.CommandText = InsertQuery & Left(Row, Len(Row) - 1)
            CmdQry.Execute
            
            ValuesRow = ""
            
            For ColsCounterArr = 1 To UBound(arr, 1)
                ValuesRow = ValuesRow & StringToMSSQLFormat(arr(ColsCounterArr, RowsCounterArr), True) & ","
            Next ColsCounterArr
            
            Row = "(" & Left(ValuesRow, Len(ValuesRow) - 1) & "),"
            
            
        End If
        
    RowsCounterArr = RowsCounterArr + 1
    Loop

Done:
    Exit Function
ErrHandler:
    MsgBox ("Error in func InsertArrayToServer: " & Err.Description)

End Function

Function CreateTemporaryTable(TableName As String, ColumnNames As String, conn As ADODB.Connection, Optional ColumnsDelimiter As String = ",")

    If Not Debug_Mode Then On Error GoTo ErrHandler

    Dim DefaultDataType     As String
    Dim ColumnsArr          As Variant
    Dim ColumnsInQuery      As String
    Dim value               As Variant
    Dim cmdQuery            As New ADODB.Command
    
    If conn.State = adStateClosed Then
        MsgBox ("Error in func CreateTemporaryTable: Connection closed.")
        Exit Function
    End If
    
    DefaultDataType = "nvarchar(255) NULL"
    ColumnsArr = Split(ColumnNames, ColumnsDelimiter)
    
    For Each value In ColumnsArr
        ColumnsInQuery = ColumnsInQuery & "[" & value & "] " & DefaultDataType & ","
    Next value
    ColumnsInQuery = Left(ColumnsInQuery, Len(ColumnsInQuery) - 1)
    
    cmdQuery.CommandType = adCmdText
    Set cmdQuery.ActiveConnection = conn
    cmdQuery.CommandText = "CREATE TABLE " & TableName & " (" & ColumnsInQuery & ")"
    cmdQuery.Execute
    

Done:
    Exit Function
ErrHandler:
    MsgBox ("Error in func CreateTemporaryTable: " & Err.Description)
    
End Function

Function ColumnNamesToSquareBrackets(ColumnNames As String, Optional Delimiter As String = ",") As String

    Dim ColumnsArr  As Variant
    Dim Column      As Variant
    Dim Result      As String
    
    ColumnsArr = Split(ColumnNames, Delimiter)
    
    For Each Column In ColumnsArr
        Result = Result & "[" & Column & "],"
    Next Column

    ColumnNamesToSquareBrackets = Left(Result, Len(Result) - 1)

End Function

Function GetSelectionRows(FirstOrLast As String)
    
    If FirstOrLast = "F" Then
        GetSelectionRows = Selection.Rows(1).Row
    ElseIf FirstOrLast = "L" Then
        GetSelectionRows = Selection.Rows.Count + Selection.Rows(1).Row - 1
    End If

End Function

Sub DeclareUserDefinedTypeTable(typeName As String, conn As ADODB.Connection)

    Dim sql As String: sql = "DECLARE @" & typeName & " " & typeName

End Sub

Function BooleanToBit(booleanVal As Boolean) As Integer


    If Not IsNull(booleanVal) Then
    
        If booleanVal = True Then
            BooleanToBit = 1
        ElseIf booleanVal = False Then
            BooleanToBit = 0
        End If
    
    End If


End Function

Function BitToBoolean(val As Integer) As Boolean


    If Not IsNull(val) Then
    
        If val = 1 Then
            BitToBoolean = True
        ElseIf val = 0 Then
            BitToBoolean = False
        End If
    
    End If


End Function

Public Sub InitListBox(LstBox As Variant, shtName As String, Optional wbObj As Workbook)

    Dim lr As Long, i As Long
    
    If wbObj Is Nothing Then Set wbObj = ThisWorkbook
    
    lr = GetBorders("LR", shtName, wbObj)
    
    With LstBox
        
        .RowSource = "" ' Clear Listbox before init.
        
        
        .List = wbObj.Sheets(shtName).Range("A1:F" & lr).value
        .ColumnCount = UBound(.List, 2) + 1
        .RowSource = "=" & shtName & "!$A$2:$F$" & lr
        .ColumnWidths = "130;40;40;30;60;100"
        
    End With

End Sub

Function GetListboxSelection(LstBox As MSForms.listBox) As Variant
    ' Write selected row from ListBox to Array.
    ' Returns null if no selection.
    
    Dim Row As Long, Column As Long
    Dim Result As Variant
   
    For Row = 0 To LstBox.ListCount - 1
        If LstBox.Selected(Row) = True Then
            
            ReDim Result(0, LstBox.ColumnCount)
            For Column = 0 To LstBox.ColumnCount - 1
               
               Result(0, Column) = LstBox.List(Row, Column)
            Next Column
            
        End If
    Next Row
    
   GetListboxSelection = Result
   
End Function

Sub InitComboBoxFromSqlQuery(CmbBox As ComboBox, sqlQuery As String, conn As ADODB.Connection)

    If Not Debug_Mode Then On Error GoTo ErrHandler

    Dim rs As ADODB.Recordset

    Set rs = conn.Execute(sqlQuery)
    
    CmbBox.Clear
    
    If Not rs.EOF Then
        Do While Not rs.EOF
            
            With CmbBox
                If Not IsNull(rs.Fields(0).value) Then
                    .AddItem rs.Fields(0).value
                End If
            End With
            
            rs.MoveNext
        Loop
    End If
    rs.Close
    
Done:
    Exit Sub
ErrHandler:
    MsgBox ("Error in sub InitComboBoxFromSqlQuery: " & Err.Description)
    
End Sub

Function replaceNullToEmptyString(val As Variant)

    If IsNull(val) Then
        replaceNullToEmptyString = ""
    Else
        replaceNullToEmptyString = val
    End If

End Function


Function getADLogin()

    Dim objAD As Variant
    Dim objUser As Variant
    Dim strDisplayName As Variant
    
    Set objAD = CreateObject("ADSystemInfo")
    Set objUser = GetObject("LDAP://" & objAD.username)
    getADLogin = Replace(objUser.Name, "CN=", "")
    
End Function

Function columnLetterToNumber(letter As String)
'PURPOSE: Convert a given letter into it's corresponding Numeric Reference
'SOURCE: www.TheSpreadsheetGuru.com/the-code-vault

    Dim ColumnNumber As Long
  
    'Convert To Column Number
    columnLetterToNumber = Range(letter & 1).Column

    
End Function


Sub freezeRowsAndCols(rowNumber As Long, colNumber As Long, wsName As String, Optional wb As Workbook)

    If wb Is Nothing Then Set wb = ThisWorkbook
    
    wb.Worksheets(wsName).Activate

    With ActiveWindow
        .SplitColumn = colNumber
        .SplitRow = rowNumber
        .FreezePanes = True
    End With


End Sub
