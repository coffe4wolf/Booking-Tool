Attribute VB_Name = "m_main"
Option Explicit

Const CELL_TO_START_DATES As String = "B6"
Const CELL_TO_START_TIMES As String = "A7"
Const COUNT_OF_DAYS_TO_DRAW As Integer = 30
Const COUNT_OF_TIMES_TO_DRAW As Integer = 28

Sub getBookings(meetingRoomID As Integer)

    Dim wsBooking   As Worksheet
    Dim counter     As Long
    Dim dateValue   As String
    Dim query       As String
    Dim rng         As Range
    
    'On Error GoTo ErrHandler
    
    Set wsBooking = ThisWorkbook.Sheets(BOOKING_WS_NAME)
    
    With wsBooking
    
        For counter = 2 To COUNT_OF_DAYS_TO_DRAW + 1
        
            dateValue = Format(.Cells(6, counter).value, "yyyyMMdd")
            
            Set rng = .Cells(6, counter)
        
            query = "EXEC [BookingConferenceRooms].[GetBookingsByDate] " & meetingRoomID & " , " & StringToMSSQLFormat(dateValue)
            Call RunSQLSelect(BOOKING_WS_NAME, query, DbServerAddress, DbName, ThisWorkbook, ColumnNumberToLetter(counter) & "7", False)
        
        Next counter
    
    End With
    
Done:
    Exit Sub
ErrHandler:
    MsgBox ("Error! Source: " & Err.Source & " Description: " & Err.Description & "(" & Err.Number & ")")
End Sub


Sub drawSchedule()

    Dim cellToStartDates    As Range
    Dim cellToStartTimes    As Range
    
    Dim dates               As Range
    Dim bookingRange        As Range
    
    Dim wsBooking As Worksheet
    
    Dim lr       As Long
    Dim lc       As Long
    Dim lcLeteer As String
    
    Dim counter  As Long
    
    'On Error GoTo ErrHandler
    
    Set wsBooking = ThisWorkbook.Sheets(BOOKING_WS_NAME)
    
    
    With wsBooking
    
        lr = GetBorders("LR", .Name, ThisWorkbook)
        lc = GetBorders("LC", .Name, ThisWorkbook)
        lcLeteer = ConvertToLetter(lc)
    
        ' Clear sheet.
        .Range("A6:AZ10000").Clear
        
        Set cellToStartDates = .Range(CELL_TO_START_DATES)
        Set cellToStartTimes = .Range(CELL_TO_START_TIMES)
        
        ' Draw dates.
        cellToStartDates.value = Date
        cellToStartDates.AutoFill .Range(cellToStartDates, cellToStartDates.Offset(0, COUNT_OF_DAYS_TO_DRAW))
        
        ' Draw times.
        cellToStartTimes = "7:00"
        cellToStartTimes.Offset(1, 0) = "7:30"
        Set cellToStartTimes = Nothing
        Set cellToStartTimes = .Range(CELL_TO_START_TIMES & ":" & Left(CELL_TO_START_TIMES, 1) & CInt(Right(CELL_TO_START_TIMES, 1)) + 1)
        
        cellToStartTimes.AutoFill .Range(cellToStartTimes, cellToStartTimes.Offset(COUNT_OF_TIMES_TO_DRAW, 0))
        
        'Format time.
        .Columns(1).ColumnWidth = 7
        .Columns(1).Font.Bold = True
        .Columns(1).Font.Size = 12
        
        Set dates = .Range("B" & Right(CELL_TO_START_DATES, 1) & ":" & lcLeteer & Right(CELL_TO_START_DATES, 1))
        
        ' Format columns.
        With dates
        
            .ColumnWidth = 25
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .Font.Size = 12
        
        End With
        
        Call freezeRowsAndCols(6, 1, .Name)
    
    End With
    
    ActiveWindow.Zoom = 85
    
Done:
    Exit Sub
ErrHandler:
    MsgBox ("Error! Source: " & Err.Source & " Description: " & Err.Description & "(" & Err.Number & ")")
End Sub

Sub unbookARoom()

    Dim bookingWs       As Worksheet
    Dim query           As String
    
    Dim conn            As New ADODB.Connection

    Dim selectedRange       As Range
    Dim startOfSelection    As Range
    Dim endOfSelection      As Range
    
    Dim selectedDate        As String
    Dim selectedTimeStart   As String
    Dim selectedTimeEnd     As String
    
    Dim selectedDtStart     As String
    Dim selectedDtEnd       As String
    
    Dim roomId              As String
    
    Dim messageTextResult   As String
    
    Dim rs                  As ADODB.Recordset
    
    query = "EXEC [BookingConferenceRooms].[unbookRoom] {1}"
    
    Set bookingWs = ThisWorkbook.Sheets(BOOKING_WS_NAME)
    
    Set selectedRange = Application.Selection
    
    With bookingWs
    
        ' Get start and end of selection.
        If selectedRange.Cells.Count > 1 Then
            Set startOfSelection = .Range(Left(selectedRange.Address, InStr(1, selectedRange.Address, ":") - 1))
            Set endOfSelection = .Range(Right(selectedRange.Address, InStr(1, selectedRange.Address, ":") - 1))
        Else
            Set startOfSelection = .Range(selectedRange.Address)
            Set endOfSelection = .Range(selectedRange.Address)
        End If
    
        selectedDate = Format(.Range(ColumnNumberToLetter(startOfSelection.Column) & Right(CELL_TO_START_DATES, 1)), "yyyy.MM.dd")
        selectedTimeStart = Format(.Range("A" & startOfSelection.Row).value, "hh:mm")
        selectedTimeEnd = Format(.Range("A" & endOfSelection.Row).value, "hh:mm")
    
        selectedDtStart = selectedDate & " " & selectedTimeStart
        selectedDtEnd = selectedDate & " " & selectedTimeEnd
    
    End With
    
    roomId = getMeetingRoomIdByName(Worksheets(BOOKING_WS_NAME).meetingRooms_ComboBox.value)

    ' Fill query with data.
    query = Replace(query, "{1}", roomId & "," & StringToMSSQLFormat(selectedDtStart) & "," & StringToMSSQLFormat(selectedDtEnd))
    
    Set conn = CreateConnection(DbServerAddress, DbName)
    Set rs = conn.Execute(query)
    
    Call updateScedule

End Sub

Sub BookARoom()

    Dim bookingWs       As Worksheet
    Dim query           As String
    
    Dim conn            As New ADODB.Connection

    Dim selectedRange       As Range
    Dim startOfSelection    As Range
    Dim endOfSelection      As Range
    
    Dim selectedDate        As String
    Dim selectedTimeStart   As String
    Dim selectedTimeEnd     As String
    
    Dim selectedDtStart     As String
    Dim selectedDtEnd       As String
    
    Dim note                As String
    Dim roomId              As String
    
    Dim messageTextResult   As String
    
    Dim rs                  As ADODB.Recordset
    
    query = "SET NOCOUNT ON; " & _
            "DECLARE @RoomIsOccupied tinyint;SET @RoomIsOccupied = (SELECT [BookingConferenceRooms].[CheckRoomIsOccupied] ({1})); " & _
            "IF @RoomIsOccupied = 1 BEGIN SELECT 'Error: Cannot book a room: its occupied.'; END " & _
            "IF @RoomIsOccupied = 0 BEGIN EXEC [BookingConferenceRooms].[BookRoom] {2}; SELECT 'Success!'; END "
    
    Set bookingWs = ThisWorkbook.Sheets(BOOKING_WS_NAME)
    
    Set selectedRange = Application.Selection
    
    With bookingWs
    
        ' Get start and end of selection.
        If selectedRange.Cells.Count > 1 Then
            Set startOfSelection = .Range(Left(selectedRange.Address, InStr(1, selectedRange.Address, ":") - 1))
            Set endOfSelection = .Range(Right(selectedRange.Address, InStr(1, selectedRange.Address, ":") - 1))
        Else
            Set startOfSelection = .Range(selectedRange.Address)
            Set endOfSelection = .Range(selectedRange.Address)
        End If
    
        selectedDate = Format(.Range(ColumnNumberToLetter(startOfSelection.Column) & Right(CELL_TO_START_DATES, 1)), "yyyy.MM.dd")
        selectedTimeStart = Format(.Range("A" & startOfSelection.Row).value, "hh:mm")
        selectedTimeEnd = Format(.Range("A" & endOfSelection.Row).value, "hh:mm")
    
        selectedDtStart = selectedDate & " " & selectedTimeStart
        selectedDtEnd = selectedDate & " " & selectedTimeEnd
    
    End With
    
    note = Trim(InputBox("Enter note:"))
    
    If note = "" Then
        Exit Sub
    End If
    
    roomId = getMeetingRoomIdByName(Worksheets(BOOKING_WS_NAME).meetingRooms_ComboBox.value)

    ' Fill query with data.
    query = Replace(query, "{1}", roomId & "," & StringToMSSQLFormat(selectedDtStart) & "," & StringToMSSQLFormat(selectedDtEnd))
    query = Replace(query, "{2}", roomId & "," & StringToMSSQLFormat(selectedDtStart) & "," & StringToMSSQLFormat(selectedDtEnd) & ",N" & StringToMSSQLFormat(note))
    
     
    Set conn = CreateConnection(DbServerAddress, DbName)
    Set rs = conn.Execute(query)
    
    Call updateScedule

End Sub

Sub loadRooms()

    Dim dataWs          As Worksheet
    Dim rangeToPast     As Range
    Dim query           As String
    Dim rangeToClear    As String
    
    Dim datesCounter    As Long
    
    'On Error GoTo ErrHandler
    
    query = "SELECT [ID],[Name] FROM [BookingConferenceRooms].[Rooms]"
    
    Set dataWs = ThisWorkbook.Sheets(DATA_WS_NAME)
    
    rangeToClear = RANGE_TO_PASTE_ROOMS_ON_DATASHEET & ":" & Left(RANGE_TO_PASTE_ROOMS_ON_DATASHEET, 1) & 100
    
    dataWs.Range(rangeToClear).Clear
    
    
Done:
    Exit Sub
ErrHandler:
    MsgBox ("Error! Source: " & Err.Source & " Description: " & Err.Description & "(" & Err.Number & ")")
End Sub

Sub getFreeRoomsByTime()

    Dim bookingWs       As Worksheet
    Dim query           As String
    
    Dim conn            As New ADODB.Connection

    Dim selectedRange       As Range
    Dim startOfSelection    As Range
    Dim endOfSelection      As Range
    
    Dim selectedDate        As String
    Dim selectedTimeStart   As String
    Dim selectedTimeEnd     As String
    
    Dim selectedDtStart     As String
    Dim selectedDtEnd       As String
    
    Dim note                As String
    Dim roomId              As String
    
    Dim messageTextResult   As String
    
    Dim rs                  As ADODB.Recordset
    
    query = "[BookingConferenceRooms].[getFreeRoomsByTime] {1}"
    
    Set bookingWs = ThisWorkbook.Sheets(BOOKING_WS_NAME)
    
    Set selectedRange = Application.Selection
    
    With bookingWs
    
        ' Get start and end of selection.
        If selectedRange.Cells.Count > 1 Then
            Set startOfSelection = .Range(Left(selectedRange.Address, InStr(1, selectedRange.Address, ":") - 1))
            Set endOfSelection = .Range(Right(selectedRange.Address, InStr(1, selectedRange.Address, ":") - 1))
        Else
            Set startOfSelection = .Range(selectedRange.Address)
            Set endOfSelection = .Range(selectedRange.Address)
        End If
    
        selectedDate = Format(.Range(ColumnNumberToLetter(startOfSelection.Column) & Right(CELL_TO_START_DATES, 1)), "yyyy.MM.dd")
        selectedTimeStart = Format(.Range("A" & startOfSelection.Row).value, "hh:mm")
        selectedTimeEnd = Format(.Range("A" & endOfSelection.Row).value, "hh:mm")
    
        selectedDtStart = selectedDate & " " & selectedTimeStart
        selectedDtEnd = selectedDate & " " & selectedTimeEnd
    
    End With
    
    roomId = getMeetingRoomIdByName(Worksheets(BOOKING_WS_NAME).meetingRooms_ComboBox.value)

    ' Fill query with data.
    query = Replace(query, "{1}", StringToMSSQLFormat(selectedDtStart) & "," & StringToMSSQLFormat(selectedDtEnd))
     
    Set conn = CreateConnection(DbServerAddress, DbName)
    Set rs = conn.Execute(query)
    
    
    ' Init userform.
    If Not rs.EOF Then
        Do While Not rs.EOF
            roomsList_UserForm.roomsList_ListBox.AddItem rs.Fields(0).value
            rs.MoveNext
        Loop
    End If
    
    roomsList_UserForm.Show
    
    Call updateScedule

End Sub

Sub initMeetingRoomsComboBox()

    Dim dataWs      As Worksheet
    Dim bookingWs   As Worksheet
    Dim roomsRange  As Range
    Dim rowsCounter As Long
    Dim cll         As Range
    
    'On Error GoTo ErrHandler
    
    Set dataWs = ThisWorkbook.Sheets(DATA_WS_NAME)
    Set bookingWs = ThisWorkbook.Sheets(BOOKING_WS_NAME)
    
    
    dataWs.Activate
    dataWs.Range(RANGE_TO_PASTE_ROOMS_ON_DATASHEET).Select
    ActiveCell.CurrentRegion.Select
    
    Set roomsRange = Application.Selection

    For Each cll In roomsRange
    
        If cll.Row > 1 And cll.Column = 2 Then
            Worksheets("booking").meetingRooms_ComboBox.AddItem cll.value
        End If
        
    Next cll
    
    bookingWs.Activate
    
Done:
    Exit Sub
ErrHandler:
    MsgBox ("Error! Source: " & Err.Source & " Description: " & Err.Description & "(" & Err.Number & ")")
End Sub

Function getMeetingRoomIdByName(roomName As String)

    'On Error GoTo ErrHandler
    
    Dim bookingWs   As Worksheet
    Dim dataWs      As Worksheet
    Dim roomsRange  As Range
    Dim cll         As Range
    
    Set bookingWs = ThisWorkbook.Sheets(BOOKING_WS_NAME)
    Set dataWs = ThisWorkbook.Sheets(DATA_WS_NAME)
    
    dataWs.Activate
    dataWs.Range(RANGE_TO_PASTE_ROOMS_ON_DATASHEET).Select
    ActiveCell.CurrentRegion.Select
    
    Set roomsRange = Application.Selection
    
    For Each cll In roomsRange
    
        If cll.value = roomName Then
            getMeetingRoomIdByName = cll.Offset(0, -1).value
        End If
        
    Next cll
    
    bookingWs.Activate
    
Done:
    Exit Function
ErrHandler:
    MsgBox ("Error! Source: " & Err.HelpContext & " Description: " & Err.Description & "(" & Err.Number & ")")
End Function

Sub updateScedule()

    Call ImprovePerformance(True)

    Call drawSchedule
    Call getBookings(getMeetingRoomIdByName(Worksheets(BOOKING_WS_NAME).meetingRooms_ComboBox.value))
    Call MergeSimilarCells(Worksheets(BOOKING_WS_NAME).Range("B7:AZ36"))
    
    Call ImprovePerformance(False)
    
Done:
    Exit Sub
ErrHandler:
    MsgBox ("Error! Source: " & Err.HelpContext & " Description: " & Err.Description & "(" & Err.Number & ")")
End Sub

Sub MergeSimilarCells(rng As Range)

    Dim myRange As Range
    Dim cell As Variant
    Set myRange = rng

CheckAgain:
    For Each cell In myRange
    
        If cell.value <> "" Then: cell.Interior.ColorIndex = BOOKING_COLOR_INDEX
        
        If cell.value = cell.Offset(1, 0).value And Not IsEmpty(cell) Then
            Range(cell, cell.Offset(1, 0)).Merge
            cell.VerticalAlignment = xlCenter
            GoTo CheckAgain
        End If
    Next

End Sub



