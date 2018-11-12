Const ATCOPath As String = "Z:\ATCO & ADCO Roster\"
Const ATFSOPath As String = "Z:\ATFSO_SATCO_ADCT Roster\"

#Const developMode = False
Const csRow As Integer = 32
Const nameRow As Integer = 33
Const noteRow As Integer = 34
Const streamRow As Integer = 35
Dim AverageTime As Single
Dim ElapsedTime As Single
Dim OpenTime As Single
Dim logFile As Object

Function ShiftDeconf(ATCO As String, ATFSO As String, ATFSONewer As Boolean, Optional cs As String = "", Optional day As Integer = 0) As String
  
    Dim AShift As String
    Dim AStream As String
    Dim AOJT As String
    Dim FShift As String
    Dim FStream As String
    Dim FOJT As String
    Dim SolvedShift As String
    
    Dim cell As String
    Dim tmpStr As Variant
    Dim shiftMatch As Boolean: shiftMatch = False
    Dim streamMatch As Boolean: shiftMatch = False
    Dim sick As Boolean: sick = False
    
    ' Return if ATCO is empty or both are the same
    If ATCO = "" Or ATFSO = ATCO Then
        ShiftDeconf = ATFSO
        Exit Function
    End If
    
    ' Split ATCO
    tmpStr = Split(ATCO, ";", 3)
    If Right(ATCO, 3) = ";S;" Then
        sick = True
    End If
    If UBound(tmpStr) = 2 Then
        AStream = tmpStr(0)
        AShift = tmpStr(1)
        AOJT = tmpStr(2)
    End If
    
    ' Split ATFSO
    tmpStr = Split(ATFSO, ";", 3)
    If UBound(tmpStr) = 2 Then
        FStream = tmpStr(0)
        FShift = Replace(tmpStr(1), " - ", "-")
        FOJT = tmpStr(2)
    End If

    ' Check stream
    If (AStream Like "TWR*" Or AStream Like "CDC*" And FStream Like "CDC*") Or AStream = FStream Or AStream Like "See Note*" Or AStream Like "*Course" Then
        streamMatch = True
    End If
    
    SolvedShift = FShift
    If streamMatch Then
        ' Check shift
        If AShift = FShift Or AShift = "" Then
            shiftMatch = True
        ElseIf Shift2Time(AShift) = FShift Then
            SolvedShift = AShift
            shiftMatch = True
        End If
    End If
    
    If Not streamMatch Or Not shiftMatch And Not ATFSONewer Then
        ' Stream and shift does not match and ATFSO is older, then use ATCO entry
        ShiftDeconf = ATCO
        
        ' Append Sick if Comp.leave is found
        If sick = False And FStream Like "*Comp.leave" Then
            ShiftDeconf = ShiftDeconf & "S;"
        End If
    Else
        ' If using ATFSO shift, replace OJT in stream
        If InStr(FShift, "OJT") Then
            FStream = Trim(Replace(FStream, "OJT", ""))
            If FOJT Like "N;*" Then
                FOJT = "Y;"
            End If
        End If
        
        ShiftDeconf = FStream & ";" & SolvedShift & ";" & FOJT
        If sick Then
            ShiftDeconf = ShiftDeconf & ";S;"
        End If
    End If
    
    #If developMode Then
        If ATFSO <> ATCO Then Debug.Print cs & "|" & day & "|" & ATCO & "|" & ATFSO & "|" & streamMatch & "|" & shiftMatch & "|" & ATFSONewer & "|" & ShiftDeconf
    #End If
    
End Function
Function Shift2Time(s As String) As String
    Static SHIFTS As New Scripting.Dictionary
    
    ' Init shifts
    If SHIFTS.Count <> 21 Then
        With SHIFTS
            'On Error Resume Next
            .CompareMode = vbTextCompare
            .Add "E*", "0730-1500"
            .Add "E1", "0745-1500"
            .Add "F#", "0800-1530"
            .Add "F1", "0845-1530"
            .Add "F2", "0845-1600"
            .Add "F3", "0845-1700"
            .Add "F4", "0900-1700"
            .Add "G1", "0945-1800"
            .Add "G2", "0945-1900"
            .Add "H1", "1045-1900"
            .Add "H2", "1045-2000"
            .Add "J1", "1245-2130"
            .Add "A*", "1415-2200"
            .Add "A1", "1430-2200"
            .Add "A2", "1430-2300"
            .Add "C1", "1530-2300"
            .Add "C2", "1545-2345"
            .Add "C3", "1715-0115"
            .Add "N*", "2130-0800"
            .Add "N1", "2145-0800"
            .Add "N2", "2100-0715"
            'On Error GoTo 0
        End With
    End If
    
    If SHIFTS.Exists(s) Then
        Shift2Time = SHIFTS(s)
    Else
        Shift2Time = ""
    End If
End Function

Function NB_DAYS(date_test As Date)
    NB_DAYS = day(DateSerial(Year(date_test), Month(date_test) + 1, 1) - 1)
End Function

Private Sub OptimizeCode_Begin()
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = False
    Application.EnableEvents = False
    ActiveSheet.DisplayPageBreaks = False
End Sub

Private Sub OptimizeCode_End()
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.DisplayStatusBar = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
End Sub

Public Function IsAlpha(strValue As String) As Boolean
    IsAlpha = strValue Like WorksheetFunction.Rept("[a-zA-Z]", Len(strValue))
End Function

Sub ATCO(RosterDate As String)
    Dim filename As String
    Dim NumberOfDays As Integer
    Dim wb1, wb2 As Workbook
    Dim ws1, ws2 As Worksheet
    Dim result As Range
    Dim firstDayCol As Range
    Dim noteStart As Integer
    Dim noteEnd As Integer
    Dim rng As Range
    Dim day As Integer
    Dim cs As String
    Dim name As String
    Dim Shift As String
    Dim stream As String
    Dim entryStr As String
    Dim asuCell As Range
    Dim OJTI As String
    Dim cellStr As String
    Dim tmpStr As Variant
    Dim notes As String
    Dim t As Single
  
    filename = ATCOPath & RosterDate & "*.xlsx"
    
    If Dir(filename) = "" Then
        Debug.Print filename & " not found"
        Exit Sub
    End If
    
    filename = ATCOPath & Dir(filename)
    
    Set wb1 = ThisWorkbook
    
    ' Close File if already opened
    On Error Resume Next
    Workbooks(filename).Close SaveChanges:=False
    On Error GoTo 0
    
    t = Timer()
    
    logFile.Write vbTab & "Open"
    Set wb2 = Workbooks.Open(filename:=filename, Password:="aerostar", UpdateLinks:=0)
    OpenTime = OpenTime + Timer() - t
    
    wb2.Windows(1).Visible = False
    
    Set ws1 = wb1.Sheets(RosterDate)
    
    Set ws2 = wb2.Sheets("MASTER")
    ws2.Unprotect ("AABABABABBAN")
    ws2.Columns.EntireColumn.Hidden = False
    ws2.Rows.EntireRow.Hidden = False
    
    'Set result = ws2.Range("B:B").Find("app", LookIn:=xlValues)
    Set result = ws2.Range("A:A").Find("E1", LookIn:=xlValues).Offset(0, 1)
    Set firstDayCol = ws2.UsedRange.Find(Format("1 " & RosterDate, "d-mmm"))
    
    If Not firstDayCol Is Nothing Then
        'Debug.Print "firstRow: " & firstDayCol.Column
    End If
    
    ' Find first cell with value above asu cell to set stream
    'Set asuCell = ws2.Range("B:B").Find("asu", LookIn:=xlValues)
    Set asuCell = ws2.Range("B:B").Find(ws2.Range("B:B").Find("asu", LookIn:=xlValues).Offset(-1, 0).Value, LookIn:=xlValues)
    
    NumberOfDays = NB_DAYS(DateValue("1 " & RosterDate))
    
    logFile.Write vbTab & "Shifts"
    
    For day = 1 To NumberOfDays
        ' Normal roster
        i = result.Row
        stream = "TMC"
        'stream = result.Text
        Do While ws2.Cells(i, result.Column).Value <> "C/S 1" ' Loop til other mannings
            cs = UCase(Trim(ws2.Cells(i, firstDayCol.Column + day - 1).Value))
            If cs <> "" Then
                Shift = ws2.Cells(i, 1).Value
                If ws2.Cells(i, result.Column).Value = "app" Then
                    stream = "APP"
                ElseIf ws2.Cells(i, result.Column).Value = "amn" Then
                    stream = "TWR"
                ElseIf ws2.Cells(i, result.Column).Value = "tre" Then
                    stream = "AREA"
                ElseIf i >= asuCell.Row Then
                    stream = UCase(ws2.Cells(i, result.Column).Value)
                    If stream = "EXR" Then
                        stream = "XRM"
                    End If
                End If
                
                If IsAlpha(cs) And (Len(cs) = 4 Or Len(cs) = 2) Then
                    ' Only Write to cell if cs is valid
                    cellStr = Left(cs, 2) & day
                    
                    ws1.Range(cellStr).Value = (stream & ";" & Shift & ";N;")
                    ws1.Range(cellStr).EntireColumn.Hidden = False
                
                    ' Trainee
                    If Len(cs) = 4 Then
                        ' OJTI cs
                        OJTI = Left(cs, 2)
                        ' OJT cs
                        cs = Right(cs, 2)
                        
                        ' Append OJT to OJTI roster
                        ws1.Range(cellStr).Value = ws1.Range(cellStr).Value & cs & ";"
                        
                        ' OJT cell
                        cellStr = cs & day
                        ws1.Range(cellStr).Value = (stream & ";" & Shift & ";Y;" & OJTI & ";")
                        ws1.Range(cellStr).EntireColumn.Hidden = False
                    Else
                        ' Append ; if no trainee
                        ws1.Range(cellStr).Value = ws1.Range(cellStr).Value & ";"
                    End If
                End If
            End If
        i = i + 1
        Loop

        ' Other manning
        Do While Len(ws2.Cells(i, result.Column).Value) <> 2 ' Loop til callsign rows
            cs = UCase(Trim(ws2.Cells(i, firstDayCol.Column + day - 1).Value))
            If IsAlpha(cs) And (Len(cs) = 4 Or Len(cs) = 2) Then
                stream = Trim(ws2.Cells(i + 1, firstDayCol.Column + day - 1).Value)
                Shift = Trim(ws2.Cells(i + 2, firstDayCol.Column + day - 1).Value)

                cellStr = Left(cs, 2) & day

                ' Append if cell not empty
                If ws1.Range(cellStr).Value <> "" Then
                    ' Split and append stream
                    stream = Split(ws1.Range(cellStr).Value, ";", 2)(0) & stream
                End If
                
                ' Write to cell
                If InStr(stream, "OJT") <> 0 Then
                    stream = Trim(Replace(stream, "OJT", ""))
                    ws1.Range(cellStr).Value = (stream & ";" & Shift & ";Y;")
                Else
                    ws1.Range(cellStr).Value = (stream & ";" & Shift & ";N;")
                End If
                ws1.Range(cellStr).EntireColumn.Hidden = False

                ' Trainee
                If Len(cs) = 4 Then
                    ' OJTI cs
                    OJTI = Left(cs, 2)
                    ' OJT cs
                    cs = Right(cs, 2)

                    ' Append OJT to OJTI roster
                    ws1.Range(cellStr).Value = ws1.Range(cellStr).Value & cs & ";"
                    
                    ' OJT cell
                    cellStr = cs & day
                    ws1.Range(cellStr).Value = (stream & ";" & Shift & ";Y;" & OJTI & ";")
                    ws1.Range(cellStr).EntireColumn.Hidden = False
                Else
                    ' Append ; if no trainee
                    ws1.Range(cellStr).Value = ws1.Range(cellStr).Value & ";"
                End If
            End If
            i = i + 4
        Loop

        ' Office & Leave
        Do While ws2.Cells(i, result.Column).Value <> ""
            cs = ws2.Cells(i, result.Column).Value
            stream = ws2.Cells(i, firstDayCol.Column + day - 1).Value
            If stream <> "" Then
                cellStr = cs & day
                Shift = ""
                ' Append stream if cell not empty
                If ws1.Range(cellStr).Value <> "" Then
                    ' Split and append stream & shift
                    tmpStr = Split(ws1.Range(cellStr).Value, ";", 3)
                    If tmpStr(0) <> "" Then
                        If stream <> "See Note" Then
                            stream = tmpStr(0) & stream
                        Else
                            stream = tmpStr(0)
                        End If
                    End If
                    Shift = tmpStr(1)
                End If

                ' Write to cell and show
                ws1.Range(cellStr).Value = (stream & ";" & Shift & ";N;;")
                ws1.Range(cellStr).EntireColumn.Hidden = False
            End If
            i = i + 1
        Loop

    Next day
    
    ' =============== Names and Personal Notes ===============
    logFile.Write vbTab & "Notes"
    Set ws2 = wb2.Sheets("CALLSIGN")
    ws2.Unprotect ("AABABABABBAN")
    ws2.Columns.EntireColumn.Hidden = False
    ws2.Rows.EntireRow.Hidden = False
    
    ' Find personal notes range
    Set result = ws2.Range("1:1").Find("Personal Notes", LookIn:=xlValues)
    
    If Not result Is Nothing Then
        noteStart = result.Column
        
        If result.MergeCells Then
            Set rng = result.MergeArea
            noteEnd = noteStart + rng.Columns.Count - 1
        Else
            noteEnd = result.Column
        End If
        
        ' Debug.Print noteStart & " " & noteEnd
        
        ' Process line by line
        i = 2
        Do While Trim(ws2.Cells(i, 2).Value) <> ""
            cs = Trim(ws2.Cells(i, 1).Value)
            name = Trim(ws2.Cells(i, 2).Value)
            ' Add notes if found
            If cs <> "" Then
                notes = ws1.Range(cs & noteRow).Value
                For j = noteStart To noteEnd
                    ' Validate cell for string comparison
                    If Not Application.WorksheetFunction.IsNA(ws2.Cells(i, j)) Then
                        ' Determine stream
                        Set result = ws2.Range("1:1").Find("HKIA")
                        
                        'Set result = ws2.Cells(i, result.Column).Resize(1, 3)
                        Set result = ws2.Cells(i, result.Column)
                        
                        If WorksheetFunction.CountA(result) <> 0 Then
                            ws1.Range(cs & streamRow).Value = "APPRoster"
                        ElseIf WorksheetFunction.CountA(result.Offset(0, 1).Resize(1, 2)) <> 0 Then
                            ws1.Range(cs & streamRow).Value = "TMCRoster"
                        ElseIf WorksheetFunction.CountA(result.Offset(0, 3).Resize(1, 7)) <> 0 Then
                            ws1.Range(cs & streamRow).Value = "AREARoster"
                        ElseIf WorksheetFunction.CountA(result.Offset(0, 12).Resize(1, 1)) <> 0 Then
                            ws1.Range(cs & streamRow).Value = "TWRRoster"
                        ElseIf WorksheetFunction.CountA(result.Offset(0, -1).Resize(1, 1)) <> 0 Then
                            ws1.Range(cs & streamRow).Value = "APPRoster"
                        End If
                        
                        ' notes
                        tmpStr = Replace(Replace(ws2.Cells(i, j).Value, Chr(10), ""), Chr(13), "")
                        If InStr(tmpStr, "Individual notes are indicated on") <> 0 Then
                            ' Skip "Individual notes...' note
                        ElseIf tmpStr <> "" And tmpStr <> "0" Then
                            ' Trim numbering
                            If Mid(tmpStr, 2, 2) = ". " Then
                                tmpStr = Right(tmpStr, Len(tmpStr) - 3)
                            End If
                            notes = notes & "- " & Trim(Replace(Replace(tmpStr, "  ", " "), ",  ", ", ")) & ";"
                        Else
                            ' Skip if note cell is empty
                            Exit For
                        End If
                    End If
                Next j
            
                If notes = "" And WorksheetFunction.CountA(ws1.Range(cs & "1:" & cs & NumberOfDays)) = 0 Then
                    'Debug.Print RosterDate & " ATCO " & cs & " is empty"
                    ws1.Range(cs & "1").EntireColumn.Hidden = True
                Else
                    ' Write info to cells
                    ws1.Range(cs & nameRow).Value = name
                    ws1.Range(cs & noteRow).Value = notes
                    ws1.Range(cs & csRow).Value = cs
                    ws1.Range(cs & "1").EntireColumn.Hidden = False
                End If
            
            End If
           
            i = i + 1
        Loop
    Else
        Debug.Print "Personal Notes Cell not found"
    End If
    
    ' =============== Sick Leave ===============
    logFile.Write vbTab & "Sick"
    Set ws2 = wb2.Sheets("SICK")
    ws2.Unprotect ("AABABABABBAN")
    ws2.Columns.EntireColumn.Hidden = False
    ws2.Rows.EntireRow.Hidden = False
    sick = 0
    For day = 1 To NumberOfDays
        i = 4
        Do While Trim(ws2.Cells(i, day + 1).Value) <> ""
            cs = Left(ws2.Cells(i, day + 1).Value, 2)
            If ws1.Range(cs & day).Value <> "" Then
                ws1.Range(cs & day).Value = ws1.Range(cs & day).Value & "S;"
                sick = sick + 1
            End If
            i = i + 1
        Loop
    Next day
    #If developMode Then
        Debug.Print RosterDate & " Sick: " & sick
    #End If
    
    ' Close roster
    wb2.Close False
    logFile.WriteLine vbTab & "Done"
End Sub
Sub ATFSO(RosterDate As String)
    Dim filename As String
    Dim NumberOfDays As Integer
    Dim wb1, wb2 As Workbook
    Dim ws1, ws2 As Worksheet
    Dim result As Range
    Dim firstDayCol As Range
    Dim noteStart As Integer
    Dim noteEnd As Integer
    Dim rng As Range
    Dim day As Integer
    Dim cs As String
    Dim name As String
    Dim Shift As String
    Dim stream As String
    Dim entryStr As String
    Dim OJTI As String
    Dim cellStr As String
    Dim tmpStr As Variant
    Dim notes As String
    Dim t As Single
    
    Dim ATFSONewer As Boolean
    
    filename = ATFSOPath & RosterDate & "*.xlsx"
    
    If Dir(filename) = "" Then
        Debug.Print filename & " not found"
        Exit Sub
    End If
    
    filename = ATFSOPath & Dir(filename)
    
    ' Check ATCO or ATFSO file is newer for deconfliction
    If Dir(ATCOPath & RosterDate & ".xlsx") = "" Then
        ATFSONewer = True
    ElseIf FileDateTime(filename) > FileDateTime(ATCOPath & RosterDate & ".xlsx") Then
        ATFSONewer = True
    Else
        ATFSONewer = False
    End If
    
    ' Close File if already opened
    On Error Resume Next
    Workbooks(filename).Close SaveChanges:=False
    On Error GoTo 0
    
    Set wb1 = ThisWorkbook
    
    t = Timer()
    
    logFile.Write vbTab & "Open"
    Set wb2 = Workbooks.Open(filename:=filename, Password:="aerostar", UpdateLinks:=0)
    
    OpenTime = OpenTime + Timer() - t
    
    wb2.Windows(1).Visible = False
    
    Set ws1 = wb1.Sheets(RosterDate)
    
    ' Unprotect and unhide sheets
    ' TODO: brute force if default password fail
    Set ws2 = wb2.Sheets("MASTER")
    ws2.Unprotect ("AAABABBABBBo")
    ws2.Rows.EntireRow.Hidden = False
    ws2.Columns.EntireColumn.Hidden = False
    
    Set result = ws2.Range("B:B").Find("fss", LookIn:=xlValues)
    Set firstDayCol = ws2.UsedRange.Find(Format("1 " & RosterDate, "d-mmm"))
    
    If Not firstDayCol Is Nothing Then
        ' Debug.Print "firstRow: " & firstDayCol.Column
    End If
    
    If Not result Is Nothing Then
        ' Debug.Print result.Address
    Else
        Debug.Print "result not found"
    End If
    
    NumberOfDays = NB_DAYS(DateValue("1 " & RosterDate))
    
    logFile.Write vbTab & "Shifts"
    For day = 1 To NumberOfDays
        ' Normal roster
        i = result.Row

        Do While ws2.Cells(i, result.Column).Value <> "C/S 1" ' Loop til other mannings
            cs = UCase(Trim(ws2.Cells(i, firstDayCol.Column + day - 1).Value))
            If IsAlpha(cs) And (Len(cs) = 4 Or Len(cs) = 2) Then
                Shift = ws2.Cells(i, 1).Value
                stream = UCase(Trim(ws2.Cells(i, result.Column).Value))
                
                entryStr = stream & ";" & Shift & ";N;"
                
                cellStr = Left(cs, 2) & day
                
                ws1.Range(cellStr).Value = ShiftDeconf(ws1.Range(cellStr).Value, entryStr, ATFSONewer, cs, day)
                ws1.Range(cellStr).EntireColumn.Hidden = False

                ' Trainee
                If Len(cs) = 4 Then
                    ' OJTI cs
                    OJTI = Left(cs, 2)
                    ' OJT cs
                    cs = Right(cs, 2)
                    
                    ' Append OJT to OJTI roster
                    ws1.Range(cellStr).Value = ws1.Range(cellStr).Value & cs & ";"
                    
                    ' OJT cell
                    cellStr = cs & day
                    entryStr = stream & ";" & Shift & ";Y;" & OJTI & ";"
                    
                    ws1.Range(cellStr).Value = ShiftDeconf(ws1.Range(cellStr).Value, entryStr, ATFSONewer, cs, day)
                    ws1.Range(cellStr).EntireColumn.Hidden = False
                Else
                    ' Append ; if no trainee
                    If Right(ws1.Range(cellStr).Value, 3) <> ";S;" Then
                        ws1.Range(cellStr).Value = ws1.Range(cellStr).Value & ";"
                    End If
                End If
            End If
        i = i + 1
        Loop

        ' Other manning
        Do While Len(ws2.Cells(i, result.Column).Value) <> 2 ' Loop til callsign rows
            cs = Trim(ws2.Cells(i, firstDayCol.Column + day - 1).Value)
            If IsAlpha(cs) And (Len(cs) = 4 Or Len(cs) = 2) Then
                stream = Trim(ws2.Cells(i + 1, firstDayCol.Column + day - 1).Value)
                Shift = Trim(ws2.Cells(i + 2, firstDayCol.Column + day - 1).Value)

                cellStr = Left(cs, 2) & day

                ' Write to cell
                
                If InStr(stream, "OJT") <> 0 Then
                    entryStr = (Trim(Replace(stream, "OJT", "")) & ";" & Shift & ";Y;")
                Else
                    entryStr = (stream & ";" & Shift & ";N;")
                End If
                
                ws1.Range(cellStr).Value = ShiftDeconf(ws1.Range(cellStr).Value, entryStr, ATFSONewer, cs, day)
                ws1.Range(cellStr).EntireColumn.Hidden = False

                ' Trainee
                If Len(cs) = 4 Then
                    ' OJTI cs
                    OJTI = Left(cs, 2)
                    ' OJT cs
                    cs = Right(cs, 2)
                    
                    ' Append OJT to OJTI roster
                    ws1.Range(cellStr).Value = ws1.Range(cellStr).Value & cs & ";"
                    
                    ' OJT cell
                    cellStr = cs & day
                    entryStr = (stream & ";" & Shift & ";Y;" & OJTI & ";")
                    
                    ws1.Range(cellStr).Value = ShiftDeconf(ws1.Range(cellStr).Value, entryStr, ATFSONewer, cs, day)
                    ws1.Range(cellStr).EntireColumn.Hidden = False
                Else
                    ' Append ; if no trainee
                    If Right(ws1.Range(cellStr).Value, 3) <> ";S;" Then
                        ws1.Range(cellStr).Value = ws1.Range(cellStr).Value & ";"
                    End If
                End If
            End If
            i = i + 4
        Loop

        ' Office & Leave
        Do While ws2.Cells(i, result.Column).Value <> ""
            cs = ws2.Cells(i, result.Column).Value
            stream = Trim(ws2.Cells(i, firstDayCol.Column + day - 1).Value)
            
            ' If stream exists and not whole month = "Course" then proceed to parsing
            If stream <> "" And Not (stream = "Course" And Application.CountIf(ws2.Cells(i, firstDayCol.Column).Resize(1, NumberOfDays), "Course") = NumberOfDays) Then
                cellStr = cs & day
                Shift = ""
                ' Append stream if cell not empty
                If ws1.Range(cellStr).Value <> "" Then
                    ' Split and append stream & shift
                    tmpStr = Split(ws1.Range(cellStr).Value, ";", 3)
                    If tmpStr(0) <> stream And tmpStr(0) <> "" And stream <> "See Note" Then
                        stream = tmpStr(0) & stream
                    Else
                        stream = tmpStr(0)
                    End If
                    Shift = tmpStr(1)
                    ' Debug.Print day & " " & RosterDate & vbTab & cs & ": " & (stream & ";" & shift & ";N;")
                End If

                ' Write to cell and show
                entryStr = (stream & ";" & Shift & ";N;;")
                ws1.Range(cellStr).Value = ShiftDeconf(ws1.Range(cellStr).Value, entryStr, ATFSONewer, cs, day)
                ws1.Range(cellStr).EntireColumn.Hidden = False
                
                'ws1.Range(cellStr).Value = (stream & ";" & Shift & ";N;;")
                'ws1.Range(cellStr).EntireColumn.Hidden = False
            End If
            i = i + 1
        Loop

    Next day
    
    ' =============== Names and Personal Notes ===============
    logFile.Write vbTab & "Notes"
    Set ws2 = wb2.Sheets("CALLSIGN")
    ws2.Unprotect ("AAABABBABBBo")
    ws2.Columns.EntireColumn.Hidden = False
    ws2.Rows.EntireRow.Hidden = False
    
    ' Find personal notes range
    Set result = ws2.Range("1:1").Find("Personal Notes", LookIn:=xlValues)
    
    If Not result Is Nothing Then
        noteStart = result.Column
        
        If result.MergeCells Then
            Set rng = result.MergeArea
            noteEnd = noteStart + rng.Columns.Count - 1
        Else
            noteEnd = result.Column
        End If
        
        ' Debug.Print noteStart & " " & noteEnd
  
        ' Process line by line
        i = 2
        Do While Trim(ws2.Cells(i, 2).Value) <> ""
            cs = Trim(ws2.Cells(i, 1).Value)
            name = Trim(ws2.Cells(i, 2).Value)
            
            ' Add notes if found
            If cs <> "" Then
                notes = ws1.Range(cs & noteRow).Value
                For j = noteStart To noteEnd
                    If Not Application.WorksheetFunction.IsNA(ws2.Cells(i, j)) Then
                        If InStr(ws2.Cells(i, j).Value, "See ATCO Watchlist for other duties") <> 0 Then
                            ' Skip
                        ElseIf ws2.Cells(i, j).Value <> "" And ws2.Cells(i, j).Value <> "0" Then
                            ' tmpStr = ws2.Cells(i, j).Value
                            tmpStr = Replace(Replace(ws2.Cells(i, j).Value, Chr(10), ""), Chr(13), "")
                            ' Replace emdash
                            tmpStr = Replace(tmpStr, ChrW(8212), "")
                            notes = notes & "- " & Trim(Replace(Replace(tmpStr, "  ", " "), ",  ", ", ")) & ";"
                        Else
                            ' Skip if note cell is empty
                            Exit For
                        End If
                    Else
                        Exit For
                    End If
                Next j
                
                If notes = "" And WorksheetFunction.CountA(ws1.Range(cs & "1:" & cs & NumberOfDays)) = 0 Then
                    ' Debug.Print RosterDate & " ATFSO " & cs & " is empty"
                    ws1.Range(cs & "1").EntireColumn.Hidden = True
                Else
                    If Len(ws1.Range(cs & nameRow).Value) < Len(name) Then
                        ' Use longer name
                        ws1.Range(cs & nameRow).Value = name
                    End If
                    ws1.Range(cs & noteRow).Value = notes
                    ws1.Range(cs & csRow).Value = cs
                    ws1.Range(cs & "1").EntireColumn.Hidden = False
                End If
                
            End If
            
            i = i + 1
        Loop
    Else
        Debug.Print "Personal Notes Cell not found"
    End If
    
    ' =============== Sick Leave ===============
    logFile.Write vbTab & "Sick"
    Set ws2 = wb2.Sheets("SICK")
    ws2.Unprotect ("AAABABBABBBo")
    ws2.Columns.EntireColumn.Hidden = False
    ws2.Rows.EntireRow.Hidden = False

    sick = 0
    For day = 1 To NumberOfDays
        i = 4
        Do While Trim(ws2.Cells(i, day + 1).Value) <> ""
            cs = Left(ws2.Cells(i, day + 1).Value, 2)
            If ws1.Range(cs & day).Value <> "" And Right(ws1.Range(cs & day).Value, 3) <> ";S;" Then
                ws1.Range(cs & day).Value = ws1.Range(cs & day).Value & "S;"
                sick = sick + 1
            End If
            i = i + 1
        Loop
    Next day

    #If developMode Then
        Debug.Print RosterDate & " Sick: " & sick
    #End If
    
    ' Close roster
    wb2.Close False
    logFile.WriteLine vbTab & "Done"
End Sub

Sub ProcessRoster(RosterDate As String)
    Dim ws1 As Worksheet
  
    ' Add Month worksheet
    With ThisWorkbook
        On Error Resume Next
        ' Remove sheet
        .Sheets(RosterDate).Delete
        On Error GoTo 0
        Set ws1 = .Sheets.Add(After:=.Sheets(.Sheets.Count))
        ws1.name = RosterDate
        ws1.Columns("A:ZZ").EntireColumn.Hidden = True
    End With
    
    logFile.Write Now & vbTab & "ATCO"
    ATCO (RosterDate)
    
    logFile.Write Now & vbTab & "ATFSO"
    ATFSO (RosterDate)

End Sub

Sub OneClick()
    Call OptimizeCode_Begin
    
    Dim Month1 As Date
    Dim Month2 As Date
    Dim MonthToProcess(2) As Variant
    Dim MonthSheet(2) As Worksheet
    Dim RosterFile(2, 2) As String
    Dim RosterModTime(2, 2) As String
    Dim RosterVersion(2) As String
    Dim ATCOfile As String
    Dim ATFSOfile As String
    Dim cs As String
    Dim ScriptStart As Single
    Dim ScriptEnd As Single
    Dim RosterRange As Range
    Dim searchRng As Range
    Dim search As Range
    
    Dim MonthHeader As String
    Dim PersonalRoster As String
    Dim VersionTxt As String
    Dim RosterTxt As String
    Dim VersionTmp As String
    Dim RosterTmp As String
    Dim NumberOfDays As Integer
    
    Dim objFSO As Object
    Dim logFileName As String
    
    ' Init log file
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    logFileName = ThisWorkbook.path & "\log.txt"
    
    If Dir(logFileName) = "" Then
        Set logFile = objFSO.CreateTextFile(logFileName)
    Else
        Set logFile = objFSO.OpenTextFile(logFileName, ForAppending)
    End If
    
    ' On Error GoTo closeLog
    logFile.WriteLine
    logFile.WriteLine "=========================================================================="
    logFile.WriteLine Now & vbTab & "Roster conversion started"
    
    ScriptStart = Timer()
    
    VersionTxt = ThisWorkbook.path & "\ATCapp_Roster_Version.txt"
    RosterTxt = ThisWorkbook.path & "\ATCapp_Rosters_new.txt"
    VersionTmp = ThisWorkbook.path & "\ATCapp_Roster_Version.tmp"
    RosterTmp = ThisWorkbook.path & "\ATCapp_Rosters_new.tmp"
    
    ' Determine Files to process
    ' Default to process current and next month
    Month1 = DateSerial(Year(Now), Month(Now), 1)
    Month2 = DateAdd("m", 1, Month1)

    ' Test paths
    ATCOfile = ATCOPath & Format(Month2, "mmmm yyyy") & "*.xlsx"
    ATFSOfile = ATFSOPath & Format(Month2, "mmmm yyyy") & "*.xlsx"
    
    'Debug.Print ATCOfile & " " & ATFSOfile
    
    ' If both files for next month does not exists, process current and previous month
    If Dir(ATCOfile) = "" And Dir(ATFSOfile) = "" Then
        ' Adjust months
        Month2 = Month1
        Month1 = DateAdd("m", -1, Month2)
    End If
    
    MonthToProcess(0) = Format(Month1, "mmmm yyyy")
    MonthToProcess(1) = Format(Month2, "mmmm yyyy")
   
    For i = 0 To 1
        RosterFile(i, 0) = Dir(ATFSOPath & MonthToProcess(i) & "*.xlsx")
        RosterFile(i, 1) = Dir(ATCOPath & MonthToProcess(i) & "*.xlsx")
        
        If RosterFile(i, 0) <> "" Then
            Debug.Print RosterFile(i, 0)
            RosterModTime(i, 0) = Format(FileDateTime(ATFSOPath & RosterFile(i, 0)), "dd/mm/yyyy HH:nn")
        Else
            RosterModTime(i, 0) = ""
        End If
        
        If RosterFile(i, 1) <> "" Then
            Debug.Print RosterFile(i, 1)
            RosterModTime(i, 1) = Format(FileDateTime(ATCOPath & RosterFile(i, 1)), "dd/mm/yyyy HH:nn")
        Else
            RosterModTime(i, 1) = ""
        End If
    Next i
    
    ' Clear and open temp files
    KillProperly (RosterTmp)
    Open RosterTmp For Output As #1
    
    KillProperly (VersionTmp)
    Open VersionTmp For Output As #2
    
    For i = 0 To 1
        
        logFile.WriteLine Now & vbTab & "Processing " & MonthToProcess(i)
        ' Process Rosters
        ProcessRoster (MonthToProcess(i))
        
        ' Roster version & time
        RosterVersion(i) = RosterModTime(i, 0) & ";" & RosterModTime(i, 1)
        Print #2, RosterVersion(i)
        
        Set MonthSheet(i) = ThisWorkbook.Sheets(MonthToProcess(i))

        MonthHeader = "Roster:" & Replace(MonthToProcess(i), " ", ";") & ";" & RosterVersion(i) & ";"
        Print #1, MonthHeader & vbNewLine & vbNewLine
        
        NumberOfDays = NB_DAYS(DateValue("1 " & MonthToProcess(i)))
        
        Set RosterRange = MonthSheet(i).Range("AA1:ZZ" & streamRow)
        logFile.WriteLine Now & vbTab & "Writing to file"
        With RosterRange
            For k = 1 To .Columns.Count
                ' If CS exists
                If .Cells(csRow, k).Value <> "" Then
                    cs = .Cells(csRow, k).Value
                    ' Determine stream
                    Set searchRng = MonthSheet(i).Range(.Cells(1, k), .Cells(NumberOfDays, k))
                    
                    'If Application.CountIf(searchRng, "SOHD;;N;") = NumberOfDays And .Cells(noteRow, K) = "" Then
                        'Debug.Print cs & " all SOHD & no notes"
                    'Else
                    If True Then
                       
                        ' Set "Roster"
                        If Application.CountIf(searchRng, "*;Y;*") > 0 Then
                            If Application.CountIf(searchRng, "APP*") > 0 Or Application.CountIf(searchRng, "APS*") > 0 Then
                                .Cells(streamRow, k).Value = "APPRoster"
                            ElseIf Application.CountIf(searchRng, "TMC*") > 0 Or Application.CountIf(searchRng, "TSU*") > 0 Then
                                .Cells(streamRow, k).Value = "TMCRoster"
                            ElseIf Application.CountIf(searchRng, "AREA*") > 0 Or Application.CountIf(searchRng, "ESU*") > 0 Then
                                .Cells(streamRow, k).Value = "AREARoster"
                            ElseIf Application.CountIf(searchRng, "WMR*") > 0 Or Application.CountIf(searchRng, "FLM*") > 0 Then
                                .Cells(streamRow, k).Value = "AREARoster"
                            ElseIf Application.CountIf(searchRng, "TWR*") > 0 Or Application.CountIf(searchRng, "ASU*") > 0 Then
                                .Cells(streamRow, k).Value = "TWRRoster"
                            ElseIf .Cells(streamRow, k).Value = "" Then
                                .Cells(streamRow, k).Value = "ATFSO"
                            End If
                        End If
                        
                        ' Manual entries
                        MonthSheet(i).Range("SL" & streamRow).Value = "AREARoster"
                        MonthSheet(i).Range("SM" & streamRow).Value = "APPRoster"
                        MonthSheet(i).Range("XR" & streamRow).Value = "TWRRoster"
                        MonthSheet(i).Range("LW" & streamRow).Value = "AREARoster"
                        
                        ' Replace Night CDC with Night TWR for Rated controllers
                        If .Cells(streamRow, k).Value Like "*Roster" Then
                            If Application.CountIf(searchRng, "TWR;*;Y;*") = 0 Then
                                searchRng.Replace What:="CDC;N", Replacement:="TWR;N"
                            End If
                        End If
                        
                        ' Set empty streams with ATFSO and replace their TWR with CDC
                        If .Cells(streamRow, k) = "" Then
                            .Cells(streamRow, k).Value = "ATFSO"
                            searchRng.Replace What:="TWR;", Replacement:="CDC;"
                        End If
                       
                        PersonalRoster = "Name:" & .Cells(nameRow, k) & ";" & .Cells(csRow, k) & ";" & .Cells(streamRow, k) & ";" & .Cells(noteRow, k)
                        Print #1, PersonalRoster
                        
                        For j = 1 To NumberOfDays
                            If .Cells(j, k).Value = "" Then
                               ' Fill Empty days
                               .Cells(j, k).Value = ";;N;;"
                               Print #1, ";;N;;"
                            Else
                               Print #1, .Cells(j, k).Value
                            End If
                        Next j
                        Print #1, ""
                    End If
                End If
            Next k
        End With
        
        ' Remove completed roster worksheets
        #If Not developMode Then
            On Error Resume Next
            MonthSheet(i).Delete
            On Error GoTo 0
        #End If
        
    Next i
    
    ' Close temp files
    Close #1
    Close #2
    
    ' Delete and rename files after conversion completed successfully
    KillProperly (RosterTxt)
    KillProperly (VersionTxt)
    Name RosterTmp As RosterTxt
    Name VersionTmp As VersionTxt
    
    ScriptEnd = Timer()
    ElapsedTime = ScriptEnd - ScriptStart
    Debug.Print "Elapsed Time: " & Format(ElapsedTime, "#.00") & "s"
    ThisWorkbook.Sheets("Main").Range("A3") = Format(ElapsedTime, "#.00") & "s"
    
    logFile.WriteLine Now & vbTab & "Elapsed time: " & ElapsedTime & "s"
    logFile.WriteLine Now & vbTab & "Roster conversion completed"

closeLog:
    ' Close log file
    logFile.Close
    
    Call OptimizeCode_End
End Sub

Sub PerformanceTest()
    AverageTime = 0
    For i = 1 To 5
        Call OneClick
        AverageTime = (AverageTime * (i - 1) + ElapsedTime) / i
    Next i
    Debug.Print "Average Time: " & Format(AverageTime, "#.00") & "s"
    ThisWorkbook.Sheets("Main").Range("A3") = Format(AverageTime, "#.00") & "s"
End Sub

Sub ProcessTime()
    OpenTime = 0
    Call OneClick
    Debug.Print "Process time: " & Format(ElapsedTime - OpenTime, "#.00")
End Sub

Public Sub KillProperly(Killfile As String)
    If Len(Dir$(Killfile)) > 0 Then
        SetAttr Killfile, vbNormal
        Kill Killfile
    End If
End Sub

Public Sub RegEx_Tester()

Set objRegExp_1 = CreateObject("vbscript.regexp")

objRegExp_1.Global = True

objRegExp_1.IgnoreCase = True

objRegExp_1.Pattern = "[a-z,A-Z]*@[a-z,A-Z]*.com."

strToSearch = "ABC@xyz.com"

Set regExp_Matches = objRegExp_1.Execute(strToSearch)

If regExp_Matches.Count = 1 Then

MsgBox ("This string is a valid email address.")
Else
MsgBox ("This string is not a valid email address.")
End If

End Sub