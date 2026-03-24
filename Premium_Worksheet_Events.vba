'*******************************************************************************
' Worksheet Event Module - Double-click toggle, header mapping, and change sync
' This code should be placed in the worksheet module (e.g., Sheet1)
' Version 17.2 - Fixed debug constant for standalone operation
'*******************************************************************************

Option Explicit

'===============================================================================
' Debug constant - local to this module for standalone operation
'===============================================================================
Private Const DEBUG_MODE As Boolean = True

'===============================================================================
' Debug helper - prints to Immediate Window
'===============================================================================
Private Sub DebugPrint(message As String)
    #If DEBUG_MODE Then
        On Error Resume Next
        Debug.Print "[" & Format(Now, "hh:nn:ss") & "] " & message
    #End If
End Sub

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    Dim ws As Worksheet
    Dim tempAliases As Object
    Dim tempLearned As Object
    Dim headerRow As Long
    Dim lastCol As Long
    Dim matchRow As Long
    Dim response As Integer
    Dim matchEndRow As Long
    Dim matchStartRow As Long

    ' Guard: prevent crash when multiple cells are selected
    If Target.Cells.Count > 1 Then Exit Sub

    On Error Resume Next

    Set ws = Me

    ' CRITICAL GUARD: If not initialized, initialize now
    If Not g_Initialized Then
        DebugPrint "Worksheet_BeforeDoubleClick: Not initialized, calling InitializeDatasetContext..."
        Call InitializeDatasetContext(ws)
        If Not g_Initialized Then
            DebugPrint "Worksheet_BeforeDoubleClick: Still not initialized after init call"
            Exit Sub
        End If
    End If

    ' Use global header row
    headerRow = g_DataHeaderRow
    lastCol = g_LastDataColumn

    ' Fallback detection if needed
    If headerRow = 0 Then
        Set tempAliases = CreateObject("Scripting.Dictionary")
        Set tempLearned = CreateObject("Scripting.Dictionary")
        headerRow = FindDataHeaderRow(ws, tempAliases, tempLearned)
        If headerRow <= 1 Then
            headerRow = FindDataHeaderRowAggressive(ws, tempAliases, tempLearned)
        End If
        If headerRow > 1 Then
            g_DataHeaderRow = headerRow
            g_DataStartRow = headerRow + 1
        End If
    End If

    If lastCol = 0 Then
        lastCol = GetLastColumn(ws, headerRow)
    End If

    On Error GoTo 0

    ' Gracefully handle if detection failed
    If headerRow = 0 Then headerRow = 10  ' Fallback
    If lastCol = 0 Then lastCol = 10  ' Default fallback

    ' Calculate dynamic UI boundaries
    matchStartRow = 5  ' Match rows start at row 5
    If g_DataHeaderRow > 0 Then
        matchEndRow = g_DataHeaderRow - 1
    Else
        matchEndRow = 9  ' Fallback: FIXED_DATA_HEADER_ROW(10) - 1, literal used as constant not accessible from worksheet module
    End If
    If matchEndRow < matchStartRow Then matchEndRow = matchStartRow + 2

    ' Check if double-clicked on the DATA HEADER ROW - assign column to match rule
    If Target.Row = headerRow And Target.Column >= 1 And Target.Column <= lastCol Then
        Cancel = True

        ' Find the first match rule that doesn't have this column selected
        For matchRow = matchStartRow To matchEndRow
            If IsNumeric(ws.Cells(matchRow, 1).Value) Then
                ' Check if this column is already marked
                If Trim(UCase(ws.Cells(matchRow, Target.Column).Value)) <> "X" Then
                    response = MsgBox("Use column '" & Target.Value & "' for Match Rule " & _
                            ws.Cells(matchRow, 1).Value & " (" & ws.Cells(matchRow, 2).Value & ")?" & vbCrLf & vbCrLf & _
                            "Click Yes to assign, No to skip to next rule.", vbYesNoCancel + vbQuestion)

                    If response = vbYes Then
                        Application.EnableEvents = False
                        ws.Cells(matchRow, Target.Column).Value = "X"
                        ws.Cells(matchRow, Target.Column).Interior.Color = RGB(255, 0, 0)
                        Application.EnableEvents = True
                        Exit Sub
                    ElseIf response = vbCancel Then
                        Exit Sub
                    End If
                    ' If No, continue to next match rule
                End If
            End If
        Next matchRow

        MsgBox "All match rules already have column '" & Target.Value & "' assigned.", vbInformation
        Exit Sub
    End If

    ' Check if double-clicked on a match row - toggle X mark
    ' CRITICAL: Use dynamic matchStartRow instead of hardcoded UI_FIRST_MATCH_ROW
    If Target.Row < matchStartRow Or Target.Row >= headerRow Then Exit Sub
    If Target.Column < 3 Or Target.Column > lastCol Then Exit Sub

    Cancel = True

    ' Toggle X mark
    Application.EnableEvents = True  ' Safety: ensure re-enabled even if prior error left it disabled
    Application.EnableEvents = False
    If Trim(UCase(Target.Value)) = "X" Then
        Target.ClearContents
        Target.Interior.Color = RGB(255, 255, 200)
    Else
        Target.Value = "X"
        Target.Interior.Color = RGB(255, 0, 0)
    End If
    Application.EnableEvents = True
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim ws As Worksheet
    Dim headerRow As Long
    Dim lastCol As Long
    Dim cell As Range
    Dim matchStartRow As Long
    Dim matchEndRow As Long

    ' Exit immediately when a large range changes (like column deletion)
    If Target.Columns.Count > 10 Then Exit Sub

    ' Guard: Prevent event from running during macro operations
    If g_MacroRunning Then Exit Sub

    On Error Resume Next

    Set ws = Me

    ' CRITICAL GUARD: If not initialized, initialize now
    If Not g_Initialized Then
        DebugPrint "Worksheet_Change: Not initialized, calling InitializeDatasetContext..."
        Call InitializeDatasetContext(ws)
        If Not g_Initialized Then
            DebugPrint "Worksheet_Change: Still not initialized after init call"
            Exit Sub
        End If
    End If

    ' Use global header row
    headerRow = g_DataHeaderRow
    lastCol = g_LastDataColumn

    ' Fallback detection if needed
    If headerRow = 0 Then
        Dim tempAliases As Object
        Dim tempLearned As Object
        Set tempAliases = CreateObject("Scripting.Dictionary")
        Set tempLearned = CreateObject("Scripting.Dictionary")
        headerRow = FindDataHeaderRow(ws, tempAliases, tempLearned)
        If headerRow <= 1 Then
            headerRow = FindDataHeaderRowAggressive(ws, tempAliases, tempLearned)
        End If
        If headerRow > 1 Then
            g_DataHeaderRow = headerRow
            g_DataStartRow = headerRow + 1
        End If
    End If

    If lastCol = 0 Then
        lastCol = GetLastColumn(ws, headerRow)
    End If

    On Error GoTo 0

    ' Gracefully handle if detection failed
    If headerRow = 0 Then headerRow = 10
    If lastCol = 0 Then lastCol = 10

    ' Determine the effective match area based on dynamic g_DataHeaderRow
    ' Use dynamic boundaries instead of constants
    matchStartRow = 5  ' Match rows start at row 5
    If g_DataHeaderRow > 0 Then
        matchEndRow = g_DataHeaderRow - 1
    Else
        matchEndRow = 9  ' Fallback: FIXED_DATA_HEADER_ROW(10) - 1, literal used as constant not accessible from worksheet module
    End If
    If matchEndRow < matchStartRow Then matchEndRow = matchStartRow + 2

    ' CRITICAL SAFETY CHECK: Validate the range before processing
    If matchStartRow >= matchEndRow Or matchEndRow >= headerRow Then
        Exit Sub  ' UI not properly set up, skip
    End If

    If Intersect(Target, ws.Range(ws.Cells(matchStartRow, 3), _
                                   ws.Cells(matchEndRow, lastCol))) Is Nothing Then Exit Sub

    ' Update colors - but only within valid match area
    Application.EnableEvents = False
    For Each cell In Target.Cells
        ' Additional safety check for each cell
        If cell.Row >= matchStartRow And cell.Row <= matchEndRow And _
           cell.Column >= 3 And cell.Column <= lastCol Then
            If Trim(UCase(cell.Value)) = "X" Then
                cell.Interior.Color = RGB(255, 0, 0)
            Else
                cell.Interior.Color = RGB(255, 255, 200)
            End If
        End If
    Next cell
    Application.EnableEvents = True
End Sub
