Attribute VB_Name = "Module1"
Sub Import_VPSRG_HN()
Attribute Import_VPSRG_HN.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Import_VPSRG_HN Macro
'

'
    ' Save workbook before importing data, just in case
    ActiveWorkbook.Save
        
    ' Open the file dialog
    With Application.FileDialog(msoFileDialogFilePicker)
        .AllowMultiSelect = False
        .InitialFileName = "c:\temp\*.txt"
        .Show

        my_path = "TEXT;" + .SelectedItems(1)
        
    End With
    
    prev_conn_count = ActiveWorkbook.Connections.Count
    workRow = ActiveCell.Row
    ActiveSheet.Cells(workRow, 4).Activate
                
    With ActiveSheet.QueryTables.Add(Connection:= _
        my_path, Destination:=ActiveCell)
        .Name = my_path
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlOverwriteCells
        .SavePassword = False
        .SaveData = False
        .AdjustColumnWidth = False
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = 65001
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = False
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = True
        .TextFileSpaceDelimiter = False
        .TextFileColumnDataTypes = Array(1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, _
        1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With
    
    
    ' Remove the connection
    ActiveWorkbook.Connections(prev_conn_count + 1).Delete
    
    'centering data
    ActiveSheet.Rows(workRow).HorizontalAlignment = xlCenter
    
    ' cleaning NaNs
    For i = 19 To 52
    If ActiveSheet.Cells(workRow, i).Value = "NaN" Then
    ActiveSheet.Cells(workRow, i).Value = ""
    End If
    Next
    
        
End Sub

Sub compare_values()
'compare values
    ' green if Ok, orange inside 2%, red more than 2%
    workRow = ActiveCell.Row
    
    'Col S (19): Brainstem Dmax<54Gy
    ActiveSheet.Cells(workRow, 19).Select
    compVal = 54 ' Gy
    If Selection.Value > compVal * 1.02 Then
    red
    ElseIf Selection.Value >= compVal * 0.98 Then
    yellow
    Else
    green
    End If
    
    'Col T (20): Cord Dmax<45Gy
    ActiveSheet.Cells(workRow, 20).Select
    compVal = 45 ' Gy
    If Selection.Value > compVal * 1.02 Then
    red
    ElseIf Selection.Value >= compVal * 0.98 Then
    yellow
    Else
    green
    End If
    
    
    'Col U (21): Cord PRV Dmax<45Gy
    ActiveSheet.Cells(workRow, 21).Select
    compVal = 45 ' Gy
    If Selection.Value > compVal * 1.02 Then
    red
    ElseIf Selection.Value >= compVal * 0.98 Then
    yellow
    Else
    green
    End If
    
    'Col V (22): Larynx Dmean .... NOT DEFINED
    'ActiveSheet.Cells(workRow, 22).Select
    'compVal = 45 ' Gy
    'If Selection.Value > 45 * 1.02 Then
    'red
    'ElseIf Selection.Value >= 45 * 0.98 Then
    'yellow
    'Else
    'green
    'End If
    
    'Col W (23): Inner ear Lt Dmax<50Gy
    ActiveSheet.Cells(workRow, 23).Select
    compVal = 50 ' Gy
    If Selection.Value > compVal * 1.02 Then
    red
    ElseIf Selection.Value >= compVal * 0.98 Then
    yellow
    Else
    green
    End If
    
    'Col X (24): Inner ear Rt Dmax<50Gy
    ActiveSheet.Cells(workRow, 24).Select
    compVal = 50 ' Gy
    If Selection.Value > compVal * 1.02 Then
    red
    ElseIf Selection.Value >= compVal * 0.98 Then
    yellow
    Else
    green
    End If
    
    'Col Y (25): Lens Lt Dmax<8Gy
    ActiveSheet.Cells(workRow, 25).Select
    compVal = 8 ' Gy
    If Selection.Value > compVal * 1.02 Then
    red
    ElseIf Selection.Value >= compVal * 0.98 Then
    yellow
    Else
    green
    End If
    
    'Col Z (26): Lens Rt Dmax<8Gy
    ActiveSheet.Cells(workRow, 26).Select
    compVal = 8 ' Gy
    If Selection.Value > compVal * 1.02 Then
    red
    ElseIf Selection.Value >= compVal * 0.98 Then
    yellow
    Else
    green
    End If
    
    'Col AA (27): MAndible Dmax report only
    'ActiveSheet.Cells(workRow, 27).Select
    'compVal = 8 ' Gy
    'If Selection.Value > compVal * 1.02 Then
    'red
    'ElseIf Selection.Value >= compVal * 0.98 Then
    'yellow
    'Else
    'green
    'End If
    
    'Col AB (28): Mandible V{TotalDose}[%] < 1%
    ActiveSheet.Cells(workRow, 28).Select
    compVal = 1 ' %
    If Selection.Value > compVal * 1.02 Then
    red
    ElseIf Selection.Value >= compVal * 0.98 Then
    yellow
    Else
    green
    End If
    
    'Col AC (29): Optic Chiasm Dmax<54Gy
    ActiveSheet.Cells(workRow, 29).Select
    compVal = 54 ' Gy
    If Selection.Value > compVal * 1.02 Then
    red
    ElseIf Selection.Value >= compVal * 0.98 Then
    yellow
    Else
    green
    End If
    
    'Col AD (30): Optic Nerve Lt Dmax<54Gy
    ActiveSheet.Cells(workRow, 30).Select
    compVal = 54 ' Gy
    If Selection.Value > compVal * 1.02 Then
    red
    ElseIf Selection.Value >= compVal * 0.98 Then
    yellow
    Else
    green
    End If
    
    'Col AE (31): Optic Nerve Rt Dmax<54Gy
    ActiveSheet.Cells(workRow, 31).Select
    compVal = 54 ' Gy
    If Selection.Value > compVal * 1.02 Then
    red
    ElseIf Selection.Value >= compVal * 0.98 Then
    yellow
    Else
    green
    End If
    
    'Col AF (32): Uninvolved Oral Cavity Dmean<30Gy
    ActiveSheet.Cells(workRow, 32).Select
    compVal = 30 ' Gy
    If Selection.Value > compVal * 1.02 Then
    red
    ElseIf Selection.Value >= compVal * 0.98 Then
    yellow
    Else
    green
    End If
    
    'Col AG (33): PArotid Lt Dmean<26Gy
    ActiveSheet.Cells(workRow, 33).Select
    compVal = 26 ' Gy
    If Selection.Value > compVal * 1.02 Then
    red
    ElseIf Selection.Value >= compVal * 0.98 Then
    yellow
    Else
    green
    End If
    
    'Col AH (34): PArotid Lt V30Gy[%]<50%
    ActiveSheet.Cells(workRow, 34).Select
    compVal = 50 ' %
    If Selection.Value > compVal * 1.02 Then
    red
    ElseIf Selection.Value >= compVal * 0.98 Then
    yellow
    Else
    green
    End If
    
    'Col AI (35): PArotid Lt V20Gy[cm3]<20cm3
    ActiveSheet.Cells(workRow, 35).Select
    compVal = 20 ' cm3
    If Selection.Value > compVal * 1.02 Then
    red
    ElseIf Selection.Value >= compVal * 0.98 Then
    yellow
    Else
    green
    End If
    
    'Col AJ (36): PArotid Rt Dmean<26Gy
    ActiveSheet.Cells(workRow, 36).Select
    compVal = 26 ' Gy
    If Selection.Value > compVal * 1.02 Then
    red
    ElseIf Selection.Value >= compVal * 0.98 Then
    yellow
    Else
    green
    End If
    
    'Col AK (37): PArotid Rt V30Gy[%]<50%
    ActiveSheet.Cells(workRow, 37).Select
    compVal = 50 ' %
    If Selection.Value > compVal * 1.02 Then
    red
    ElseIf Selection.Value >= compVal * 0.98 Then
    yellow
    Else
    green
    End If
    
    'Col AL (38): PArotid Rt V20Gy[cm3]<20cm3
    ActiveSheet.Cells(workRow, 38).Select
    compVal = 20 ' cm3
    If Selection.Value > compVal * 1.02 Then
    red
    ElseIf Selection.Value >= compVal * 0.98 Then
    yellow
    Else
    green
    End If
    
    'Col AM (39): PArotid Rt Dmean<60Gy
    ActiveSheet.Cells(workRow, 39).Select
    compVal = 60 ' Gy
    If Selection.Value > compVal * 1.02 Then
    red
    ElseIf Selection.Value >= compVal * 0.98 Then
    yellow
    Else
    green
    End If
    
    'Col AN (40): Subman Lt Dmean report only
    'ActiveSheet.Cells(workRow, 40).Select
    'compVal = 8 ' Gy
    'If Selection.Value > compVal * 1.02 Then
    'red
    'ElseIf Selection.Value >= compVal * 0.98 Then
    'yellow
    'Else
    'green
    'End If
    
    'Col AO (41): Subman Rt Dmean report only
    'ActiveSheet.Cells(workRow, 41).Select
    'compVal = 8 ' Gy
    'If Selection.Value > compVal * 1.02 Then
    'red
    'ElseIf Selection.Value >= compVal * 0.98 Then
    'yellow
    'Else
    'green
    'End If
    
    'Col AP (42): Mass Muscle Lt  Dmean<66Gy
    ActiveSheet.Cells(workRow, 42).Select
    compVal = 66 ' Gy
    If Selection.Value > compVal * 1.02 Then
    red
    ElseIf Selection.Value >= compVal * 0.98 Then
    yellow
    Else
    green
    End If
    
    'Col AQ (43): Mass Muscle Rt  Dmean<66Gy
    ActiveSheet.Cells(workRow, 43).Select
    compVal = 66 ' Gy
    If Selection.Value > compVal * 1.02 Then
    red
    ElseIf Selection.Value >= compVal * 0.98 Then
    yellow
    Else
    green
    End If
    
    'Col AR (44): Brachial Plexus Dmax<66Gy
    ActiveSheet.Cells(workRow, 44).Select
    compVal = 66 ' Gy
    If Selection.Value > compVal * 1.02 Then
    red
    ElseIf Selection.Value >= compVal * 0.98 Then
    yellow
    Else
    green
    End If
    
    ' PTVs
    'Col AT (46): PTVHigh D2%[Gy] not stablished
    'ActiveSheet.Cells(workRow, 46).Select
    'compVal = 66 ' Gy
    'If Selection.Value > compVal * 1.02 Then
    'red
    'ElseIf Selection.Value >= compVal * 0.98 Then
    'yellow
    'Else
    'green
    'End If
    
    'Col AU (47): PTVHigh V95%[Gy] not stablished
    'ActiveSheet.Cells(workRow, 47).Select
    'compVal = 66 ' Gy
    'If Selection.Value > compVal * 1.02 Then
    'red
    'ElseIf Selection.Value >= compVal * 0.98 Then
    'yellow
    'Else
    'green
    'End If
    
    'Col AV (48): PTVHigh Dmean[Gy] not stablished
    'ActiveSheet.Cells(workRow, 48).Select
    'compVal = 66 ' Gy
    'If Selection.Value > compVal * 1.02 Then
    'red
    'ElseIf Selection.Value >= compVal * 0.98 Then
    'yellow
    'Else
    'green
    'End If
    
    'Col AW (49): PTVInt V95%[Gy] not stablished
    'ActiveSheet.Cells(workRow, 49).Select
    'compVal = 66 ' Gy
    'If Selection.Value > compVal * 1.02 Then
    'red
    'ElseIf Selection.Value >= compVal * 0.98 Then
    'yellow
    'Else
    'green
    'End If
    
    'Col AX (50): PTVInt Dmean[Gy] not stablished
    'ActiveSheet.Cells(workRow, 50).Select
    'compVal = 66 ' Gy
    'If Selection.Value > compVal * 1.02 Then
    'red
    'ElseIf Selection.Value >= compVal * 0.98 Then
    'yellow
    'Else
    'green
    'End If
    
    'Col AY (51): PTVLow V95%[Gy] not stablished
    'ActiveSheet.Cells(workRow, 51).Select
    'compVal = 66 ' Gy
    'If Selection.Value > compVal * 1.02 Then
    'red
    'ElseIf Selection.Value >= compVal * 0.98 Then
    'yellow
    'Else
    'green
    'End If
    
    'Col AZ (52): PTVLow Dmean[Gy] not stablished
    'ActiveSheet.Cells(workRow, 52).Select
    'compVal = 66 ' Gy
    'If Selection.Value > compVal * 1.02 Then
    'red
    'ElseIf Selection.Value >= compVal * 0.98 Then
    'yellow
    'Else
    'green
    'End If

End Sub


Private Sub red()

' Red bold
    'ActiveCell.Select
    With Selection.Font
        .Color = -16776961
        .TintAndShade = 0
        .Bold = True
    End With

End Sub


Private Sub green()

'Green Bold
    ActiveCell.Select
    With Selection.Font
        .Color = -11489280
        .TintAndShade = 0
        .Bold = True
    End With

End Sub

Private Sub yellow()

'Yellow Bold
    ActiveCell.Select
    With Selection.Font
        .Color = -16213508
        .TintAndShade = 0
        .Bold = True
    End With

End Sub
