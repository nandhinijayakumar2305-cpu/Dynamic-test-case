Attribute VB_Name = "Module1"
Option Explicit

'==========================================================
' GenerateTestValues - Auto generates Positive & Negative
' test values based on Min/Max constraints in each row
'==========================================================
Sub GenerateTestValues()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim minLen As Variant, maxLen As Variant
    Dim minVal As Variant, maxVal As Variant
    Dim defVal As Variant
    Dim p1 As Variant, p2 As Variant, p3 As Variant, n1 As Variant
    
    Application.ScreenUpdating = False
    
    Set ws = ThisWorkbook.Sheets("IT 001")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    Dim greenFill As Long, redFill As Long
    greenFill = RGB(198, 239, 206)
    redFill = RGB(255, 199, 206)
    
    Dim processed As Long
    processed = 0
    
    For i = 3 To lastRow
        ' Read constraints
        minLen = ws.Cells(i, 11).Value  ' Col K - Min Field Length
        maxLen = ws.Cells(i, 12).Value  ' Col L - Max Field Length
        minVal = ws.Cells(i, 13).Value  ' Col M - Min Value
        maxVal = ws.Cells(i, 14).Value  ' Col N - Max Value
        defVal = ws.Cells(i, 26).Value  ' Col Z - Default Value
        
        ' Skip rows with no constraints
        If IsEmpty(minLen) And IsEmpty(maxLen) And IsEmpty(minVal) And IsEmpty(maxVal) Then
            ws.Cells(i, 30).Value = "N.A"
            ws.Cells(i, 31).Value = "N.A"
            ws.Cells(i, 32).Value = "N.A"
            ws.Cells(i, 33).Value = "N.A"
        Else
            processed = processed + 1
            
            '--- POSITIVE 1: Min valid value (lower boundary) ---
            If Not IsEmpty(minVal) And CStr(minVal) <> "N.A" Then
                p1 = minVal
            ElseIf Not IsEmpty(minLen) And CStr(minLen) <> "N.A" Then
                p1 = 0
            Else
                p1 = "N.A"
            End If
            
            '--- POSITIVE 2: Max valid value (upper boundary) ---
            If Not IsEmpty(maxLen) And CStr(maxLen) <> "N.A" Then
                p2 = maxLen
            ElseIf Not IsEmpty(maxVal) And CStr(maxVal) <> "N.A" Then
                p2 = maxVal
            Else
                p2 = "N.A"
            End If
            
            '--- POSITIVE 3: Default value (if exists) ---
            If Not IsEmpty(defVal) And CStr(defVal) <> "" Then
                p3 = defVal
            Else
                p3 = "N.A"
            End If
            
            '--- NEGATIVE 1: Value exceeding max (invalid boundary) ---
            If Not IsEmpty(maxVal) And CStr(maxVal) <> "N.A" Then
                If IsNumeric(maxVal) Then
                    n1 = CDbl(maxVal) + 1
                Else
                    n1 = "INVALID_" & CStr(maxVal)
                End If
            ElseIf Not IsEmpty(maxLen) And CStr(maxLen) <> "N.A" And IsNumeric(maxLen) Then
                n1 = CLng(maxLen) + 1
            Else
                n1 = -1
            End If
            
            ' Write Positive values (green)
            With ws.Cells(i, 30)
                .Value = p1
                .Interior.Color = greenFill
                .Font.Color = RGB(55, 86, 35)
            End With
            With ws.Cells(i, 31)
                .Value = p2
                .Interior.Color = greenFill
                .Font.Color = RGB(55, 86, 35)
            End With
            With ws.Cells(i, 32)
                .Value = p3
                .Interior.Color = greenFill
                .Font.Color = RGB(55, 86, 35)
            End With
            
            ' Write Negative value (red)
            With ws.Cells(i, 33)
                .Value = n1
                .Interior.Color = redFill
                .Font.Color = RGB(156, 0, 6)
            End With
        End If
    Next i
    
    Application.ScreenUpdating = True
    
    MsgBox "Test values generated!" & Chr(10) & Chr(10) & _
           "Rows processed: " & processed & Chr(10) & _
           "Green cells = Positive (VALID) values" & Chr(10) & _
           "Red cells = Negative (INVALID) values", _
           vbInformation, "Generate Test Values - Done"
End Sub

'==========================================================
' ClearTestValues - Clears all generated test values
'==========================================================
Sub ClearTestValues()
    Dim ws As Worksheet
    Dim lastRow As Long
    
    If MsgBox("Clear all generated test values?", vbYesNo + vbQuestion, "Confirm") = vbNo Then Exit Sub
    
    Set ws = ThisWorkbook.Sheets("IT 001")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    With ws.Range("AD3:AG" & lastRow)
        .ClearContents
        .Interior.ColorIndex = xlNone
        .Font.ColorIndex = xlAutomatic
    End With
    
    MsgBox "Cleared!", vbInformation, "Done"
End Sub
