Attribute VB_Name = "Adapter_Build"
Sub Build_Table()
'New button

'----------Variables----------
Dim x As Integer
Dim y As Integer
Dim z As Integer
Dim i As Integer
Dim LastPin As Integer
Dim A As String
Dim B As String
Dim C As String
Dim D As String
Dim E As String
Dim myvalue As Variant
Dim NewNumberINT As Integer
Dim CheckNumberINT As Integer
Dim lrow As Integer
Dim Pin_Start As Integer

Dim Pin8 As Integer
Dim Pin6 As Integer
Dim Pin4 As Integer
Dim Pin3 As Integer
Dim Pin2 As Integer
Dim Pin1 As Integer
Dim AdapterRow As Integer

'----------Sheet Names----------
A = "Bank Layout"
B = "New Adapter Build"
C = "New Adapter BOM"
D = "Sheet4"
E = "Sheet5"
F = "Component List"

'-------------------------Enable Excel Applications-------------------------
'Disable application auto-updating during code execution
Application.ScreenUpdating = False
Application.DisplayStatusBar = False
Application.EnableEvents = False

'----------Input Box-----------
myvalue = InputBox("Please enter number of cavities for new test adapter")

'Check numeric

If IsNumeric(myvalue) = False Then
    MsgBox "Please enter a number.", vbOKOnly + vbCritical, "Error"
    Exit Sub
End If

'Check Blank

If myvalue = "" Then
    MsgBox "Please enter a number.", vbOKOnly + vbCritical, "Error"
    Exit Sub
End If


'----------Locate next open row-----------

lrow = Worksheets(B).Cells(Rows.Count, 3).End(xlUp).Row
lrow = lrow + 2
Worksheets(D).Range("C5:L9").Copy Worksheets(B).Range("C" & lrow & ":L" & lrow + 5)

'----------Select Tables-----------

For x = 5 To 154
    Pin_Start = Worksheets(B).Cells(Rows.Count, 3).End(xlUp).Row
    NewNumberINT = myvalue
    CheckNumberINT = Worksheets(E).Cells(x, 2).Value
    
    If CheckNumberINT = NewNumberINT Then
    
        Pin8 = Worksheets(E).Cells(x, 3).Value
        Pin6 = Worksheets(E).Cells(x, 4).Value
        Pin4 = Worksheets(E).Cells(x, 5).Value
        Pin3 = Worksheets(E).Cells(x, 6).Value
        Pin2 = Worksheets(E).Cells(x, 7).Value
        Pin1 = Worksheets(E).Cells(x, 8).Value
    
        For y = 1 To Pin8
        
            lrow = Worksheets(B).Cells(Rows.Count, 3).End(xlUp).Row
            lrow = lrow + 1
            Worksheets(D).Range("C13:L20").Copy Worksheets(B).Range("C" & lrow & ":L" & lrow + 7)
        
        Next y
        
        For y = 1 To Pin6
        
            lrow = Worksheets(B).Cells(Rows.Count, 3).End(xlUp).Row
            lrow = lrow + 1
            Worksheets(D).Range("C22:L27").Copy Worksheets(B).Range("C" & lrow & ":L" & lrow + 5)
                
        Next y

        For y = 1 To Pin4
        
            lrow = Worksheets(B).Cells(Rows.Count, 3).End(xlUp).Row
            lrow = lrow + 1
            Worksheets(D).Range("C29:L32").Copy Worksheets(B).Range("C" & lrow & ":L" & lrow + 3)
                
        Next y

        For y = 1 To Pin3
        
            lrow = Worksheets(B).Cells(Rows.Count, 3).End(xlUp).Row
            lrow = lrow + 1
            Worksheets(D).Range("C34:L36").Copy Worksheets(B).Range("C" & lrow & ":L" & lrow + 2)
                
        Next y

        For y = 1 To Pin2
        
            lrow = Worksheets(B).Cells(Rows.Count, 3).End(xlUp).Row
            lrow = lrow + 1
            Worksheets(D).Range("C38:L39").Copy Worksheets(B).Range("C" & lrow & ":L" & lrow + 1)
                
        Next y

        For y = 1 To Pin1
        
            lrow = Worksheets(B).Cells(Rows.Count, 3).End(xlUp).Row
            lrow = lrow + 1
            Worksheets(D).Range("C41:L41").Copy Worksheets(B).Range("C" & lrow & ":L" & lrow)
                
        Next y
        
        '---------- Adjust Pincount----------
        
        lrow = Worksheets(B).Cells(Rows.Count, 3).End(xlUp).Row
        i = 1
        For y = (Pin_Start + 1) To lrow
            Worksheets(B).Cells(y, 3).Value = Worksheets(B).Cells(y, 3).Value & " (" & i & ")"
            i = i + 1
        Next y
        
    End If
Next x

'----------Adjust Print Area----------

LastRow = Worksheets(B).Cells(Rows.Count, 3).End(xlUp).Row
Worksheets(B).PageSetup.PrintArea = "B2:M" & LastRow + 6

'-------------------------Enable Excel Applications-------------------------
'Disable application auto-updating during code execution
Application.ScreenUpdating = True
Application.DisplayStatusBar = True
Application.EnableEvents = True


End Sub

Sub Clear()
'Clear button

'Assign workbook tab to letter
Dim A As String
A = "New Adapter Build"

'Message prompt before clearing sheet contents
If MsgBox("Are you sure you want to clear the sheet?", vbYesNo + vbExclamation, "Clear") = vbNo Then
    Exit Sub
Else
    Worksheets(A).Range("B5:M1000").Cells.Clear
End If

End Sub
