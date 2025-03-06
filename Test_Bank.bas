Attribute VB_Name = "Test_Bank"

Sub Create_Test_Bank()

'----------Variables----------
Dim x As Integer
Dim y As Integer
Dim z As Integer
Dim i As Integer
Dim h As Integer

Dim LastRow As Integer
Dim PrevLastRow As Integer
Dim EndRow As Integer

Dim Pin8 As Integer
Dim Pin6 As Integer
Dim Pin4 As Integer
Dim Pin3 As Integer
Dim Pin2 As Integer
Dim Pin1 As Integer

'----------Sheet Names----------

Dim A As String
Dim B As String
Dim C As String
Dim D As String
Dim E As String
Dim F As String

A = "Sheet1"
B = "Sheet2"
C = "Sheet3"
D = "Sheet4"
E = "Sheet5"
F = "Sheet6"

'----------Ranges----------

Dim TempRange As Range



'-------------------------Disable Excel Applications-------------------------
'Disable application auto-updating during code execution
Application.ScreenUpdating = False
Application.DisplayStatusBar = False
Application.EnableEvents = False

LastRow = Worksheets(A).Cells(Rows.Count, 66).End(xlUp).Row
If LastRow < 4 Then
    LastRow = 4
End If
Worksheets(A).Range("B4:BO" & LastRow).Cells.Clear
    
    Worksheets(A).Range("B4:BO64").Select
    With Selection.Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With


'----------------------Create test bank outline-----------------------

'-----Check assigned test bank locations----
LastRow = Worksheets(F).Cells(Rows.Count, 2).End(xlUp).Row

Dim BankID_1 As String
Dim BankID As String

For x = 65 To 69 '74 '90
    
    BankID_1 = Chr(x)
    
    Select Case BankID_1
        Case "A"
            y = 5
            z = 8
            i = 64
        Case "B"
            y = 11
            z = 14
            i = 128
        Case "C"
            y = 17
            z = 20
            i = 192
        Case "D"
            y = 23
            z = 26
            i = 256
    End Select

    BankID = "*" & BankID_1 & "*"
    
    If Not IsError(Application.Match(BankID, Worksheets(F).Range("B8:B" & LastRow), 0)) Then
        Worksheets(A).Range("C" & y & ":BN" & z).Select
        With Selection.Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
        End With
        With Selection.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
        End With
        With Selection.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
        End With
        With Selection.Borders(xlEdgeRight)
            .LineStyle = xlContinuous
        End With
        Worksheets(A).Range("C" & y & ":BN" & y).Select
        With Selection.Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
        End With
        With Selection.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
        End With
        With Selection.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
        End With
        With Selection.Borders(xlEdgeRight)
            .LineStyle = xlContinuous
        End With
        With Selection.Borders(xlInsideVertical)
            .LineStyle = xlContinuous
        End With
        With Selection.Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
        End With
            
        For j = 3 To 66
            Worksheets(A).Cells(y, j).Value = i
            i = i - 1
        Next j
        
            Range("C10:BN10").Select
        
        If BankID_1 <> "A" Then
            i = 64
            For j = 3 To 66
                Worksheets(A).Cells(y - 1, j).Value = i
                i = i - 1
            Next j
        End If
        
        Set TempRange = Worksheets(A).Range(Worksheets(A).Cells(y - 1, 3), Worksheets(A).Cells(y - 1, 66))
        TempRange.Select
        With Selection.Font
            .Name = "Arial"
            .Size = 8
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = -0.349986266670736
        End With
        With Selection
            .HorizontalAlignment = xlCenter
        End With
            
        Worksheets(A).Range("C" & y & ":BN" & y).Select
        With Selection
            .HorizontalAlignment = xlCenter
        End With
        Worksheets(A).Range("BN" & z).Select
        ActiveCell.FormulaR1C1 = "RED WIRE"
        Worksheets(A).Range("BN" & z).Select
        With Selection
            .HorizontalAlignment = xlRight
        End With
        With Selection.Font
            .Color = -16776961
            .TintAndShade = 0
            .Bold = True
        End With
        Worksheets(A).Range("C" & z).Select
        ActiveCell.FormulaR1C1 = "BANK " & BankID_1
        Worksheets(A).Range("C" & z).Select
        With Selection.Font
            .Bold = True
        End With
        Worksheets(A).Range("C" & y & ":BN" & y).Select
        With Selection.Font
            .Name = "Arial"
            .Size = 7.5
        End With
        
    End If
Next x

Dim CavityNumber As Integer
Dim SearchBankLoc As Range
Dim SearchPinout As Range
Dim RighPin As Integer
Dim RightPin_old As Integer
Dim LeftPin As Integer
Dim CompID As String
Dim TestID As String

For x = 65 To 69 '74 '90
    BankID_1 = Chr(x)
    
    Select Case BankID_1
        Case "A"
            y = 5
        Case "B"
            y = 11
        Case "C"
            y = 17
        Case "D"
            y = 23
    End Select
    
    RightPin = 66
    
    For w = 1 To 20
        RightPin_old = RightPin
        Set SearchBankLoc = Worksheets(F).Range("B8:B" & LastRow).Find(What:=BankID_1 & w)
        If Not SearchBankLoc Is Nothing Then
        
            CavityNumber = Worksheets(F).Cells(SearchBankLoc.Row, 7).Value
            CompID = Worksheets(F).Cells(SearchBankLoc.Row, 3).Value
            TestID = Worksheets(F).Cells(SearchBankLoc.Row, 8).Value
            
            Set SearchPinout = Worksheets(E).Range("B5:B154").Find(What:=CavityNumber)
            If Not SearchPinout Is Nothing Then
            
                Pin8 = Worksheets(E).Cells(SearchPinout.Row, 3).Value
                Pin6 = Worksheets(E).Cells(SearchPinout.Row, 4).Value
                Pin4 = Worksheets(E).Cells(SearchPinout.Row, 5).Value
                Pin3 = Worksheets(E).Cells(SearchPinout.Row, 6).Value
                Pin2 = Worksheets(E).Cells(SearchPinout.Row, 7).Value
                Pin1 = Worksheets(E).Cells(SearchPinout.Row, 8).Value
                
            End If
            
            For z = 1 To Pin8
                
                LeftPin = RightPin - 7
                Set TempRange = Worksheets(A).Range(Worksheets(A).Cells(y + 1, RightPin), Worksheets(A).Cells(y + 1, LeftPin))
                
                TempRange.Select
                With Selection
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlBottom
                    .Orientation = 90
                    .MergeCells = True
                End With
                With Selection.Borders(xlEdgeLeft)
                    .LineStyle = xlContinuous
                End With
                With Selection.Borders(xlEdgeTop)
                    .LineStyle = xlContinuous
                End With
                With Selection.Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous
                End With
                With Selection.Borders(xlEdgeRight)
                    .LineStyle = xlContinuous
                End With
                With Selection.Borders(xlInsideVertical)
                    .LineStyle = xlContinuous
                End With
                With Selection.Borders(xlInsideHorizontal)
                    .LineStyle = xlContinuous
                End With
                ActiveCell.FormulaR1C1 = "8 Pin"
                Rows((y + 1) & ":" & (y + 1)).RowHeight = 35
                
                RightPin = LeftPin - 1
            Next z
            
            For z = 1 To Pin6
                
                LeftPin = RightPin - 5
                Set TempRange = Worksheets(A).Range(Worksheets(A).Cells(y + 1, RightPin), Worksheets(A).Cells(y + 1, LeftPin))
                
                TempRange.Select
                With Selection
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlBottom
                    .Orientation = 90
                    .MergeCells = True
                End With
                With Selection.Borders(xlEdgeLeft)
                    .LineStyle = xlContinuous
                End With
                With Selection.Borders(xlEdgeTop)
                    .LineStyle = xlContinuous
                End With
                With Selection.Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous
                End With
                With Selection.Borders(xlEdgeRight)
                    .LineStyle = xlContinuous
                End With
                With Selection.Borders(xlInsideVertical)
                    .LineStyle = xlContinuous
                End With
                With Selection.Borders(xlInsideHorizontal)
                    .LineStyle = xlContinuous
                End With
                ActiveCell.FormulaR1C1 = "6 Pin"
                Rows((y + 1) & ":" & (y + 1)).RowHeight = 35
                
                RightPin = LeftPin - 1
            Next z
            
            For z = 1 To Pin4
                
                LeftPin = RightPin - 3
                Set TempRange = Worksheets(A).Range(Worksheets(A).Cells(y + 1, RightPin), Worksheets(A).Cells(y + 1, LeftPin))
                
                TempRange.Select
                With Selection
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlBottom
                    .Orientation = 90
                    .MergeCells = True
                End With
                With Selection.Borders(xlEdgeLeft)
                    .LineStyle = xlContinuous
                End With
                With Selection.Borders(xlEdgeTop)
                    .LineStyle = xlContinuous
                End With
                With Selection.Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous
                End With
                With Selection.Borders(xlEdgeRight)
                    .LineStyle = xlContinuous
                End With
                With Selection.Borders(xlInsideVertical)
                    .LineStyle = xlContinuous
                End With
                With Selection.Borders(xlInsideHorizontal)
                    .LineStyle = xlContinuous
                End With
                ActiveCell.FormulaR1C1 = "4 Pin"
                Rows((y + 1) & ":" & (y + 1)).RowHeight = 35
                
                RightPin = LeftPin - 1
            Next z
            
            For z = 1 To Pin3
                
                LeftPin = RightPin - 2
                Set TempRange = Worksheets(A).Range(Worksheets(A).Cells(y + 1, RightPin), Worksheets(A).Cells(y + 1, LeftPin))
                
                TempRange.Select
                With Selection
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlBottom
                    .Orientation = 90
                    .MergeCells = True
                End With
                With Selection.Borders(xlEdgeLeft)
                    .LineStyle = xlContinuous
                End With
                With Selection.Borders(xlEdgeTop)
                    .LineStyle = xlContinuous
                End With
                With Selection.Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous
                End With
                With Selection.Borders(xlEdgeRight)
                    .LineStyle = xlContinuous
                End With
                With Selection.Borders(xlInsideVertical)
                    .LineStyle = xlContinuous
                End With
                With Selection.Borders(xlInsideHorizontal)
                    .LineStyle = xlContinuous
                End With
                ActiveCell.FormulaR1C1 = "3 Pin"
                Rows((y + 1) & ":" & (y + 1)).RowHeight = 35
                
                RightPin = LeftPin - 1
            Next z
            
            For z = 1 To Pin2
                
                LeftPin = RightPin - 1
                Set TempRange = Worksheets(A).Range(Worksheets(A).Cells(y + 1, RightPin), Worksheets(A).Cells(y + 1, LeftPin))
                
                TempRange.Select
                With Selection
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlBottom
                    .Orientation = 90
                    .MergeCells = True
                End With
                With Selection.Borders(xlEdgeLeft)
                    .LineStyle = xlContinuous
                End With
                With Selection.Borders(xlEdgeTop)
                    .LineStyle = xlContinuous
                End With
                With Selection.Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous
                End With
                With Selection.Borders(xlEdgeRight)
                    .LineStyle = xlContinuous
                End With
                With Selection.Borders(xlInsideVertical)
                    .LineStyle = xlContinuous
                End With
                With Selection.Borders(xlInsideHorizontal)
                    .LineStyle = xlContinuous
                End With
                ActiveCell.FormulaR1C1 = "2 Pin"
                Rows((y + 1) & ":" & (y + 1)).RowHeight = 35
                
                RightPin = LeftPin - 1
            Next z
                
            For z = 1 To Pin1
                
                LeftPin = RightPin
                Set TempRange = Worksheets(A).Range(Worksheets(A).Cells(y + 1, RightPin), Worksheets(A).Cells(y + 1, LeftPin))
                
                TempRange.Select
                With Selection
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlBottom
                    .Orientation = 90
                    .MergeCells = True
                End With
                With Selection.Borders(xlEdgeLeft)
                    .LineStyle = xlContinuous
                End With
                With Selection.Borders(xlEdgeTop)
                    .LineStyle = xlContinuous
                End With
                With Selection.Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous
                End With
                With Selection.Borders(xlEdgeRight)
                    .LineStyle = xlContinuous
                End With
                With Selection.Borders(xlInsideVertical)
                    .LineStyle = xlContinuous
                End With
                With Selection.Borders(xlInsideHorizontal)
                    .LineStyle = xlContinuous
                End With
                ActiveCell.FormulaR1C1 = "1 Pin"
                Rows((y + 1) & ":" & (y + 1)).RowHeight = 35
                
                RightPin = LeftPin - 1
            Next z
            
            Set TempRange = Worksheets(A).Range(Worksheets(A).Cells(y + 2, RightPin_old), Worksheets(A).Cells(y + 2, LeftPin))
            TempRange.Select
            With Selection
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlBottom
                .Orientation = 90
                .MergeCells = True
            End With
            With Selection.Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
            End With
            With Selection.Borders(xlEdgeTop)
                .LineStyle = xlContinuous
            End With
            With Selection.Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
            End With
            With Selection.Borders(xlEdgeRight)
                .LineStyle = xlContinuous
            End With
            With Selection.Borders(xlInsideVertical)
                .LineStyle = xlContinuous
            End With
            With Selection.Borders(xlInsideHorizontal)
                .LineStyle = xlContinuous
            End With
            With Selection.Font
            .Name = "Arial"
            .Size = 8
            End With
            ActiveCell.FormulaR1C1 = CompID & " [" & TestID & "]"
            Rows((y + 2) & ":" & (y + 2)).RowHeight = 75
            
        End If
        
    Next w
    
Next x

'-------------------------Set Print View-------------------------
LastRow = Worksheets(A).Cells(Rows.Count, 66).End(xlUp).Row
Worksheets(A).PageSetup.PrintArea = "B1:BO" & LastRow + 6


'-------------------------Enable Excel Applications-------------------------
'Disable application auto-updating during code execution
Application.ScreenUpdating = True
Application.DisplayStatusBar = True
Application.EnableEvents = True

End Sub

Sub Clear_Test_Bank()


'----------Variables----------
Dim LastRow As Integer

'----------Sheet Names----------

Dim A As String

A = "Sheet1"

LastRow = Worksheets(A).Cells(Rows.Count, 66).End(xlUp).Row
If LastRow < 4 Then
    LastRow = 4
End If
Worksheets(A).Range("B4:BO" & LastRow).Cells.Clear


    Worksheets(A).Range("B4:BO64").Select
    With Selection.Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With


End Sub

