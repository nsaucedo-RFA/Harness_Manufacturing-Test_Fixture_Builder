Attribute VB_Name = "Test_Bank_2"
Sub Create_Test_Bank_2()

'----------Variables----------
Dim x As Integer
Dim y As Integer
Dim z As Integer
Dim v As Integer
Dim i As Integer
Dim h As Integer
Dim u As Integer

Dim LastRow As Integer

Dim Pin8 As Integer
Dim Pin6 As Integer
Dim Pin4 As Integer
Dim Pin3 As Integer
Dim Pin2 As Integer
Dim Pin1 As Integer

Dim BankID_1 As String
Dim BankID As String

Dim CavityNumber As Integer
Dim SearchBankLoc As Range
Dim SearchPinout As Range
Dim RighPin As Integer
Dim RightPin_old As Integer
Dim LeftPin As Integer
Dim CompID As String
Dim TestID As String


'----------Sheet Names----------

Dim A As String
Dim B As String
Dim C As String
Dim D As String
Dim E As String
Dim F As String

A = "Bank Layout"
B = "New Adapter Build"
C = "New Adapter BOM"
D = "Sheet4"
E = "Sheet5"
F = "Component List"

'----------Ranges----------

Dim TempRange As Range

'-------------------------Define Arrays ------------------------------------

'Define BankLocations in Array
Dim ArrayID() As String
ReDim ArrayID(0 To 0) As String

'Define CompID in Array
Dim CompID2() As String
ReDim CompID2(0 To 0) As String

'Define CavityNumber in Array
Dim CavID2() As String
ReDim CavID2(0 To 0) As String

'Define TEstID in Array
Dim TestID2() As String
ReDim TestID2(0 To 0) As String

'-------------------------Disable Excel Applications-------------------------
'Disable application auto-updating during code execution
Application.ScreenUpdating = False
Application.DisplayStatusBar = False
Application.EnableEvents = False

'----------------------Clear Test Bank Layout----------------------

LastRow = Worksheets(A).Cells(Rows.Count, 66).End(xlUp).Row
If LastRow < 5 Then
    LastRow = 5
End If
Worksheets(A).Range("B5:BO" & LastRow).Cells.Clear
    
    Worksheets(A).Range("B5:BO66").Select
    With Selection.Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With

'----------------------Create Test Banks-----------------------

'-----Check for assigned test bank locations-----

LastRow = Worksheets(F).Cells(Rows.Count, 2).End(xlUp).Row
For x = 7 To LastRow

    If Worksheets(F).Cells(x, 2).Value = "" Then
        GoTo Skip_Location
    End If
    
    If Worksheets(F).Cells(x, 8).Value = "" Then
        GoTo Skip_Location
    End If
    
    If IsNumeric(Worksheets(F).Cells(x, 8).Value) = False Then
        MsgBox "Invalid number for bank location " & Worksheets(F).Cells(x, 2).Value & ".", vbOKOnly + vbExclamation
        GoTo Skip_Location
    End If
    
    ReDim Preserve ArrayID(0 To UBound(ArrayID) + 1) As String
    ArrayID(UBound(ArrayID)) = Worksheets(F).Cells(x, 2).Value
    
    ReDim Preserve CompID2(0 To UBound(CompID2) + 1) As String
    CompID2(UBound(CompID2)) = Worksheets(F).Cells(x, 3).Value
    
    ReDim Preserve CavID2(0 To UBound(CavID2) + 1) As String
    CavID2(UBound(CavID2)) = Worksheets(F).Cells(x, 8).Value
    
    ReDim Preserve TestID2(0 To UBound(TestID2) + 1) As String
    TestID2(UBound(TestID2)) = Worksheets(F).Cells(x, 9).Value
    
Skip_Location:
Next x

For x = 65 To 72
    
    BankID_1 = Chr(x)
    
    Select Case BankID_1
        Case "A"
            y = 6
            z = 9
            i = 64
        Case "B"
            y = 12
            z = 15
            i = 128
        Case "C"
            y = 18
            z = 21
            i = 192
        Case "D"
            y = 24
            z = 27
            i = 256
        Case "E"
            y = 30
            z = 33
            i = 320
        Case "F"
            y = 36
            z = 39
            i = 384
        Case "G"
            y = 42
            z = 45
            i = 448
        Case "H"
            y = 48
            z = 51
            i = 512
        Case "I"
            y = 48
            z = 57
            i = 576
        Case "J"
            y = 48
            z = 63
            i = 640
    End Select

    BankID = "*" & BankID_1 & "*"
    
    If Not IsError(Application.Match(BankID, ArrayID(), 0)) Then
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

'-----Loop through assigned locations-----

For x = 65 To 72 '90
    BankID_1 = Chr(x)
    
    Select Case BankID_1
        Case "A"
            y = 6
        Case "B"
            y = 12
        Case "C"
            y = 18
        Case "D"
            y = 24
        Case "E"
            y = 30
        Case "F"
            y = 36
        Case "G"
            y = 42
        Case "H"
            y = 48
        Case "I"
            y = 54
        Case "J"
            y = 60
    End Select
    
    RightPin = 66
    
    For w = 1 To 64
        RightPin_old = RightPin
        
        For v = 1 To UBound(ArrayID)
            
            If ArrayID(v) = BankID_1 & w Then
            
                CavityNumber = CavID2(v)
                CompID = CompID2(v)
                TestID = TestID2(v)
                
                If Left(TestID, 2) = "8P" Then

                    If CavityNumber <> 8 Then
                    
                        LeftPin = RightPin - (CavityNumber - 1)
                        
                        Set TempRange = Worksheets(A).Range(Worksheets(A).Cells(y + 1, RightPin), Worksheets(A).Cells(y + 1, LeftPin))
                
                        TempRange.Select
                        With Selection
                            .HorizontalAlignment = xlCenter
                            .VerticalAlignment = xlBottom
                            .Orientation = 90
                            .MergeCells = True
                        End With
                                         
                        With Selection.Font
                            .Name = "Arial"
                            .Size = 9
                        End With
                        
                        TempRange.Borders(xlEdgeLeft).LineStyle = xlContinuous
                        TempRange.Borders(xlEdgeBottom).LineStyle = xlContinuous
                        TempRange.Borders(xlEdgeRight).LineStyle = xlContinuous
                        TempRange.Borders(xlInsideVertical).LineStyle = xlContinuous
                        TempRange.Borders(xlInsideHorizontal).LineStyle = xlContinuous
                        TempRange.Borders(xlEdgeTop).LineStyle = xlContinuous
                        
                        ActiveCell.FormulaR1C1 = "8 Pin SL*"
                        Rows((y + 1) & ":" & (y + 1)).RowHeight = 45
                        
                        RightPin = LeftPin - 1
                        
                        Set TempRange = Worksheets(A).Range(Worksheets(A).Cells(y + 2, RightPin_old), Worksheets(A).Cells(y + 2, LeftPin))
                        TempRange.Select
                        With Selection
                            .HorizontalAlignment = xlCenter
                            .VerticalAlignment = xlBottom
                            .Orientation = 90
                            .MergeCells = True
                        End With
                   
                        With Selection.Font
                            .Name = "Arial"
                            .Size = 9
                        End With
                   
                        TempRange.Borders(xlEdgeLeft).LineStyle = xlContinuous
                        TempRange.Borders(xlEdgeBottom).LineStyle = xlContinuous
                        TempRange.Borders(xlEdgeRight).LineStyle = xlContinuous
                        TempRange.Borders(xlInsideVertical).LineStyle = xlContinuous
                        TempRange.Borders(xlInsideHorizontal).LineStyle = xlContinuous
                        TempRange.Borders(xlEdgeTop).LineStyle = xlContinuous
                        
                        ActiveCell.FormulaR1C1 = CompID & " [" & TestID & "]"
                        Rows((y + 2) & ":" & (y + 2)).RowHeight = 75
                                            
                        GoTo NextSlot
                    End If
                End If
                     
                For u = 5 To 154
                    If CavityNumber = Worksheets(E).Cells(u, 2).Value Then
                        Pin8 = Worksheets(E).Cells(u, 3).Value
                        Pin6 = Worksheets(E).Cells(u, 4).Value
                        Pin4 = Worksheets(E).Cells(u, 5).Value
                        Pin3 = Worksheets(E).Cells(u, 6).Value
                        Pin2 = Worksheets(E).Cells(u, 7).Value
                        Pin1 = Worksheets(E).Cells(u, 8).Value
                    End If
                Next u
                
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
                    
                    With Selection.Font
                        .Name = "Arial"
                        .Size = 9
                    End With

                    TempRange.Borders(xlEdgeLeft).LineStyle = xlContinuous
                    TempRange.Borders(xlEdgeBottom).LineStyle = xlContinuous
                    TempRange.Borders(xlEdgeRight).LineStyle = xlContinuous
                    TempRange.Borders(xlInsideVertical).LineStyle = xlContinuous
                    TempRange.Borders(xlInsideHorizontal).LineStyle = xlContinuous
                    TempRange.Borders(xlEdgeTop).LineStyle = xlContinuous

                    ActiveCell.FormulaR1C1 = "8 Pin"
                    Rows((y + 1) & ":" & (y + 1)).RowHeight = 45
                    
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
                    
                    With Selection.Font
                        .Name = "Arial"
                        .Size = 9
                    End With
                    
                    TempRange.Borders(xlEdgeLeft).LineStyle = xlContinuous
                    TempRange.Borders(xlEdgeBottom).LineStyle = xlContinuous
                    TempRange.Borders(xlEdgeRight).LineStyle = xlContinuous
                    TempRange.Borders(xlInsideVertical).LineStyle = xlContinuous
                    TempRange.Borders(xlInsideHorizontal).LineStyle = xlContinuous
                    TempRange.Borders(xlEdgeTop).LineStyle = xlContinuous
        
                    ActiveCell.FormulaR1C1 = "6 Pin"
                    Rows((y + 1) & ":" & (y + 1)).RowHeight = 45
                    
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
                    
                    With Selection.Font
                        .Name = "Arial"
                        .Size = 9
                    End With
                    
                    TempRange.Borders(xlEdgeLeft).LineStyle = xlContinuous
                    TempRange.Borders(xlEdgeBottom).LineStyle = xlContinuous
                    TempRange.Borders(xlEdgeRight).LineStyle = xlContinuous
                    TempRange.Borders(xlInsideVertical).LineStyle = xlContinuous
                    TempRange.Borders(xlInsideHorizontal).LineStyle = xlContinuous
                    TempRange.Borders(xlEdgeTop).LineStyle = xlContinuous
                    
                    ActiveCell.FormulaR1C1 = "4 Pin"
                    Rows((y + 1) & ":" & (y + 1)).RowHeight = 45
                    
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
                    
                    With Selection.Font
                        .Name = "Arial"
                        .Size = 9
                    End With
                    
                    TempRange.Borders(xlEdgeLeft).LineStyle = xlContinuous
                    TempRange.Borders(xlEdgeBottom).LineStyle = xlContinuous
                    TempRange.Borders(xlEdgeRight).LineStyle = xlContinuous
                    TempRange.Borders(xlInsideVertical).LineStyle = xlContinuous
                    TempRange.Borders(xlInsideHorizontal).LineStyle = xlContinuous
                    TempRange.Borders(xlEdgeTop).LineStyle = xlContinuous
                    
                    ActiveCell.FormulaR1C1 = "3 Pin"
                    Rows((y + 1) & ":" & (y + 1)).RowHeight = 45
                    
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
                    
                    With Selection.Font
                        .Name = "Arial"
                        .Size = 9
                    End With
                    
                    TempRange.Borders(xlEdgeLeft).LineStyle = xlContinuous
                    TempRange.Borders(xlEdgeBottom).LineStyle = xlContinuous
                    TempRange.Borders(xlEdgeRight).LineStyle = xlContinuous
                    TempRange.Borders(xlInsideVertical).LineStyle = xlContinuous
                    TempRange.Borders(xlInsideHorizontal).LineStyle = xlContinuous
                    TempRange.Borders(xlEdgeTop).LineStyle = xlContinuous
                                     
                    ActiveCell.FormulaR1C1 = "2 Pin"
                    Rows((y + 1) & ":" & (y + 1)).RowHeight = 45
                    
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
                    
                    With Selection.Font
                        .Name = "Arial"
                        .Size = 9
                    End With
                    
                    TempRange.Borders(xlEdgeLeft).LineStyle = xlContinuous
                    TempRange.Borders(xlEdgeBottom).LineStyle = xlContinuous
                    TempRange.Borders(xlEdgeRight).LineStyle = xlContinuous
                    TempRange.Borders(xlInsideVertical).LineStyle = xlContinuous
                    TempRange.Borders(xlInsideHorizontal).LineStyle = xlContinuous
                    TempRange.Borders(xlEdgeTop).LineStyle = xlContinuous
                    
                    'ActiveCell.FormulaR1C1 = ""
                    Rows((y + 1) & ":" & (y + 1)).RowHeight = 45
                    
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
                
                With Selection.Font
                    .Name = "Arial"
                    .Size = 9
                End With
                
                TempRange.Borders(xlEdgeLeft).LineStyle = xlContinuous
                TempRange.Borders(xlEdgeBottom).LineStyle = xlContinuous
                TempRange.Borders(xlEdgeRight).LineStyle = xlContinuous
                TempRange.Borders(xlInsideVertical).LineStyle = xlContinuous
                TempRange.Borders(xlInsideHorizontal).LineStyle = xlContinuous
                TempRange.Borders(xlEdgeTop).LineStyle = xlContinuous

                ActiveCell.FormulaR1C1 = CompID & " [" & TestID & "]"
                Rows((y + 2) & ":" & (y + 2)).RowHeight = 75
            
            End If
                
        Next v
NextSlot:
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

Sub Clear_Test_Bank_2()


'----------Variables----------
Dim LastRow As Integer

'----------Sheet Names----------

Dim A As String

A = "Bank Layout"

LastRow = Worksheets(A).Cells(Rows.Count, 66).End(xlUp).Row
If LastRow < 5 Then
    LastRow = 5
End If
Worksheets(A).Range("B5:BO" & LastRow).Cells.Clear


    Worksheets(A).Range("B5:BO66").Select
    With Selection.Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With


End Sub


