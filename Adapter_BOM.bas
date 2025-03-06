Attribute VB_Name = "Adapter_BOM"
Sub Create_BOMTable()
'Create button

'----------Variables----------
Dim x As Integer
Dim y As Integer
Dim z As Integer
Dim i As Integer
Dim h As Integer

Dim TermPN As String

Dim LastRow As Integer
Dim PrevLastRow As Integer
Dim EndRow As Integer

Dim ArrayLoc As Integer

Dim A As String
Dim B As String
Dim C As String
Dim D As String
Dim E As String
Dim F As String

Dim Duplicate As Boolean

'Define Connector Array
Dim Conn() As String
ReDim Conn(0 To 0) As String

'Define ConnectorQty Array
Dim ConnQty() As Integer
ReDim ConnQty(0 To 0) As Integer

'Define Accessories Array
Dim Accy() As String
ReDim Accy(0 To 0) As String

'Define AccessoryQty Array
Dim AccyQty() As Integer
ReDim AccyQty(0 To 0) As Integer

'Define Terminal Array
Dim Term() As String
ReDim Term(0 To 0) As String

'Define TerminalQty Array
Dim TermQty() As Integer
ReDim TermQty(0 To 0) As Integer

'Define Seal Array
Dim Seal() As String
ReDim Seal(0 To 0) As String

'Define Seal Array
Dim SealQty() As Integer
ReDim SealQty(0 To 0) As Integer

'Define Wire Array
Dim Wire() As String
ReDim Wire(0 To 0) As String

'Define WireQty Array
Dim WireQty() As Integer
ReDim WireQty(0 To 0) As Integer

'Define Amphenol Connector Array
Dim AmpConn() As String
ReDim AmpConn(0 To 0) As String

'Define Amphenol Connector Array
Dim AmpConnQty() As Integer
ReDim AmpConnQty(0 To 0) As Integer

'Define Amphenol Connector Wedge Array
Dim AmpWedg() As String
ReDim AmpWedg(0 To 0) As String

'Define Amphenol Connector Wedge Array
Dim AmpWedgQty() As Integer
ReDim AmpWedgQty(0 To 0) As Integer

'Define Amphenol Term Array
Dim AmpTerm() As String
ReDim AmpTerm(0 To 0) As String

'Define Amphenol Term Qty Array
Dim AmpTermQty() As Integer
ReDim AmpTermQty(0 To 0) As Integer

'----------Sheet Names----------
A = "Bank Layout"
B = "New Adapter Build"
C = "New Adapter BOM"
D = "Sheet4"
E = "Sheet5"
F = "Component List"

'-------------------------Disable Excel Applications-------------------------
'Disable application auto-updating during code execution
Application.ScreenUpdating = False
Application.DisplayStatusBar = False
Application.EnableEvents = False

Worksheets(C).Range("B5:G1000").Cells.Clear

'-------------------------Create BOM table header-------------------------

LastRow = 6

'Table
Worksheets(C).Cells(LastRow, 3).Value = "BOM"
Worksheets(C).Range("C" & LastRow & ":F" & LastRow).Merge
Worksheets(C).Range("C" & LastRow & ":F" & LastRow).Interior.ThemeColor = xlThemeColorLight1
Worksheets(C).Range("C" & LastRow & ":F" & LastRow).Font.ThemeColor = xlThemeColorDark1
Worksheets(C).Range("C" & LastRow & ":F" & LastRow).Font.Bold = True
    
LastRow = 7

'Table
Worksheets(C).Cells(LastRow, 3).Value = "Index"
Worksheets(C).Cells(LastRow, 4).Value = "Part Number"
Worksheets(C).Cells(LastRow, 5).Value = "Quantity"
Worksheets(C).Cells(LastRow, 6).Value = "Location"

Worksheets(C).Range("C" & LastRow & ":F" & LastRow).Interior.ThemeColor = xlThemeColorDark1
Worksheets(C).Range("C" & LastRow & ":F" & LastRow).Interior.Color = RGB(225, 225, 225)
Worksheets(C).Range("C" & LastRow & ":F" & LastRow).Font.Bold = True

'-------------------------Index Adapter Part Numbers-------------------------

'Look for adpter part # in column 3
'store & check quantity

LastRow = Worksheets(B).Cells(Rows.Count, 3).End(xlUp).Row

For x = 6 To LastRow
    
    If InStr(UCase(Worksheets(B).Cells(x, 3).Value), "QTY") > 0 Then
        
        Qty = Worksheets(B).Cells(x, 4).Value
        
        If Qty = "" Then
            Qty = 1
        End If
    
        '---------Find end of table row----------
    
        EndRow = Worksheets(B).Range("C" & x & ":C" & LastRow).End(xlDown).Offset(1).Row
    
        EndRow = EndRow - 1 'Table End Row
    
        '----------Add Conn Array----------
    
        If Worksheets(B).Cells(x + 1, 4).Value = "" Then
            GoTo AddAccy
        Else
                
            If UBound(Conn) = 0 Then
                    
                ReDim Preserve Conn(0 To UBound(Conn) + 1) As String
                Conn(UBound(Conn)) = Worksheets(B).Cells(x + 1, 4).Value
                    
                ReDim Preserve ConnQty(0 To UBound(ConnQty) + 1) As Integer
                ConnQty(UBound(ConnQty)) = Qty
                    
            Else
                
                Duplicate = False
                For y = 1 To UBound(Conn)
                    If Worksheets(B).Cells(x + 1, 4).Value = Conn(y) Then
                        Duplicate = True
                        ArrayLoc = y
                        GoTo DuplicateConn
                    End If
                Next y
                    
DuplicateConn:
    
                If Duplicate = True Then
                    ConnQty(ArrayLoc) = ConnQty(ArrayLoc) + Qty
                Else
                    ReDim Preserve Conn(0 To UBound(Conn) + 1) As String
                    Conn(UBound(Conn)) = Worksheets(B).Cells(x + 1, 4).Value
                        
                    ReDim Preserve ConnQty(0 To UBound(ConnQty) + 1) As Integer
                    ConnQty(UBound(ConnQty)) = Qty
                End If
            
            End If
            
        End If

AddAccy: '----------Add Accy----------
   
        If Worksheets(B).Cells(x + 2, 4).Value = "" Then
            GoTo AddTerm
        Else
                
            If UBound(Accy) = 0 Then
                        
                ReDim Preserve Accy(0 To UBound(Accy) + 1) As String
                Accy(UBound(Accy)) = Worksheets(B).Cells(x + 2, 4).Value
                            
                ReDim Preserve AccyQty(0 To UBound(AccyQty) + 1) As Integer
                AccyQty(UBound(AccyQty)) = Qty
                            
            Else
                        
                Duplicate = False
                For y = 1 To UBound(Accy)
                    If Worksheets(B).Cells(x + 2, 4).Value = Accy(y) Then
                        Duplicate = True
                        ArrayLoc = y
                        GoTo DuplicateAccy
                    End If
                Next y
                
DuplicateAccy:
        
                If Duplicate = True Then
                    AccyQty(ArrayLoc) = AccyQty(ArrayLoc) + Qty
                Else
                    ReDim Preserve Accy(0 To UBound(Accy) + 1) As String
                    Accy(UBound(Accy)) = Worksheets(B).Cells(x + 2, 4).Value
                                    
                    ReDim Preserve AccyQty(0 To UBound(AccyQty) + 1) As Integer
                    AccyQty(UBound(AccyQty)) = Qty
                End If
                        
            End If
                        
        End If

AddTerm: '----------Add Term----------
        
        For z = (x + 4) To EndRow
        
            If Worksheets(B).Cells(z, 4).Value = "" Then
                GoTo Nextz_Term
            Else
            
                If UBound(Term) = 0 Then
                            
                    ReDim Preserve Term(0 To UBound(Term) + 1) As String
                    Term(UBound(Term)) = Worksheets(B).Cells(z, 4).Value
                                
                    ReDim Preserve TermQty(0 To UBound(TermQty) + 1) As Integer
                    TermQty(UBound(TermQty)) = Qty
                                
                Else
                            
                    Duplicate = False
                    For y = 1 To UBound(Term)
                        If Worksheets(B).Cells(z, 4).Value = Term(y) Then
                            Duplicate = True
                            ArrayLoc = y
                            GoTo DuplicateTerm
                        End If
                    Next y
                
DuplicateTerm:
        
                    If Duplicate = True Then
                        TermQty(ArrayLoc) = TermQty(ArrayLoc) + Qty
                    Else
                        ReDim Preserve Term(0 To UBound(Term) + 1) As String
                        Term(UBound(Term)) = Worksheets(B).Cells(z, 4).Value
                                    
                        ReDim Preserve TermQty(0 To UBound(TermQty) + 1) As Integer
                        TermQty(UBound(TermQty)) = Qty
                    End If
                        
                End If
                        
            End If
        
Nextz_Term:
        Next z
        
AddSeal: '----------Add Seal----------
        
        For z = (x + 4) To EndRow
        
            If Worksheets(B).Cells(z, 5).Value = "" Then
                GoTo Nextz_Seal
            Else
            
                If UBound(Seal) = 0 Then
                            
                    ReDim Preserve Seal(0 To UBound(Seal) + 1) As String
                    Seal(UBound(Seal)) = Worksheets(B).Cells(z, 5).Value
                                
                    ReDim Preserve SealQty(0 To UBound(SealQty) + 1) As Integer
                    SealQty(UBound(SealQty)) = Qty
                                
                Else
                            
                    Duplicate = False
                    For y = 1 To UBound(Seal)
                        If Worksheets(B).Cells(z, 5).Value = Seal(y) Then
                            Duplicate = True
                            ArrayLoc = y
                            GoTo DuplicateSeal
                        End If
                    Next y
                
DuplicateSeal:
        
                    If Duplicate = True Then
                        SealQty(ArrayLoc) = SealQty(ArrayLoc) + Qty
                    Else
                        ReDim Preserve Seal(0 To UBound(Seal) + 1) As String
                        Seal(UBound(Seal)) = Worksheets(B).Cells(z, 5).Value
                                    
                        ReDim Preserve SealQty(0 To UBound(SealQty) + 1) As Integer
                        SealQty(UBound(SealQty)) = Qty
                    End If
                        
                End If
                        
            End If
        
Nextz_Seal:
        Next z
        
AddWire: '----------Add Wire----------
        
        For z = (x + 4) To EndRow
        
            If Worksheets(B).Cells(z, 6).Value = "" Then
                GoTo Nextz_Wire
            Else
            
                If UBound(Wire) = 0 Then
                            
                    ReDim Preserve Wire(0 To UBound(Wire) + 1) As String
                    Wire(UBound(Wire)) = Worksheets(B).Cells(z, 6).Value & "-TXL-" & Worksheets(B).Cells(z, 7).Value
                    
                    ReDim Preserve WireQty(0 To UBound(WireQty) + 1) As Integer
                    WireQty(UBound(WireQty)) = Qty
                
                Else
                    
                    Duplicate = False
                    For y = 1 To UBound(Wire)
                        If Worksheets(B).Cells(z, 6).Value & "-TXL-" & Worksheets(B).Cells(z, 7).Value = Wire(y) Then
                            Duplicate = True
                            ArrayLoc = y
                            GoTo DuplicateWire
                        End If
                    Next y
                
DuplicateWire:
        
                    If Duplicate = True Then
                        WireQty(ArrayLoc) = WireQty(ArrayLoc) + Qty
                    Else
                        ReDim Preserve Wire(0 To UBound(Wire) + 1) As String
                        Wire(UBound(Wire)) = Worksheets(B).Cells(z, 6).Value & "-TXL-" & Worksheets(B).Cells(z, 7).Value
                                    
                        ReDim Preserve WireQty(0 To UBound(WireQty) + 1) As Integer
                        WireQty(UBound(WireQty)) = Qty
                    End If
                        
                End If
                        
            End If
            
            '----------- Add AmpTerm ------------
            
            If Worksheets(B).Cells(z, 6).Value >= 16 Then
                TermPN = "AT60-202-16141"
            End If
            
            If Worksheets(B).Cells(z, 6).Value <= 14 Then
                TermPN = "AT60-215-16141"
            End If
            
            If UBound(AmpTerm) = 0 Then
                    
                ReDim Preserve AmpTerm(0 To UBound(AmpTerm) + 1) As String
                AmpTerm(UBound(AmpTerm)) = TermPN
                    
                ReDim Preserve AmpTermQty(0 To UBound(AmpTermQty) + 1) As Integer
                AmpTermQty(UBound(AmpTermQty)) = Qty
                
            Else
                   
                Duplicate = False
                For y = 1 To UBound(AmpTerm)
                    If TermPN = AmpTerm(y) Then
                        Duplicate = True
                        ArrayLoc = y
                        GoTo DuplicateAmpTerm
                    End If
                Next y
                
DuplicateAmpTerm:
        
                If Duplicate = True Then
                    AmpTermQty(ArrayLoc) = AmpTermQty(ArrayLoc) + Qty
                Else
                    ReDim Preserve AmpTerm(0 To UBound(AmpTerm) + 1) As String
                    AmpTerm(UBound(AmpTerm)) = TermPN
                    
                    ReDim Preserve AmpTermQty(0 To UBound(AmpTermQty) + 1) As Integer
                    AmpTermQty(UBound(AmpTermQty)) = Qty
                End If
                        
            End If
            
Nextz_Wire:
        Next z
                     
        
        '----------Add Amphenol Connectors----------
            
        For z = (x + 4) To EndRow
        
            If Worksheets(B).Cells(z, 8).Value = "" Then
                GoTo Nextz_AmpConn
            Else
            
                If UBound(AmpConn) = 0 Then
                            
                    ReDim Preserve AmpConn(0 To UBound(AmpConn) + 1) As String
                    AmpConn(UBound(AmpConn)) = Worksheets(B).Cells(z, 8).Value
                                
                    ReDim Preserve AmpConnQty(0 To UBound(AmpConnQty) + 1) As Integer
                    AmpConnQty(UBound(AmpConnQty)) = Qty
                                
                Else
                    
                    Duplicate = False
                    For y = 1 To UBound(AmpConn)
                        If Worksheets(B).Cells(z, 8).Value = AmpConn(y) Then
                            Duplicate = True
                            ArrayLoc = y
                            GoTo DuplicateAmpConn
                        End If
                    Next y
                
DuplicateAmpConn:
        
                    If Duplicate = True Then
                        AmpConnQty(ArrayLoc) = AmpConnQty(ArrayLoc) + Qty
                    Else
                        ReDim Preserve AmpConn(0 To UBound(AmpConn) + 1) As String
                        AmpConn(UBound(AmpConn)) = Worksheets(B).Cells(z, 8).Value
                                    
                        ReDim Preserve AmpConnQty(0 To UBound(AmpConnQty) + 1) As Integer
                        AmpConnQty(UBound(AmpConnQty)) = Qty
                    End If
                        
                End If
                        
            End If
        
Nextz_AmpConn:
        Next z
    
        '----------Add Amphenol Connector Wedges----------
            
        For z = (x + 4) To EndRow
        
            If Worksheets(B).Cells(z, 9).Value = "" Then
                GoTo Nextz_AmpWedg
            Else
            
                If UBound(AmpWedg) = 0 Then
                            
                    ReDim Preserve AmpWedg(0 To UBound(AmpWedg) + 1) As String
                    AmpWedg(UBound(AmpWedg)) = Worksheets(B).Cells(z, 9).Value
                                
                    ReDim Preserve AmpWedgQty(0 To UBound(AmpWedgQty) + 1) As Integer
                    AmpWedgQty(UBound(AmpWedgQty)) = Qty
                                
                Else
                    
                    Duplicate = False
                    For y = 1 To UBound(AmpWedg)
                        If Worksheets(B).Cells(z, 9).Value = AmpWedg(y) Then
                            Duplicate = True
                            ArrayLoc = y
                            GoTo DuplicateAmpWedg
                        End If
                    Next y
                
DuplicateAmpWedg:
        
                    If Duplicate = True Then
                        AmpWedgQty(ArrayLoc) = AmpWedgQty(ArrayLoc) + Qty
                    Else
                        ReDim Preserve AmpWedg(0 To UBound(AmpWedg) + 1) As String
                        AmpWedg(UBound(AmpWedg)) = Worksheets(B).Cells(z, 9).Value
                                    
                        ReDim Preserve AmpWedgQty(0 To UBound(AmpWedgQty) + 1) As Integer
                        AmpWedgQty(UBound(AmpWedgQty)) = Qty
                    End If
                        
                End If
                        
            End If
        
Nextz_AmpWedg:
        Next z
    
    ' else go to next row
    Else
        GoTo Nextx
    End If 'Instr()

Nextx:

Next x


'----------Populate BOM----------
    
i = 8
    
For x = 1 To UBound(Conn)
    Worksheets(C).Cells(i, 4).Value = Conn(x)
    Worksheets(C).Cells(i, 5).Value = ConnQty(x)
    i = i + 1
Next x
    
For x = 1 To UBound(Accy)
    Worksheets(C).Cells(i, 4).Value = Accy(x)
    Worksheets(C).Cells(i, 5).Value = AccyQty(x)
    i = i + 1
Next x

For x = 1 To UBound(Term)
    Worksheets(C).Cells(i, 4).Value = Term(x)
    Worksheets(C).Cells(i, 5).Value = TermQty(x)
    i = i + 1
Next x

For x = 1 To UBound(Seal)
    Worksheets(C).Cells(i, 4).Value = Seal(x)
    Worksheets(C).Cells(i, 5).Value = SealQty(x)
    i = i + 1
Next x

For x = 1 To UBound(AmpConn)
    Worksheets(C).Cells(i, 4).Value = AmpConn(x)
    Worksheets(C).Cells(i, 5).Value = AmpConnQty(x)
    i = i + 1
Next x

For x = 1 To UBound(AmpWedg)
    Worksheets(C).Cells(i, 4).Value = AmpWedg(x)
    Worksheets(C).Cells(i, 5).Value = AmpWedgQty(x)
    i = i + 1
Next x

For x = 1 To UBound(AmpTerm)
    Worksheets(C).Cells(i, 4).Value = AmpTerm(x)
    Worksheets(C).Cells(i, 5).Value = AmpTermQty(x)
    i = i + 1
Next x

For h = 8 To i
    Worksheets(C).Cells(h, 3).Value = h - 7
Next h

'Add Wire Header Seperator
Worksheets(C).Cells(i, 3).Value = "//Wires   (Cut wire in 2ft [610mm] sections)"
Worksheets(C).Range("C" & i & ":F" & i).Merge
Worksheets(C).Range("C" & i & ":F" & i).Interior.ThemeColor = xlThemeColorDark1
Worksheets(C).Range("C" & i & ":F" & i).Interior.Color = RGB(225, 225, 225)
Worksheets(C).Range("C" & i & ":F" & i).Font.Bold = True
i = i + 1

For x = 1 To UBound(Wire)
    Worksheets(C).Cells(i, 4).Value = Wire(x)
    Worksheets(C).Cells(i, 5).Value = WireQty(x)
    i = i + 1
Next x

'----------Format Table----------

LastRow = Worksheets(C).Cells(Rows.Count, 4).End(xlUp).Row

Worksheets(C).Range("C6:F" & LastRow).Borders(xlEdgeLeft).LineStyle = xlContinuous
Worksheets(C).Range("C6:F" & LastRow).Borders(xlEdgeBottom).LineStyle = xlContinuous
Worksheets(C).Range("C6:F" & LastRow).Borders(xlEdgeRight).LineStyle = xlContinuous
Worksheets(C).Range("C6:F" & LastRow).Borders(xlInsideVertical).LineStyle = xlContinuous
Worksheets(C).Range("C6:F" & LastRow).Borders(xlInsideHorizontal).LineStyle = xlContinuous
Worksheets(C).Range("C6:F" & LastRow).Borders(xlEdgeTop).LineStyle = xlContinuous

Worksheets(C).Columns("D:F").HorizontalAlignment = xlCenter
Worksheets(C).Columns("C:F").ColumnWidth = 20

'----------Adjust Print Area----------

LastRow = Worksheets(C).Cells(Rows.Count, 3).End(xlUp).Row
Worksheets(C).PageSetup.PrintArea = "B2:G" & LastRow + 6

'-------------------------Enable Excel Applications-------------------------
'Disable application auto-updating during code execution
Application.ScreenUpdating = True
Application.DisplayStatusBar = True
Application.EnableEvents = True

End Sub
Sub Clear_BOM()
'Clear button

'Assign workbook tab to letter
Dim C As String

'----------Sheet Names----------
C = "New Adapter BOM"

'Message prompt before clearing sheet contents
If MsgBox("Are you sure you want to clear the sheet?", vbYesNo + vbExclamation, "Clear") = vbNo Then
    Exit Sub
Else
    Worksheets(C).Range("B5:G1000").Cells.Clear
End If

End Sub

Sub Add_Loc_BOM()

'--------------------Message Prompt--------------------

Dim result As VbMsgBoxResult
result = MsgBox("Location data will be pulled from the Manufacturing Database. Please save and close the database to avoid losing data. Do you want to continue?" & vbCrLf & vbCrLf & "[Yes]: Proceed with finding locations." & vbCrLf & "[No]: Cancel find location.", vbYesNo + vbQuestion, "Confirmation")

If result = vbNo Then
    Exit Sub
End If

'-------------------------Disable Excel Applications-------------------------
'Disable application auto-updating during code execution
Application.ScreenUpdating = False
Application.DisplayStatusBar = False
Application.EnableEvents = False

'-------------------------Define Sheets -------------------------

'----------Sheet Names----------

Dim C As String
C = "New Adapter BOM"

Dim x As Integer
Dim y As Integer
Dim z As Integer

'--------------------Setup Workbooks--------------------

'Harness Database Workbook
Dim TF As Workbook
Set TF = ThisWorkbook

'Tool chart database workbook setup and open
Dim Data As Workbook

On Error Resume Next
Set Data = Workbooks.Open(Filename:="I:\Harness Manufacturing\1_Documents\Manufacturing Database (V1.1).xlsx", ReadOnly:=True)

Dim A1 As String
A1 = "Connectors"
Dim B1 As String
B1 = "Contacts"
Dim C1 As String
C1 = "Seals"
Dim F1 As String
F1 = "Accessories"
Dim G1 As String
G1 = "Terminals"
Dim H1 As String
H1 = "Fastening Hardware"
Dim I1 As String
I1 = "Circuit Elements"

'--------------------Populate Locations--------------------
LastRow = TF.Worksheets(C).Cells(Rows.Count, 4).End(xlUp).Row

For x = 8 To LastRow

    '----------Connector Locations----------
    LastRow_Data = Data.Worksheets(A1).Cells(Rows.Count, 2).End(xlUp).Row
    For y = 4 To LastRow_Data
        If Data.Worksheets(A1).Cells(y, 1).Value <> "" Then
            If UCase(TF.Worksheets(C).Cells(x, 4).Value) = UCase(Data.Worksheets(A1).Cells(y, 2).Value) Then
                
                TF.Worksheets(C).Cells(x, 6).Value = Data.Worksheets(A1).Cells(y, 1).Value 'Location
                
            End If
        End If
    Next y
    
    '----------Contact Locations----------
    LastRow_Data = Data.Worksheets(B1).Cells(Rows.Count, 3).End(xlUp).Row
    For y = 4 To LastRow_Data
        If Data.Worksheets(B1).Cells(y, 1).Value <> "" Then
            If UCase(TF.Worksheets(C).Cells(x, 4).Value) = UCase(Data.Worksheets(B1).Cells(y, 3).Value) Then
                
                TF.Worksheets(C).Cells(x, 6).Value = Data.Worksheets(B1).Cells(y, 1).Value 'Location
                
            End If
        End If
    Next y
    
    '----------Seal Locations----------
    LastRow_Data = Data.Worksheets(C1).Cells(Rows.Count, 3).End(xlUp).Row
    For y = 4 To LastRow_Data
        If Data.Worksheets(C1).Cells(y, 1).Value <> "" Then
            If UCase(TF.Worksheets(C).Cells(x, 4).Value) = UCase(Data.Worksheets(C1).Cells(y, 3).Value) Then
                
                TF.Worksheets(C).Cells(x, 6).Value = Data.Worksheets(C1).Cells(y, 1).Value 'Location
                
            End If
        End If
    Next y
    
    '----------Accessory Locations----------
    
    LastCol_Data = Data.Worksheets(F1).Cells(3, Columns.Count).End(xlToLeft).Column

    For y = 3 To LastCol_Data
        If Data.Worksheets(F1).Cells(2, y).Value <> "" Then
            If UCase(TF.Worksheets(C).Cells(x, 4).Value) = UCase(Data.Worksheets(F1).Cells(3, y).Value) Then
                
                TF.Worksheets(C).Cells(x, 6).Value = Data.Worksheets(F1).Cells(2, y).Value 'Location
                
            End If
        End If
    Next y
    
    '----------Terminal Locations----------
    LastRow_Data = Data.Worksheets(G1).Cells(Rows.Count, 2).End(xlUp).Row
    For y = 4 To LastRow_Data
        If Data.Worksheets(G1).Cells(y, 1).Value <> "" Then
            If UCase(TF.Worksheets(C).Cells(x, 4).Value) = UCase(Data.Worksheets(G1).Cells(y, 2).Value) Then
                
                TF.Worksheets(C).Cells(x, 6).Value = Data.Worksheets(G1).Cells(y, 1).Value 'Location
                
            End If
        End If
    Next y
    
    '----------Fastening Hardware Locations----------
    LastRow_Data = Data.Worksheets(H1).Cells(Rows.Count, 2).End(xlUp).Row
    For y = 4 To LastRow_Data
        If Data.Worksheets(H1).Cells(y, 1).Value <> "" Then
            If UCase(TF.Worksheets(C).Cells(x, 4).Value) = UCase(Data.Worksheets(H1).Cells(y, 2).Value) Then
                
                TF.Worksheets(C).Cells(x, 6).Value = Data.Worksheets(H1).Cells(y, 1).Value 'Location
                
            End If
        End If
    Next y
    
    '----------Circuit Elements Locations----------
    LastRow_Data = Data.Worksheets(I1).Cells(Rows.Count, 2).End(xlUp).Row
    For y = 4 To LastRow_Data
        If Data.Worksheets(I1).Cells(y, 1).Value <> "" Then
            If UCase(TF.Worksheets(C).Cells(x, 4).Value) = UCase(Data.Worksheets(I1).Cells(y, 2).Value) Then
                
                TF.Worksheets(C).Cells(x, 6).Value = Data.Worksheets(I1).Cells(y, 1).Value 'Location
                
            End If
        End If
    Next y

Next x

'----------Close Database (Read Only)----------

'Application.DisplayAlerts = False
Data.Close savechanges:=False

'-------------------------Enable Excel Applications-------------------------
'Enable application auto-updating during code execution
Application.ScreenUpdating = True
Application.DisplayStatusBar = True
Application.EnableEvents = True


End Sub

