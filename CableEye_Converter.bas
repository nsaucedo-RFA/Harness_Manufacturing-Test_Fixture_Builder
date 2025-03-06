Attribute VB_Name = "CableEye_Converter"

Sub AutoFit_Netlist()
'Autofit button

'Assign workbook tab to letter
Dim A As String
A = "CableEye Converter"

'Adjust cell columns
Worksheets(A).Columns("D:L").AutoFit
'Worksheets(A).Range("D7:K").HorizontalAlignment = xlCenter
    
End Sub
Sub Create_Netlist()

'Define worksheets
Dim A As String
Dim B As String
A = "CableEye Converter"

'-------------------------Unlock and Format Worksheet-------------------------

'Password Protection
'Worksheets(A).Unprotect Password:="rfamec"
'Worksheets(B).Unprotect Password:="rfamec"

'Clear netlist cells on second page connection list
Worksheets(A).Range("N7:Q1000").Cells.ClearContents

'-------------------------Disable Excel Applications-------------------------

'Disable application auto-updating during code execution
Application.ScreenUpdating = False
Application.DisplayStatusBar = False
Application.EnableEvents = False

'--------------------Define Splice Array--------------------

'Define Splice() Array
Dim SpliceID() As String
ReDim SpliceID(0 To 0) As String

Dim LastRow As Integer
LastRow = Worksheets(A).Cells(Rows.Count, 4).End(xlUp).Row 'Need to change to last row in table.

Dim x As Integer
Dim y As Integer
Dim Duplicate As Boolean

'X-HSG
For x = 7 To LastRow
    If Left(UCase(Worksheets(A).Cells(x, 9).Value), 2) = "S-" Then
        
        If UBound(SpliceID) = 0 Then
    
            ReDim Preserve SpliceID(0 To UBound(SpliceID) + 1) As String
            SpliceID(UBound(SpliceID)) = UCase(Worksheets(A).Cells(x, 9).Value)
           
        
        Else
    
            Duplicate = False
            For y = 1 To UBound(SpliceID)
                If UCase(Worksheets(A).Cells(x, 9).Value) = SpliceID(y) Then
                    Duplicate = True
                End If
            Next y
                
            If Duplicate = False Then
                    
                'Record splice ID
                ReDim Preserve SpliceID(0 To UBound(SpliceID) + 1) As String
                SpliceID(UBound(SpliceID)) = UCase(Worksheets(A).Cells(x, 9).Value)
                'MsgBox Worksheets(A).Cells(x, 9).Value
                
            Else
                GoTo Nextx1
            End If
        End If
    End If
Nextx1:
Next x

'Y-HSG

For x = 7 To LastRow
    If Left(UCase(Worksheets(A).Cells(x, 11).Value), 2) = "S-" Then
        
        If UBound(SpliceID) = 0 Then
    
            ReDim Preserve SpliceID(0 To UBound(SpliceID) + 1) As String
            SpliceID(UBound(SpliceID)) = UCase(Worksheets(A).Cells(x, 11).Value)
            'MsgBox Worksheets(A).Cells(x, 11).Value
        
        Else

            Duplicate = False
            For y = 1 To UBound(SpliceID)
                If UCase(Worksheets(A).Cells(x, 11).Value) = SpliceID(y) Then
                    Duplicate = True
                End If
            Next y
                
            If Duplicate = False Then
                    
                'Record splice ID
                ReDim Preserve SpliceID(0 To UBound(SpliceID) + 1) As String
                SpliceID(UBound(SpliceID)) = UCase(Worksheets(A).Cells(x, 11).Value)
                'MsgBox Worksheets(A).Cells(x, 11).Value
                
            Else
                GoTo Nextx2
            End If
        End If
    End If
Nextx2:
Next x

'--------------------Define Splice Replacement Components Array--------------------

'Define SComp() Array
Dim SpliceComp() As String
ReDim SpliceComp(0 To 0) As String

Dim StoreSplice() As String
ReDim StoreSplice(0 To 0) As String

Dim Splice As String
Dim FoundReplaceComp As Boolean
Dim z As Integer

For x = 1 To UBound(SpliceID)

    Splice = SpliceID(x)
    
    FoundReplaceComp = False
    
    Do Until FoundReplaceComp = True
        
        Splice = Splice
    
        'check Y-HSG for replacement component

        For y = 7 To LastRow
            If UCase(Worksheets(A).Cells(y, 9).Value) = Splice Then
                If Left(UCase(Worksheets(A).Cells(y, 11).Value), 2) = "S-" Then
                    GoTo NextRow_y
                Else
                    ReDim Preserve SpliceComp(0 To UBound(SpliceComp) + 1) As String
                    SpliceComp(UBound(SpliceComp)) = UCase(Worksheets(A).Cells(y, 11).Value) & ":" & Worksheets(A).Cells(y, 12).Value
                    FoundReplaceComp = True
                    GoTo NextSplice
                End If
            End If
NextRow_y:
        Next y
    
        'check X-HSG for replacement component
     
        For y = 7 To LastRow
            If UCase(Worksheets(A).Cells(y, 11).Value) = Splice Then
                If Left(UCase(Worksheets(A).Cells(y, 9).Value), 2) = "S-" Then
                    GoTo NextRow_y2
                Else
                    ReDim Preserve SpliceComp(0 To UBound(SpliceComp) + 1) As String
                    SpliceComp(UBound(SpliceComp)) = UCase(Worksheets(A).Cells(y, 9).Value) & ":" & Worksheets(A).Cells(y, 10).Value
                    FoundReplaceComp = True
                    GoTo NextSplice
                End If
            End If
NextRow_y2:
        Next y

        If UBound(SpliceID) <> UBound(SpliceComp) Then
        
        'check Y-HSG for replacement splice

            For y = 7 To LastRow
                If UCase(Worksheets(A).Cells(y, 9).Value) = Splice Then
                    If Left(UCase(Worksheets(A).Cells(y, 11).Value), 2) <> "S-" Then
                        GoTo NextRow_y3
                    Else
                        
                        If UBound(StoreSplice) = 0 Then
    
                            ReDim Preserve StoreSplice(0 To UBound(StoreSplice) + 1) As String
                            StoreSplice(UBound(StoreSplice)) = Splice
                            Splice = UCase(Worksheets(A).Cells(y, 11).Value)
                            GoTo NewSpliceSearch
                            
                        Else
                        
                            Duplicate = False
                            For z = 1 To UBound(StoreSplice)
                                If UCase(Worksheets(A).Cells(y, 11).Value) = StoreSplice(z) Then
                                    Duplicate = True
                                End If
                            Next z
                        
                            If Duplicate = False Then
                                ReDim Preserve StoreSplice(0 To UBound(StoreSplice) + 1) As String
                                StoreSplice(UBound(StoreSplice)) = Splice
                                Splice = UCase(Worksheets(A).Cells(y, 11).Value)
                                GoTo NewSpliceSearch
                            End If
                        End If
                    End If
                End If
NextRow_y3:
            Next y
              
            'check X-HSG for replacement splice

            For y = 7 To LastRow
                If UCase(Worksheets(A).Cells(y, 11).Value) = Splice Then
                    If Left(UCase(Worksheets(A).Cells(y, 9).Value), 2) <> "S-" Then
                        GoTo NextRow_y4
                    Else
                        
                        If UBound(StoreSplice) = 0 Then
    
                            ReDim Preserve StoreSplice(0 To UBound(StoreSplice) + 1) As String
                            StoreSplice(UBound(StoreSplice)) = Splice
                            Splice = UCase(Worksheets(A).Cells(y, 9).Value)
                            GoTo NewSpliceSearch
                            
                        Else
                        
                            Duplicate = False
                            For z = 1 To UBound(StoreSplice)
                                If Splice = StoreSplice(z) Then
                                    Duplicate = True
                                End If
                            Next z
                        
                            If Duplicate = False Then
                                ReDim Preserve StoreSplice(0 To UBound(StoreSplice) + 1) As String
                                StoreSplice(UBound(StoreSplice)) = Splice
                                Splice = UCase(Worksheets(A).Cells(y, 9).Value)
                                GoTo NewSpliceSearch
                            End If
                        End If
                    End If
                End If
NextRow_y4:
            Next y
            
            '* No splice found
            MsgBox "No equivalent circuit component found for " & Splice & "."
            ReDim Preserve SpliceComp(0 To UBound(SpliceComp) + 1) As String
            SpliceComp(UBound(SpliceComp)) = ""
            GoTo NextSplice

        End If
NewSpliceSearch:

    Loop
    
NextSplice:
ReDim StoreSplice(0 To 0) As String

Next x

'--------------------Generate CableEye Netlist--------------------

Dim X_HSG As String
Dim Y_HSG As String

For x = 7 To LastRow
    
    'Add daisy chain connection
    
    If Left(UCase(Worksheets(A).Cells(x, 9).Value), 2) = "S-" Then

        For y = 1 To UBound(SpliceID)
            If UCase(Worksheets(A).Cells(x, 9).Value) = SpliceID(y) Then
                X_HSG = SpliceComp(y)
                GoTo Next_col1
            End If
        Next y
    Else
        X_HSG = UCase(Worksheets(A).Cells(x, 9).Value) & ":" & Worksheets(A).Cells(x, 10).Value
    End If
    
Next_col1:
    
    If Left(UCase(Worksheets(A).Cells(x, 11).Value), 2) = "S-" Then

        For y = 1 To UBound(SpliceID)
            If UCase(Worksheets(A).Cells(x, 11).Value) = SpliceID(y) Then
                Y_HSG = SpliceComp(y)
                GoTo Next_col2
            End If
        Next y
    Else
        Y_HSG = UCase(Worksheets(A).Cells(x, 11).Value) & ":" & Worksheets(A).Cells(x, 12).Value
    End If
    
Next_col2:
    
    If X_HSG <> Y_HSG Then
        
        Worksheets(A).Cells(x, 14).Value = X_HSG & "," & Y_HSG
          
        'Add circuit description
        Worksheets(A).Cells(x, 17).Value = Worksheets(A).Cells(x, 4).Value & " (" & Worksheets(A).Cells(x, 6).Value & "-" & Worksheets(A).Cells(x, 7).Value & "-" & Worksheets(A).Cells(x, 5).Value & ")"
        
        'Add cable
        Worksheets(A).Cells(x, 16).Value = Worksheets(A).Cells(x, 8).Value
    
    End If
    
Next x


'-------------------------Enable Excel Applications-------------------------

'Enable application auto-updating during code execution
Application.ScreenUpdating = False
Application.DisplayStatusBar = False
Application.EnableEvents = False

''-------------------------Relock and Format Worksheet-------------------------
Worksheets(A).Columns("N:Q").AutoFit

'Worksheets(A).Range("B5:L1000").Locked = False
'Worksheets(B).Range("B4:E1000").Locked = False

'Worksheets(A).Protect Password:="rfamec", DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowSorting:=True, AllowFiltering:=True
'Worksheets(B).Protect Password:="rfamec", DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowSorting:=True, AllowFiltering:=True

End Sub

'-------------------------Unlock and Format Worksheet-------------------------
Sub Clear_To_From()

If MsgBox("Are you sure you want to clear the To/From Table?", vbYesNo + vbExclamation, "Clear") = vbNo Then
    Exit Sub
Else

    Dim A As String
    A = "CableEye Converter"
    Worksheets(A).Range("D7:L1000").Cells.ClearContents
    Worksheets(A).Columns("D:L").AutoFit
    
End If

End Sub

'-------------------------Unlock and Format Worksheet-------------------------
Sub Clear_Import_Data()

    'Define worksheets

    Dim A As String
    A = "CableEye Converter"
    
    'Clear netlist cells on second page connection list
    Worksheets(A).Range("N7:Q1000").Cells.ClearContents
    Worksheets(A).Columns("N:Q").AutoFit

End Sub

Sub Copy_Import_Data()

    Dim A As String
    A = "CableEye Converter"
     
Dim LastRow As Integer
LastRow = Worksheets(A).Cells(Rows.Count, 4).End(xlUp).Row 'Need to change to last row in table.

Worksheets(A).Range("N7:Q" & LastRow).Select
    Selection.Copy


End Sub


