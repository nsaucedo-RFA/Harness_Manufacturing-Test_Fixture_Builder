Attribute VB_Name = "Component_List"
Sub Check_Adapters()

'-------------------------Disable Excel Applications-------------------------

'Disable application auto-updating during code execution
Application.ScreenUpdating = False
Application.DisplayStatusBar = False
Application.EnableEvents = False

'-------------------------Set Variables-------------------------
Dim x As Integer
Dim y As Integer
Dim LastRow As Integer
Dim Data_LastCol As Integer
Dim Data_LastRow As Integer

Dim A As String
A = "Component List"

'-------------------------Setup Workbooks -----------------------
            
'Harness Database Workbook
Dim TFB As Workbook
Set TFB = ThisWorkbook
            
'Tool chart database workbook setup and open
Dim Data As Workbook
On Error Resume Next
Set Data = Workbooks.Open(Filename:="I:\Harness Manufacturing\1_Documents\Manufacturing Database (V1.1).xlsx", ReadOnly:=True)
            
Dim T As String
T = "Test Adapters"

'-------------------------Loop Through Test Adapaters -----------------------
Dim SearchPN As Range
Dim PN As String

LastRow = TFB.Worksheets(A).Cells(Rows.Count, 5).End(xlUp).Row
Data_LastRow = Data.Worksheets(T).Cells(Rows.Count, 1).End(xlUp).Row

For x = 7 To LastRow
    
    If TFB.Worksheets(A).Cells(x, 9).Value <> "" Or TFB.Worksheets(A).Cells(x, 10).Value <> "" Then
        GoTo Nextx1
    End If
    
    If TFB.Worksheets(A).Cells(x, 6).Value <> "" Then
        PN = UCase(Trim(TFB.Worksheets(A).Cells(x, 6).Value))
    Else
        
        If TFB.Worksheets(A).Cells(x, 5).Value <> "" Then
            PN = UCase(Trim(TFB.Worksheets(A).Cells(x, 5).Value))
        Else
            GoTo Nextx1
        End If
    
        GoTo Nextx1
    
    End If
    
    For z = 4 To Data_LastRow
    
        If PN = UCase(Trim(Data.Worksheets(T).Cells(z, 3).Value)) Then
            
            If Data.Worksheets(T).Cells(z, 3).Interior.TintAndShade <> 0 Then

                For y = z To 4 Step -1
                
                    If Data.Worksheets(T).Range("A" & y).Interior.TintAndShade = 0 And Data.Worksheets(T).Range("A" & y).Interior.Pattern = xlNone Then
                            
                        If TFB.Worksheets(A).Cells(x, 9).Value = "" Then
                            TFB.Worksheets(A).Cells(x, 9).Value = Data.Worksheets(T).Cells(y, 2).Value 'Adapter PN
                        Else
                            TFB.Worksheets(A).Cells(x, 9).Value = TFB.Worksheets(A).Cells(x, 9).Value & Chr(10) & Data.Worksheets(T).Cells(y, 2).Value
                        End If
                                
                        If TFB.Worksheets(A).Cells(x, 10).Value = "" Then
                            TFB.Worksheets(A).Cells(x, 10).Value = Data.Worksheets(T).Cells(y, 3).Value 'ADapter Component PN
                        Else
                            TFB.Worksheets(A).Cells(x, 10).Value = TFB.Worksheets(A).Cells(x, 10).Value & Chr(10) & Data.Worksheets(T).Cells(y, 3).Value
                        End If
                            
                        GoTo NextSearch
                    End If
                    
                Next y
                
            End If
            
        End If
NextSearch:
    Next z

Nextx1:
Next x

'-------------------------Close workbooks-------------------------

Data.Close savechanges:=False

'-------------------------Format Cells-------------------------
        
TFB.Worksheets(A).Columns("I:J").HorizontalAlignment = xlCenter
TFB.Worksheets(A).Columns("I:J").AutoFit

'-------------------------Enable Excel Applications-------------------------

Application.ScreenUpdating = True
Application.DisplayStatusBar = True
Application.EnableEvents = True

End Sub

Sub Clear_Component_List()

Dim A As String
A = "Component List"

Worksheets(A).Range("B7:K150").ClearContents

End Sub


