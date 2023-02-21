VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim inputs As Range
    Dim num As Long
    Dim check As Long
    Dim over5 As Boolean
    Dim checknum As Integer
    Dim i, j, n, c As Integer
    Dim s As String
                
    Set inputs = Range("C2")
    num = Cells(2, "C").Value
    c = 14
    
    If Not Application.Intersect(inputs, Range(Target.Address)) Is Nothing Then
        
        check = num \ 100000
        
        If (check > 0 And check < 10) Then
            
            check = num
            
            For i = 0 To 5
                checknum = check Mod 10
                check = check \ 10
                
                If (checknum >= 6) Then
                    over5 = True
                    MsgBox "1~5까지의 숫자를 입력해주십시오."
                    Exit For
                    
                End If
            Next i
            
            
            If (Not over5) Then
                
                For j = 0 To 5
                
                    n = num Mod 10
                    num = num \ 10
                    s = ""
                
                    If n = 1 Then
                        s = "B"
                    
                    ElseIf n = 2 Then
                        s = "C"
                        
                    ElseIf n = 3 Then
                        s = "D"
                        
                    ElseIf n = 4 Then
                        s = "E"
                        
                    ElseIf n = 5 Then
                        s = "F"
                        
                    Else
                        MsgBox "오류가 발생했습니다."
                        Exit For
                            
                    End If
                    
                    Set inputs = Cells(c, s)
                    inputs.Value = inputs.Value + 1
                    c = c - 2
                    
                Next j
                
                                
            End If
        
        Else
            MsgBox "6자리 숫자를 다시 입력해주십시오."
                
        End If

        Range("C2").Select
            
    End If
               

End Sub
