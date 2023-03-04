Sub jingCalculate()
    Flush
End Sub
Sub scroller()
    Range("L3:N3").Calculate
End Sub
Sub button()
    CommandButton2_Click
End Sub
Sub Flush()
    Dim num As Double
    Dim n As Long
    
    For n = 1 To 1000
        Range("F5").Value = n
        Application.Calculate
        num = Range("E6").Value
        If n = 1000 Then
            MsgBox "未能刷出符合要求的数据，请检查参数设置"
        ElseIf num = 0 Then
            Exit For
        End If
    Next n
End Sub
Sub CommandButton2_Click()
    Range("A1:AO187").Copy
    Sheets("数据固定").Activate
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Worksheets("输出表").Calculate
    ActiveWindow.Zoom = 100
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
    Sheets("输出表").Activate
    ActiveWindow.Zoom = 100
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
End Sub

Sub SpinButton1_Click()
    Dim 位数1 As Integer
    位数1 = Range("J2").Value
    Range("D177:D184,F12:M71,T12:AA71,AH12:AO71,T76:V90,D76:F90,L76:N90,E95:L100,E103:L108,E111:L116,E119:L124,E127:L132,E135:L140,E143:L148,E151:L156,E159:L164").NumberFormat = "General"
    Select Case 位数1
        Case 0:
            Range("D177:D184,F12:M71,T12:AA71,AH12:AO71,T76:V90,D76:F90,L76:N90,E95:L100,E103:L108,E111:L116,E119:L124,E127:L132,E135:L140,E143:L148,E151:L156,E159:L164").NumberFormat = "0"
        Case 1:
            Range("D177:D184,F12:M71,T12:AA71,AH12:AO71,T76:V90,D76:F90,L76:N90,E95:L100,E103:L108,E111:L116,E119:L124,E127:L132,E135:L140,E143:L148,E151:L156,E159:L164").NumberFormat = "0.0"
        Case 2:
            Range("D177:D184,F12:M71,T12:AA71,AH12:AO71,T76:V90,D76:F90,L76:N90,E95:L100,E103:L108,E111:L116,E119:L124,E127:L132,E135:L140,E143:L148,E151:L156,E159:L164").NumberFormat = "0.00"
        Case 3:
            Range("D177:D184,F12:M71,T12:AA71,AH12:AO71,T76:V90,D76:F90,L76:N90,E95:L100,E103:L108,E111:L116,E119:L124,E127:L132,E135:L140,E143:L148,E151:L156,E159:L164").NumberFormat = "0.000"
        Case Else:
            Range("D177:D184,F12:M71,T12:AA71,AH12:AO71,T76:V90,D76:F90,L76:N90,E95:L100,E103:L108,E111:L116,E119:L124,E127:L132,E135:L140,E143:L148,E151:L156,E159:L164").NumberFormat = "0.0000"
    End Select
    Range("F5").Value = 1
    Application.Calculate
    ActiveWindow.Zoom = 100
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
End Sub
Sub 复制数据1()
    Dim p As Integer
    p = Range("A3").Value2
    ActiveWindow.ScrollRow = 191
    If Range("O1").Value2 = "true" Then
        Range("B2:S182").Copy
    ElseIf Range("Q1").Value2 = "true" Then
        Range("G2:N182").Copy
    Else
        Select Case p
            Case 3
                Range("B2:I182").Copy
            Case 6
                Range("B2:L182").Copy
            Case 8
                Range("B2:N182").Copy
            Case Else
                ActiveWindow.ScrollRow = 1
                MsgBox "无法自动选择！"
        End Select
    End If
End Sub
Sub 复制数据2()
    Dim p As Integer
    p = Range("A3").Value
    ActiveWindow.ScrollRow = 254
    If Range("O1").Value = "true" Then
        Range("D184:S256").Copy
    ElseIf Range("Q1").Value = "true" Then
        Range("G184:N256").Copy
    Else
        Select Case p
            Case 3
                Range("D184:I256").Copy
            Case 6
                Range("D184:L256").Copy
            Case 8
                Range("D184:N256").Copy
            Case Else
                ActiveWindow.ScrollRow = 191
                MsgBox "无法自动选择！"
        End Select
    End If
End Sub
Sub 复制数据3()
    Dim p As Integer
    p = Range("A3").Value
    ActiveWindow.ScrollRow = 264
    If Range("O1").Value = "true" Then
        Range("F258:S263").Copy
    ElseIf Range("Q1").Value = "true" Then
        Range("G258:N263").Copy
    Else
        Select Case p
            Case 3
                Range("F258:I263").Copy
            Case 6
                Range("F258:L263").Copy
            Case 8
                Range("F258:N263").Copy
            Case Else
                ActiveWindow.ScrollRow = 254
                MsgBox "无法自动选择！"
                Exit Sub
        End Select
    End If
End Sub
Sub 复制数据4()
    Dim p As Long
    Dim ws As Worksheet
    
    Set ws = ThisWorkbook.Sheets("数据输入及生成页")
    p = Range("A3").Value2
    
    ws.Activate
    ActiveWindow.Zoom = 100
    
    If ThisWorkbook.Worksheets("输出表").Range("O1").Value2 = "true" Then
        ThisWorkbook.Worksheets("输出表").Range("B265:O279").Copy
    ElseIf ThisWorkbook.Worksheets("输出表").Range("Q1").Value2 = "true" Then
        ThisWorkbook.Worksheets("输出表").Range("B267:O274").Copy
    Else
        Select Case p
            Case 3
                ThisWorkbook.Worksheets("输出表").Range("B265:O269").Copy
            Case 6
                ThisWorkbook.Worksheets("输出表").Range("B265:O272").Copy
            Case 8
                ThisWorkbook.Worksheets("输出表").Range("B265:O274").Copy
            Case Else
                ThisWorkbook.Worksheets("输出表").Activate
                ActiveWindow.ScrollRow = 264
                MsgBox "无法自动选择！"
        End Select
    End If
End Sub
