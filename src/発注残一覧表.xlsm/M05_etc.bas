Attribute VB_Name = "M05_etc"
Option Explicit

Sub PRN_SHT()
Attribute PRN_SHT.VB_Description = "マクロ記録日 : 2003/2/7  ユーザー名 : 下田　康夫"
Attribute PRN_SHT.VB_ProcData.VB_Invoke_Func = " \n14"
Dim lngR As Long
Dim strSCD As String
Dim strKIN As String

    lngR = 7
    strSCD = Cells(lngR, 1)
    strKIN = Cells(lngR, 11)
    
    If strSCD = "" And strKIN = "" Then
        MsgBox "明細がありません。"
        Exit Sub
    End If

    Do
        If Cells(lngR, 1) = "" And Cells(lngR, 11) = "" Then
            If Cells(lngR + 1, 1) = "" And Cells(lngR + 1, 11) = "" Then
                Exit Do
            End If
        End If
        lngR = lngR + 1
    Loop
    
    Range(Cells(1, 1), Cells(lngR - 1, 11)).Select
    Selection.PrintOut Copies:=1, Collate:=True
    ActiveSheet.DisplayAutomaticPageBreaks = False
    Cells(7, 1).Select

End Sub

Sub PRN_TAN()
Dim strI As String
Dim lngI As Long

    strI = Format(Now(), "m")

Retry:
    strI = InputBox("何月分の発注残を印刷しますか？" & vbCrLf & "数字を入力して下さい。　(全部出す場合はAと入れて下さい。)", "印刷", strI)
    If strI = "" Then GoTo Retry
    strI = StrConv(strI, vbNarrow + vbUpperCase)
    If strI = "A" Then
    Else
        On Error Resume Next
        lngI = CLng(strI)
        If Err Then
            lngI = MsgBox("数字を入力して下さい。", vbCritical, "エラー")
            GoTo Retry
        Else
            strI = Format(lngI, "00")
        End If
        On Error GoTo 0
    End If

    Call Get_DATAT(strI, "2")
    Call PRN_SHT

End Sub

Sub PRN_Sum()
    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
    ActiveSheet.DisplayAutomaticPageBreaks = False
End Sub

Sub 終了()

    Dim myBook As Workbook
    Dim strFN As String
    Dim boolB As Boolean
    
     Application.ReferenceStyle = xlA1
    
    Application.DisplayAlerts = False
    
    strFN = ThisWorkbook.Name 'このブックの名前
    boolB = False
    For Each myBook In Workbooks
        If myBook.Name <> strFN Then boolB = True
    Next
    If boolB Then
        ThisWorkbook.Close False  'ファイルを閉じる
    Else
        Application.Quit  'Excellを終了
        ThisWorkbook.Saved = True
        ThisWorkbook.Close False
    End If
      
End Sub

Sub TK_Search()
    UserForm1.Show
End Sub

Sub Clear_ALL()
    Call Clear_Head
    Call Clear_Sht
End Sub

Sub Clear_Head()
    Range("E3:F3").ClearContents
    ActiveWindow.ScrollRow = 7
    Range("E3").Select
End Sub

Sub Clear_Sht()
    Dim lngR As Long
    lngR = 7
    If Range("E7") = "" Then Exit Sub
    Do
        If Cells(lngR, 12) = "E" Or lngR = 6000 Then Exit Do
        lngR = lngR + 1
    Loop
    Range(Cells(7, 1), Cells(lngR, 12)).ClearContents
    Range("A1").Select
    
End Sub

Sub Clear_Head2()
    Range("S3:V22").ClearContents
    Range("S1") = 1
    Range("U1") = 1
    Range("W1") = 1
    ActiveWindow.ScrollRow = 7
End Sub

Sub Clear_Sht2()
    
    Dim lngR As Long
    
    Sheets("担当者").Select
    Application.DisplayAlerts = False
    
    lngR = 7
    If Range("A7") = "" Then Exit Sub
    Do
        If Cells(lngR, 13) = "E" Or lngR = 5000 Then Exit Do
        lngR = lngR + 1
    Loop
    
    Cells(lngR, 13).ClearContents
    lngR = 5000
    Range(Cells(7, 1), Cells(lngR, 12)).Select
    Selection.ClearContents
    Selection.MergeCells = False
    Selection.Font.Bold = False
            
    '罫線
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlHairline
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlHairline
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlHairline
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlHairline
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlHairline
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlHairline
        .ColorIndex = xlAutomatic
    End With
    
    Range(Cells(7, 1), Cells(lngR, 5)).Select
    Selection.HorizontalAlignment = xlCenter
    Range("A7").Select
    
End Sub
