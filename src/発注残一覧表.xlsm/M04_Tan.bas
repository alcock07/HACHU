Attribute VB_Name = "M04_Tan"
Option Explicit

Sub ALL_DATA()
    Call Get_DATAT("A", "4")
End Sub

Sub M_DATA()
Dim strM As String
Dim strKBN As String

    If IsError(Sheets("集計").Range("U2")) = False Then
        strM = Sheets("集計").Range("U2")
        strKBN = Sheets("集計").Range("U1")
        Call Get_DATAT(strM, strKBN)
    End If
    
End Sub

Sub Get_DATAT(strM As String, strKBN As String)

Dim whole_time As Toriikinzoku.TimerObject
Set whole_time = Toriikinzoku.CreateTimer

Dim db      As Toriikinzoku.DataBaseAccess
Dim rsA As New ADODB.Recordset
Dim strSQL As String
Dim strCD  As String
Dim strKB  As String
Dim strNW  As String
Dim lngC   As Long
Dim lngR   As Long
Dim lngG   As Long

    'ｺｰﾄﾞﾁｪｯｸ
    Sheets("担当者").Select
    If Cells(1, 19) = 0 Then Cells(1, 19) = "1"
    If Cells(2, 19) = "" Then Exit Sub
    strCD = Cells(2, 19)
    strKB = Range("W1")
    
    'シート初期化
    Call Clear_Sht2
    
    Set db = Toriikinzoku.Instance.CreateDB
    db.Connect ("process_os")
    
    strNW = Format(Now(), "yyyymmdd")
    
    strSQL = ""
    strSQL = strSQL & "SELECT NOKDT,"
    strSQL = strSQL & "       HDNDT,"
    strSQL = strSQL & "       DENNO,"
    strSQL = strSQL & "       SOKONM,"
    strSQL = strSQL & "       HINCD,"
    strSQL = strSQL & "       HINNM,"
    strSQL = strSQL & "       SODSU,"
    strSQL = strSQL & "       SODTK,"
    strSQL = strSQL & "       SODKN,"
    strSQL = strSQL & "       ZANSU,"
    strSQL = strSQL & "       ZANKN,"
    strSQL = strSQL & "       DENKB,"
    strSQL = strSQL & "       SIRCD,"
    strSQL = strSQL & "       SIRNM"
    strSQL = strSQL & "              FROM HACTBZ"
    strSQL = strSQL & "                         WHERE DENKB = '" & strKB & "'"
    strSQL = strSQL & "                         And   BMNCD = '" & strCD & "'"
    
    If strKBN = "1" Then '当月
        strSQL = strSQL & "          And NOKDT < '" & strM & "'"
    ElseIf strKBN = "2" Then '次月
        strSQL = strSQL & "          And NOKDT Like '____" & strM & "__'"
    ElseIf strKBN = "3" Then '以降
        strSQL = strSQL & "          And NOKDT >= '" & strM & "'"
    End If
    strSQL = strSQL & "              ORDER BY SIRCD, NOKDT, HDNDT, DENNO, LINNO"
    Set rsA = db.Execute(strSQL)
    
    If rsA.EOF = False Then rsA.MoveFirst

    lngR = 7
    lngG = 0
    Do Until rsA.EOF
        '仕入先
        If rsA.Fields(12) <> n_TOK Then
            If lngR > 7 Then
                Cells(lngR, 6) = "仕入先計"
                Cells(lngR, 11) = lngG
                lngG = 0
                Range(Cells(lngR, 6), Cells(lngR, 11)).Font.Bold = True
                Range(Cells(lngR - 1, 1), Cells(lngR - 1, 11)).Borders(xlEdgeBottom).Weight = xlThin
                lngR = lngR + 2
                If lngR > 5000 Then
                    MsgBox "データが表からはみ出しました。"
                    Exit Do
                End If
            End If
            Cells(lngR, 1) = Right(rsA.Fields(12), 6)
            Range(Cells(lngR, 2), Cells(lngR, 5)).Merge
            Range(Cells(lngR, 2), Cells(lngR, 5)).HorizontalAlignment = xlLeft
            Cells(lngR, 2) = rsA.Fields(13)
            Range(Cells(lngR, 1), Cells(lngR, 5)).Font.Bold = True
            n_TOK = rsA.Fields(12)
            n_NOK = ""
            n_HDN = ""
            lngG = 0
            lngR = lngR + 1
            If lngR > 5000 Then
                MsgBox "データが表からはみ出しました。"
                Exit Do
            End If
        End If

        '納期
        If rsA.Fields(0) = n_NOK Then
            Cells(lngR, 1) = ""
        Else
            Cells(lngR, 1) = Date_in(rsA.Fields(0))
            n_NOK = rsA.Fields(0)
        End If

        '受注日
        If rsA.Fields(1) = n_HDN Then
            Cells(lngR, 2) = ""
        Else
            Cells(lngR, 2) = Date_in(rsA.Fields(1))
            n_HDN = rsA.Fields(1)
        End If

        '伝票№
        If rsA.Fields(2) = n_DEN Then
            Cells(lngR, 3) = ""
        Else
            Cells(lngR, 3) = Trim(rsA.Fields(2))
            n_DEN = rsA.Fields(2)
        End If

        '倉庫
        If rsA.Fields(11) = "2" Then
            Cells(lngR, 4) = "直送"
        Else
            Cells(lngR, 4) = Trim(rsA.Fields(3))
        End If

        'その他
        For lngC = 4 To 11
            Cells(lngR, lngC + 1) = Trim(rsA.Fields(lngC))
        Next lngC
        lngG = lngG + rsA.Fields(10)
        lngR = lngR + 1
        If lngR > 5000 Then
            MsgBox "データが表からはみ出しました。"
            Exit Do
        End If

        rsA.MoveNext
    Loop
    
    Cells(lngR, 6) = "仕入先計"
    Cells(lngR, 11) = lngG
    Range(Cells(lngR, 6), Cells(lngR, 11)).Font.Bold = True
    Range(Cells(lngR - 1, 1), Cells(lngR - 1, 11)).Borders(xlEdgeBottom).Weight = xlThin
    If lngR > 7 Then Cells(lngR, 13) = "E"
    
    If Range("A7") = "" Then
        MsgBox "この担当者は発注残がありません。"
    End If
    
Exit_DB:

    db.Disconnect
    If Not rsA Is Nothing Then
        If rsA.State = adStateOpen Then rsA.Close
        Set rsA = Nothing
    End If

    Range("A7").Select
    whole_time.DebugTime ("Get_DATA")
    
End Sub

Sub BUMON_Get()
    Call Proc_BUMON
End Sub

Sub Proc_BUMON()
    
    Dim cnB    As New ADODB.Connection
    Dim rsB    As New ADODB.Recordset
    Dim strSQL As String
    Dim strSTN As String
    Dim lngRC  As Long
    Dim boolC  As Boolean
    
    Const SQL1 = "SELECT 部門ｺｰﾄﾞ, First(部門名) FROM 部門区分 WHERE (((支店) = '"
    Const SQL2 = "') And ((区分) = 'S')) GROUP BY 部門ｺｰﾄﾞ ORDER BY 部門ｺｰﾄﾞ"
    
    'strDB = DR1 & dbA
    'cnB.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & strDB
    cnB.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & dbA
    cnB.Open
    strSTN = Sheets("担当者").Cells(2, 18)
    strSQL = SQL1 & strSTN & SQL2
    rsB.Open strSQL, cnB, adOpenStatic, adLockReadOnly
    If rsB.EOF Then
        MsgBox "データベースにアクセス出来ません。", vbCritical
        GoTo Exit_DB
    Else
        rsB.MoveFirst
    End If
    
    'ﾘｽﾄ表示域ｸﾘｱ
    Range(Sheets("担当者").Cells(3, 19), Sheets("担当者").Cells(22, 20)).ClearContents
    
    With rsB
        If .RecordCount > 0 Then
            lngRC = 3
            Do Until .EOF
                Sheets("担当者").Cells(lngRC, 20) = .Fields(0)
                Sheets("担当者").Cells(lngRC, 19) = .Fields(1)
                lngRC = lngRC + 1
                If lngRC = 23 Then Exit Do
                .MoveNext
            Loop
        End If
    End With
    
    Range("S1") = ""
    
Exit_DB:
    rsB.Close
    cnB.Close
    Set rsB = Nothing
    Set cnB = Nothing

End Sub
