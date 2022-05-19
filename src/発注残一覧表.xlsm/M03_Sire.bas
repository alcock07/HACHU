Attribute VB_Name = "M03_Sire"
Option Explicit

Sub Code_in()

    Dim strCode As String
    Dim strNAM  As String
    Dim lngCD   As Long
    
    strCode = Cells(3, 5)
    If strCode = "" Then Exit Sub
    lngCD = CLng(strCode)
    strCode = Format(lngCD, "0000000000000")
    strNAM = Get_NAME(strCode)
    Cells(3, 6) = strNAM
    Cells(3, 5).Select
    
    If strNAM = "" Then
        MsgBox "この仕入先コードは使用されていません。"
    End If
    
End Sub

Sub Get_All()

Dim strK    As String
Dim strM    As String

    strK = Sheets("集計").Range("U1")
    strM = Sheets("集計").Range("U2")
    
    Call Get_DATA(strK, strM)
    
End Sub

Sub Get_DATA(strK As String, strM As String)

Dim whole_time As Toriikinzoku.TimerObject
Set whole_time = Toriikinzoku.CreateTimer

Dim db      As Toriikinzoku.DataBaseAccess
Dim rsA     As ADODB.Recordset
Dim strSQL  As String
Dim strC    As String
Dim intC    As Integer
Dim intR    As Integer
Dim lngCD   As Long

    '仕入先ｺｰﾄﾞ取得
    strC = Cells(3, 5)
    If strC = "" Then Exit Sub
    lngCD = CLng(strC)
    strC = Format(lngCD, "0000000000000")
    
    Call Clear_Sht
    
    Set db = Toriikinzoku.Instance.CreateDB
    db.Connect ("process_os")

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
    strSQL = strSQL & "       DENKB"
    strSQL = strSQL & "              FROM HACTBZ"
    strSQL = strSQL & "                         WHERE SIRCD = '" & strC & "'"
    
    If strK = "1" Then '当月まで
        strSQL = strSQL & "          And NOKDT < '" & strM & "'"
    ElseIf strK = "2" Then '次月
        strSQL = strSQL & "          And NOKDT Like '____" & strM & "__'"
    ElseIf strK = "3" Then '以降
        strSQL = strSQL & "          And NOKDT >= '" & strM & "'"
    End If
    strSQL = strSQL & "              ORDER BY NOKDT, HDNDT, DENNO, LINNO"
    Set rsA = db.Execute(strSQL)

    If rsA.EOF = False Then rsA.MoveFirst
    
    intR = 7
    n_NOK = ""
    n_HDN = ""
    n_DEN = ""
    
    Do Until rsA.EOF
        '納期
        If rsA.Fields(0) = n_NOK Then
            Cells(intR, 1) = ""
        Else
            Cells(intR, 1) = Date_in(rsA.Fields(0))
            n_NOK = rsA.Fields(0)
        End If
        '発注日
        If rsA.Fields(1) = n_HDN Then
            Cells(intR, 2) = ""
        Else
            Cells(intR, 2) = Date_in(rsA.Fields(1))
            n_HDN = rsA.Fields(1)
        End If
        '伝票№
        If rsA.Fields(2) = n_DEN Then
            Cells(intR, 3) = ""
        Else
            Cells(intR, 3) = Trim(rsA.Fields(2))
            n_DEN = rsA.Fields(2)
        End If
        '倉庫
        If rsA.Fields(3) = "2" Then
            Cells(intR, 4) = "直送"
        Else
            Cells(intR, 4) = rsA.Fields(3)
        End If
        
        For intC = 3 To 10
            Cells(intR, intC + 1) = Trim(rsA.Fields(intC))
            
        Next intC
        intR = intR + 1
        If intR > 6000 Then
            MsgBox "データが表からはみ出しました！"
            Exit Do
        End If
        rsA.MoveNext
    Loop
    
    If intR > 7 Then Cells(intR - 1, 12) = "E"
    
Exit_DB:
    
    db.Disconnect
    If Not rsA Is Nothing Then
        If rsA.State = adStateOpen Then rsA.Close
        Set rsA = Nothing
    End If

    Range("A1").Select
    
    If Range("K7") = "" Then
        MsgBox "この仕入先は指定の期間 発注残がありません。"
    End If
    
    whole_time.DebugTime ("Get_DATA")
    
End Sub
