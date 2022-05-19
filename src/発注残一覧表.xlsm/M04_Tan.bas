Attribute VB_Name = "M04_Tan"
Option Explicit

Sub ALL_DATA()
    Call Get_DATAT("A", "4")
End Sub

Sub M_DATA()
Dim strM As String
Dim strKBN As String

    If IsError(Sheets("�W�v").Range("U2")) = False Then
        strM = Sheets("�W�v").Range("U2")
        strKBN = Sheets("�W�v").Range("U1")
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

    '��������
    Sheets("�S����").Select
    If Cells(1, 19) = 0 Then Cells(1, 19) = "1"
    If Cells(2, 19) = "" Then Exit Sub
    strCD = Cells(2, 19)
    strKB = Range("W1")
    
    '�V�[�g������
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
    
    If strKBN = "1" Then '����
        strSQL = strSQL & "          And NOKDT < '" & strM & "'"
    ElseIf strKBN = "2" Then '����
        strSQL = strSQL & "          And NOKDT Like '____" & strM & "__'"
    ElseIf strKBN = "3" Then '�ȍ~
        strSQL = strSQL & "          And NOKDT >= '" & strM & "'"
    End If
    strSQL = strSQL & "              ORDER BY SIRCD, NOKDT, HDNDT, DENNO, LINNO"
    Set rsA = db.Execute(strSQL)
    
    If rsA.EOF = False Then rsA.MoveFirst

    lngR = 7
    lngG = 0
    Do Until rsA.EOF
        '�d����
        If rsA.Fields(12) <> n_TOK Then
            If lngR > 7 Then
                Cells(lngR, 6) = "�d����v"
                Cells(lngR, 11) = lngG
                lngG = 0
                Range(Cells(lngR, 6), Cells(lngR, 11)).Font.Bold = True
                Range(Cells(lngR - 1, 1), Cells(lngR - 1, 11)).Borders(xlEdgeBottom).Weight = xlThin
                lngR = lngR + 2
                If lngR > 5000 Then
                    MsgBox "�f�[�^���\����͂ݏo���܂����B"
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
                MsgBox "�f�[�^���\����͂ݏo���܂����B"
                Exit Do
            End If
        End If

        '�[��
        If rsA.Fields(0) = n_NOK Then
            Cells(lngR, 1) = ""
        Else
            Cells(lngR, 1) = Date_in(rsA.Fields(0))
            n_NOK = rsA.Fields(0)
        End If

        '�󒍓�
        If rsA.Fields(1) = n_HDN Then
            Cells(lngR, 2) = ""
        Else
            Cells(lngR, 2) = Date_in(rsA.Fields(1))
            n_HDN = rsA.Fields(1)
        End If

        '�`�[��
        If rsA.Fields(2) = n_DEN Then
            Cells(lngR, 3) = ""
        Else
            Cells(lngR, 3) = Trim(rsA.Fields(2))
            n_DEN = rsA.Fields(2)
        End If

        '�q��
        If rsA.Fields(11) = "2" Then
            Cells(lngR, 4) = "����"
        Else
            Cells(lngR, 4) = Trim(rsA.Fields(3))
        End If

        '���̑�
        For lngC = 4 To 11
            Cells(lngR, lngC + 1) = Trim(rsA.Fields(lngC))
        Next lngC
        lngG = lngG + rsA.Fields(10)
        lngR = lngR + 1
        If lngR > 5000 Then
            MsgBox "�f�[�^���\����͂ݏo���܂����B"
            Exit Do
        End If

        rsA.MoveNext
    Loop
    
    Cells(lngR, 6) = "�d����v"
    Cells(lngR, 11) = lngG
    Range(Cells(lngR, 6), Cells(lngR, 11)).Font.Bold = True
    Range(Cells(lngR - 1, 1), Cells(lngR - 1, 11)).Borders(xlEdgeBottom).Weight = xlThin
    If lngR > 7 Then Cells(lngR, 13) = "E"
    
    If Range("A7") = "" Then
        MsgBox "���̒S���҂͔����c������܂���B"
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
    
    Const SQL1 = "SELECT ���庰��, First(���喼) FROM ����敪 WHERE (((�x�X) = '"
    Const SQL2 = "') And ((�敪) = 'S')) GROUP BY ���庰�� ORDER BY ���庰��"
    
    'strDB = DR1 & dbA
    'cnB.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & strDB
    cnB.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & dbA
    cnB.Open
    strSTN = Sheets("�S����").Cells(2, 18)
    strSQL = SQL1 & strSTN & SQL2
    rsB.Open strSQL, cnB, adOpenStatic, adLockReadOnly
    If rsB.EOF Then
        MsgBox "�f�[�^�x�[�X�ɃA�N�Z�X�o���܂���B", vbCritical
        GoTo Exit_DB
    Else
        rsB.MoveFirst
    End If
    
    'ؽĕ\����ر
    Range(Sheets("�S����").Cells(3, 19), Sheets("�S����").Cells(22, 20)).ClearContents
    
    With rsB
        If .RecordCount > 0 Then
            lngRC = 3
            Do Until .EOF
                Sheets("�S����").Cells(lngRC, 20) = .Fields(0)
                Sheets("�S����").Cells(lngRC, 19) = .Fields(1)
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
