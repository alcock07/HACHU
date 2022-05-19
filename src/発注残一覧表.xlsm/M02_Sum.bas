Attribute VB_Name = "M02_Sum"
Option Explicit

Sub Sel_OS()
    Sheets("�S����").Range("R1") = 1
    Call BUMON_Get
End Sub
Sub Sel_TK()
    Sheets("�S����").Range("R1") = 2
    Call BUMON_Get
End Sub

Sub Sel_Z()
    Sheets("�W�v").Range("U1") = 1
End Sub

Sub Sel_T()
    Sheets("�W�v").Range("U1") = 2
End Sub

Sub Sel_N()
    Sheets("�W�v").Range("U1") = 3
End Sub

Sub Sel_S()
    Sheets("�S����").Range("W1") = 1
End Sub

Sub Sel_C()
    Sheets("�S����").Range("W1") = 2
End Sub

'�W�v��ʂ̔����c���擾 ===================
Sub Get_SumD()

Dim whole_time As Toriikinzoku.TimerObject
Set whole_time = Toriikinzoku.CreateTimer

Dim db      As Toriikinzoku.DataBaseAccess
Dim rsA     As New ADODB.Recordset
Dim strSQL  As String
Dim strYY   As String
Dim strMM   As String
Dim strNY   As String
Dim strNM   As String
Dim lngM    As Long
Dim lngB    As Long

    '��������
    strYY = Format(Now(), "yyyy")
    strMM = Format(Now(), "mm")
    lngM = CLng(strMM) + 1
    If lngM = 13 Then
        strNY = CStr(CLng(strYY) + 1)
        strNM = "01"
    Else
        strNY = strYY
        strNM = Format(lngM, "00")
    End If
    
    '�ر
    Range("F5:K10").Select
    Selection.ClearContents
    Range("F13:K14").Select
    Selection.ClearContents
    
    'DB�ݒ�
    Set db = Toriikinzoku.Instance.CreateDB
    db.Connect ("process_os")
    strSQL = ""
    strSQL = strSQL & "SELECT BMNCD,"
    strSQL = strSQL & "       NOKDT,"
    strSQL = strSQL & "       Sum(ZANKN),"
    strSQL = strSQL & "       DENKB"
    strSQL = strSQL & "                  FROM HACTBZ"
    strSQL = strSQL & "                              GROUP BY BMNCD,"
    strSQL = strSQL & "                                       TANCD,"
    strSQL = strSQL & "                                       NOKDT,"
    strSQL = strSQL & "                                       DENKB"
    strSQL = strSQL & "                              ORDER BY BMNCD,"
    strSQL = strSQL & "                                       TANCD"
    Set rsA = db.Execute(strSQL)
    
    If rsA.EOF = False Then rsA.MoveFirst
    Do Until rsA.EOF
        '������
        If Left(rsA.Fields(1), 6) <= strYY & strMM Then
            lngM = 6
        ElseIf Left(rsA.Fields(1), 6) = strNY & strNM Then
            lngM = 7
        Else
            lngM = 8
        End If
        
        '���唻��
'        If rsA.Fields(0) = "010190" Then
'            If rsA.Fields(3) = "2" Then
'                Stop
'            End If
'        End If
        If Trim(rsA.Fields(3)) = "2" Then lngM = lngM + 3 '��������
        For lngB = 4 To 14
            If rsA.Fields(0) = Cells(lngB, 5) Then
                Cells(lngB, lngM) = Cells(lngB, lngM) + rsA.Fields(2)
            End If
        Next lngB
        rsA.MoveNext
    Loop
    
    Range("A1").Select
    whole_time.DebugTime ("Get_SumD")
    
Exit_DB:
    
    db.Disconnect
    If Not rsA Is Nothing Then
        If rsA.State = adStateOpen Then rsA.Close
        Set rsA = Nothing
    End If
    Range("A1").Select
    
End Sub


