VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "名称検索"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3210
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton1_Click()

'変数の宣言
Dim db     As Toriikinzoku.DataBaseAccess
Dim rsA    As ADODB.Recordset
Dim strSQL As String
Dim strNM  As String
Dim lngR   As Long

    
    strNM = UserForm1.TextBox1.Text
    If strNM = "" Then Exit Sub
    
    Set db = Toriikinzoku.Instance.CreateDB
    db.Connect ("process_os")

    Erase strTOK
    strSQL = ""
    strSQL = strSQL & "SELECT * FROM OPENQUERY ([ORA],"
    strSQL = strSQL & "                              'SELECT SIRCD,"
    strSQL = strSQL & "                                      SIRNMA,"
    strSQL = strSQL & "                                      SIRNMB,"
    strSQL = strSQL & "                                      SIRRN"
    strSQL = strSQL & "                               FROM SIRMTA"
    strSQL = strSQL & "                               WHERE SIRNMA Like ''%" & strNM & "%''"
    strSQL = strSQL & "                               ORDER BY SIRCD"
    strSQL = strSQL & "                              ')"
    Set rsA = db.Execute(strSQL)
    
    If rsA.EOF Then
        MsgBox "仕入先が見つかりません"
    Else
        rsA.MoveFirst
    End If
    
    lngR = 0
    Do Until rsA.EOF
        strTOK(0, lngR) = rsA.Fields(0)
        strTOK(1, lngR) = Trim(rsA.Fields(1)) & " " & Trim(rsA.Fields(2))
        UserForm1.ListBox1.AddItem rsA.Fields(0) & rsA.Fields(3)
        lngR = lngR + 1
        If lngR > 9 Then Exit Do
        rsA.MoveNext
    Loop
    
Exit_DB:
    '接続のクローズ
    If Not rsA Is Nothing Then
        If rsA.State = adStateOpen Then rsA.Close
        Set rsA = Nothing
    End If

End Sub

Private Sub ListBox1_Click()

    Dim lngR   As Long
    Dim strK As String
    Dim strM As String
    
    lngR = UserForm1.ListBox1.ListIndex
    Sheets("仕入先").Range("E3") = Right(strTOK(0, lngR), 6)
    Sheets("仕入先").Range("F3") = strTOK(1, lngR)
    UserForm1.Hide
    strK = Sheets("集計").Range("U1")
    strM = Sheets("集計").Range("U2")
    
    Call Clear_Sht
    Call Get_DATA(strK, strM)
    
End Sub

Private Sub UserForm_Activate()
    UserForm1.ListBox1.Clear
    UserForm1.TextBox1.Text = ""
    UserForm1.TextBox1.SetFocus
    UserForm1.TextBox1.IMEMode = fmIMEModeOn
End Sub

