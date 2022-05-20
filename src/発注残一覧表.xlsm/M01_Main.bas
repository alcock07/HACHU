Attribute VB_Name = "M01_Main"
Option Explicit

Public n_TOK        As String
Public n_NOK        As String
Public n_HDN        As String
Public n_DEN        As String
Public strTOK(1, 9) As String
Public strDB        As String

Sub 仕入先()
    If Sheets("集計").Range("U1") = 4 Then
        Sheets("集計").Range("U1") = 1
    End If
    Sheets("仕入先").Select
    Call Clear_Sht
    Range("E3:F3") = ""
    Range("E3").Select
End Sub

Sub 担当者()
    Sheets("担当者").Select
    Range("S1") = ""
    Call Clear_Sht2
    Range("A1").Select
End Sub

Sub 集計()
    Sheets("集計").Select
    Call Get_SumD
    Range("A1").Select
End Sub

Function Get_NAME(strC As String) As String

'定数の宣言
Const SQL1 = "SELECT * FROM 仕入先 WHERE (((CODE)='"
Const SQL2 = "'))"

'変数の宣言
Dim cnA    As New ADODB.Connection
Dim db     As Toriikinzoku.DataBaseAccess
Dim rsA    As ADODB.Recordset
Dim strSQL As String

    Set db = Toriikinzoku.Instance.CreateDB
    db.Connect ("process_os")

    'レコードセットのオープン
    Set rsA = New ADODB.Recordset
    strSQL = ""
    strSQL = strSQL & "SELECT * FROM OPENQUERY ([ORA],"
    strSQL = strSQL & "                              'SELECT SIRNMA,"
    strSQL = strSQL & "                                      SIRNMB"
    strSQL = strSQL & "                               FROM SIRMTA"
    strSQL = strSQL & "                               WHERE SIRCD = ''" & strC & "''"
    strSQL = strSQL & "                              ')"
    Set rsA = db.Execute(strSQL)

    If rsA.EOF = False Then
        rsA.MoveFirst
        Get_NAME = Trim(rsA.Fields(0)) & " " & Trim(rsA.Fields(1))
    Else
        Get_NAME = ""
    End If
    
    Call Clear_Sht
    
Exit_DB:

    '接続のクローズ
    If Not rsA Is Nothing Then
        If rsA.State = adStateOpen Then rsA.Close
        Set rsA = Nothing
    End If
    db.Disconnect
    
End Function

Function Date_in(strDate As String) As String
    Dim DateA As Date
    
    If Len(strDate) = 8 Then
        DateA = CDate(Left(strDate, 4) & "/" & Mid(strDate, 5, 2) & "/" & Right(strDate, 2))
        Date_in = Format(DateA, "yy/mm/dd")
    Else
        Date_in = ""
    End If
    
End Function

Sub LB_Set()
    ActiveSheet.Shapes("LB01").ScaleHeight 2#, msoFalse, msoScaleFromTopLeft
    ActiveSheet.Shapes("LB01").ScaleHeight 0.5, msoFalse, msoScaleFromTopLeft
    ActiveSheet.Shapes("LB02").ScaleHeight 2#, msoFalse, msoScaleFromTopLeft
    ActiveSheet.Shapes("LB02").ScaleHeight 0.5, msoFalse, msoScaleFromTopLeft
    ActiveSheet.Shapes("LB03").ScaleHeight 2#, msoFalse, msoScaleFromTopLeft
    ActiveSheet.Shapes("LB03").ScaleHeight 0.5, msoFalse, msoScaleFromTopLeft
    ActiveSheet.Shapes("LB04").ScaleHeight 2#, msoFalse, msoScaleFromTopLeft
    ActiveSheet.Shapes("LB04").ScaleHeight 0.5, msoFalse, msoScaleFromTopLeft
End Sub
