Attribute VB_Name = "M01_Main"
Option Explicit

'�R���s���[�^�[�����擾����֐��̐錾
'#If VBA7 Then
'    Declare PtrSafe Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
'    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'#Else
'    Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
'    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'#End If
    
'Public Const MAX_COMPUTERNAME_LENGTH = 15
Public Const dbA = "\\192.168.128.4\hb\SYS\DATA\�����c.accdb"

Public n_TOK        As String
Public n_NOK        As String
Public n_HDN        As String
Public n_DEN        As String
Public strTOK(1, 9) As String
Public strDB        As String

'Public Function CP_NAME() As String
'
'    Const COMPUTERNAMBUFFER_LENGTH = MAX_COMPUTERNAME_LENGTH + 1
'    Dim strComputerNameBuffer As String * COMPUTERNAMBUFFER_LENGTH
'    Dim lngComputerNameLength As Long
'    Dim lngWin32apiResultCode As Long
'
'    ' �R���s���[�^�[���̒�����ݒ�
'    lngComputerNameLength = Len(strComputerNameBuffer)
'    ' �R���s���[�^�[�����擾
'    lngWin32apiResultCode = GetComputerName(strComputerNameBuffer, _
'                                            lngComputerNameLength)
'    ' �R���s���[�^�[����\��
'    CP_NAME = Left(strComputerNameBuffer, InStr(strComputerNameBuffer, vbNullChar) - 1)
'
'End Function

Sub �d����()
    If Sheets("�W�v").Range("U1") = 4 Then
        Sheets("�W�v").Range("U1") = 1
    End If
    Sheets("�d����").Select
    Call Clear_Sht
    Range("E3:F3") = ""
    Range("E3").Select
End Sub

Sub �S����()
    Sheets("�S����").Select
    Range("S1") = ""
    Call Clear_Sht2
    Range("A1").Select
End Sub

Sub �W�v()
    Sheets("�W�v").Select
    Call Get_SumD
    Range("A1").Select
End Sub

Function Get_NAME(strC As String) As String

'�萔�̐錾
Const SQL1 = "SELECT * FROM �d���� WHERE (((CODE)='"
Const SQL2 = "'))"

'�ϐ��̐錾
Dim cnA    As New ADODB.Connection
Dim db     As Toriikinzoku.DataBaseAccess
Dim rsA    As ADODB.Recordset
Dim strSQL As String

    Set db = Toriikinzoku.Instance.CreateDB
    db.Connect ("process_os")

    '���R�[�h�Z�b�g�̃I�[�v��
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

    '�ڑ��̃N���[�Y
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
