Attribute VB_Name = "�K�w�쐬"
Option Explicit
Sub �f�B���N�g���ꊇ�쐬()
    Dim �I�s As Long, �s As Long
    Dim �� As String
    With Sheets("�K�w�t�H���_�쐬")
        �I�s = .Cells(Rows.Count, 1).End(xlUp).Row
        If �I�s = 1 Then MsgBox "�쐬����f�B���N�g�����ݒ肳��Ă��܂���": Exit Sub
        �� = "�����F" & �I�s - 1 & vbCrLf & vbCrLf & "�ꊇ�K�w�t�H���_�쐬���������s���܂����H"
        If MsgBox(��, vbYesNo) <> vbYes Then Exit Sub
        ReDim �z��(2 To �I�s)
        For �s = 2 To �I�s
            �K�w�t�H���_�쐬 (.Cells(�s, 1))
        Next
        MsgBox "�������������܂���"
    End With
End Sub
Function �K�w�t�H���_�쐬(�f�B���N�g�� As String)
    Dim ���� As Variant
    Dim �\���t�H���_ As String
    Dim �Y�� As Long
    ���� = Split(�f�B���N�g��, "\")
    �\���t�H���_ = ����(0)
    If Dir(�\���t�H���_, vbDirectory) = "" Then MkDir �\���t�H���_
    For �Y�� = 1 To UBound(����)
        �\���t�H���_ = �\���t�H���_ & "\" & ����(�Y��)
        If Dir(�\���t�H���_, vbDirectory) = "" Then MkDir �\���t�H���_
    Next
End Function
