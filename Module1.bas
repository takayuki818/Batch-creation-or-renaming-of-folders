Attribute VB_Name = "Module1"
Option Explicit
Sub �t�H���_�ꊇ�쐬()
    Dim �I�s As Long, �s As Long
    Dim �� As String, �쐬�t�H���_�� As String, ���l�[���� As String
    With Sheets("���̓t�H�[��")
        �I�s = .Cells(Rows.Count, 1).End(xlUp).Row
        If �I�s = 1 Then MsgBox "�t�H���_�����ݒ肳��Ă��܂���": Exit Sub
        �� = "�����F" & �I�s - 1 & vbCrLf & vbCrLf & "�ꊇ�t�H���_�쐬�E���l�[�����������s���܂����H"
        If MsgBox(��, vbYesNo) <> vbYes Then Exit Sub
        �� = ""
        For �s = 2 To �I�s
            �쐬�t�H���_�� = ThisWorkbook.Path & "\" & .Cells(�s, 1)
            Select Case .Cells(�s, 2)
                Case "": ���l�[���� = ""
                Case Else: ���l�[���� = ThisWorkbook.Path & "\" & .Cells(�s, 2)
            End Select
            Select Case Dir(�쐬�t�H���_��, vbDirectory)
                Case ""
                    MkDir �쐬�t�H���_��
                    If ���l�[���� <> "" Then GoTo �쐬�㑦���l�[��
                Case .Cells(�s, 1)
�쐬�㑦���l�[��:
                    If ���l�[���� <> "" Then
                        On Error GoTo ���l�[���G���[��
                        Name �쐬�t�H���_�� As ���l�[����
                    End If
            End Select
        Next
        Select Case ��
            Case "": �� = "�������������܂���"
            Case Else
                �� = "�ȉ��̍쐬�t�H���_���̃��l�[���Ɏ��s���܂���" & vbCrLf & ��
                �� = �� & vbCrLf & vbCrLf & "�����F�u�t�H���_���̃t�@�C����ҏW���v" & "�u�����t�H���_�����݁v��"
        End Select
        MsgBox ��
        Exit Sub
���l�[���G���[��:
        �� = �� & vbCrLf & .Cells(�s, 1)
        �� = �� & "�@���@" & .Cells(�s, 2) & "�F���s"
        Resume Next
    End With
End Sub
