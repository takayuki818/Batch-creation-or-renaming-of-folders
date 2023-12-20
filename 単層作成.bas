Attribute VB_Name = "単層作成"
Option Explicit
Sub フォルダ一括作成()
    Dim 終行 As Long, 行 As Long
    Dim 文 As String, 作成フォルダ名 As String, リネーム名 As String
    With Sheets("単層フォルダ作成・リネーム")
        終行 = .Cells(Rows.Count, 1).End(xlUp).Row
        If 終行 = 1 Then MsgBox "フォルダ名が設定されていません": Exit Sub
        文 = "件数：" & 終行 - 1 & vbCrLf & vbCrLf & "一括フォルダ作成・リネーム処理を実行しますか？"
        If MsgBox(文, vbYesNo) <> vbYes Then Exit Sub
        文 = ""
        For 行 = 2 To 終行
            作成フォルダ名 = ThisWorkbook.Path & "\" & .Cells(行, 1)
            Select Case .Cells(行, 2)
                Case "": リネーム名 = ""
                Case Else: リネーム名 = ThisWorkbook.Path & "\" & .Cells(行, 2)
            End Select
            Select Case Dir(作成フォルダ名, vbDirectory)
                Case ""
                    MkDir 作成フォルダ名
                    If リネーム名 <> "" Then GoTo 作成後即リネーム
                Case .Cells(行, 1)
作成後即リネーム:
                    If リネーム名 <> "" Then
                        On Error GoTo リネームエラー時
                        Name 作成フォルダ名 As リネーム名
                    End If
            End Select
        Next
        Select Case 文
            Case "": 文 = "処理が完了しました"
            Case Else
                文 = "以下の作成フォルダ名のリネームに失敗しました" & vbCrLf & 文
                文 = 文 & vbCrLf & vbCrLf & "原因：「フォルダ内のファイルを編集中」" & "「同名フォルダが存在」等"
        End Select
        MsgBox 文
        Exit Sub
リネームエラー時:
        文 = 文 & vbCrLf & .Cells(行, 1)
        文 = 文 & "　→　" & .Cells(行, 2) & "：失敗"
        Resume Next
    End With
End Sub
Sub フォルダ名書き出し()
    Dim FSO As New FileSystemObject
    Dim 文 As String
    Dim 行 As Long
    Dim 起点フォルダ As Folder, フォルダ As Folder
    With Sheets("単層フォルダ作成・リネーム")
        文 = "本ツールと同階層にあるフォルダ名をA2セル以下に書き出します" & vbCrLf & vbCrLf & "処理を実行してよろしいですか？"
        If MsgBox(文, vbYesNo) <> vbYes Then Exit Sub
        Set 起点フォルダ = FSO.GetFolder(ThisWorkbook.Path)
        行 = 1
        For Each フォルダ In 起点フォルダ.SubFolders
            行 = 行 + 1
            .Cells(行, 1) = フォルダ.Name
        Next
    End With
End Sub
