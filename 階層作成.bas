Attribute VB_Name = "階層作成"
Option Explicit
Sub ディレクトリ一括作成()
    Dim 終行 As Long, 行 As Long
    Dim 文 As String
    With Sheets("階層フォルダ作成")
        終行 = .Cells(Rows.Count, 1).End(xlUp).Row
        If 終行 = 1 Then MsgBox "作成するディレクトリが設定されていません": Exit Sub
        文 = "件数：" & 終行 - 1 & vbCrLf & vbCrLf & "一括階層フォルダ作成処理を実行しますか？"
        If MsgBox(文, vbYesNo) <> vbYes Then Exit Sub
        ReDim 配列(2 To 終行)
        For 行 = 2 To 終行
            階層フォルダ作成 (.Cells(行, 1))
        Next
        MsgBox "処理が完了しました"
    End With
End Sub
Function 階層フォルダ作成(ディレクトリ As String)
    Dim 分割 As Variant
    Dim 構成フォルダ As String
    Dim 添字 As Long
    分割 = Split(ディレクトリ, "\")
    構成フォルダ = 分割(0)
    If Dir(構成フォルダ, vbDirectory) = "" Then MkDir 構成フォルダ
    For 添字 = 1 To UBound(分割)
        構成フォルダ = 構成フォルダ & "\" & 分割(添字)
        If Dir(構成フォルダ, vbDirectory) = "" Then MkDir 構成フォルダ
    Next
End Function
