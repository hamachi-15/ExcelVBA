Attribute VB_Name = "Module1"

' シート名からバイナリデータを出力する
Public Function ExportSheetDataBinary(sheetName As String, ExportPath As String)
    Dim fso As New FileSystemObject
    Dim file As TextStream

    ' ファイル生成
    Set file = fso.CreateTextFile(ExportPath & "/" & sheetName & ".gdb", overwrite:=True, Unicode:=False)

    '
    file.Close

End Function
