Attribute VB_Name = "Module1"

' �V�[�g������o�C�i���f�[�^���o�͂���
Public Function ExportSheetDataBinary(sheetName As String, ExportPath As String)
    Dim fso As New FileSystemObject
    Dim file As TextStream

    ' �t�@�C������
    Set file = fso.CreateTextFile(ExportPath & "/" & sheetName & ".gdb", overwrite:=True, Unicode:=False)

    '
    file.Close

End Function
