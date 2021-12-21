Attribute VB_Name = "Module3"
Option Explicit

' �\�[�X�R�[�h�̐擪�ɋL������R�����g��
Public Const CommentPrefixStart = "/*!"
Public Const CommentPrefixFile = vbTab & "@file" & vbTab
Public Const CommentPrefixEnum = vbTab & "@enum" & vbTab
Public Const CommentPrefixStruct = vbTab & "@struct" & vbTab
Public Const CommentPrefixBrief = vbTab & "@brief" & vbTab
Public Const CommentPrefixAutor = vbTab & "@autor" & vbTab
Public Const CommentPrefixData = vbTab & "@data" & vbTab
Public Const CommentPrefixEnd = "**/"

Public Const AutoGeneratedBrief = "�c�[���ɂĎ����������Ă��܂��B��΂ɏ��������Ȃ��ł��������B"
Public Const IncludeGurard = "#pragma" & vbTab & "once"

'API��`
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

' �����J�n�Z����
Public Const IndexStartName = "A3"

' �����FTargetArray -> �Y�������m�F�������z��
' �����Felement -> �v�f(������)
' �߂�l�F�C���f�b�N�X�ԍ�
Public Function IndexOf(TargetArray, element)
    Dim i As Integer
    For i = 0 To UBound(TargetArray)
        If TargetArray(i) = element Then Exit For
    Next
    IndexOf = i
End Function

' �����FTargetArray -> �Y�������m�F�������z��
' �����Felement -> �v�f(������)
' �߂�l�F�C���f�b�N�X�ԍ�
Public Function LongToBytes(ByRef valueArray() As Byte, ByRef value As Long)

    Call CopyMemory(valueArray(LBound(valueArray)), value, 4)

End Function

' �����FTargetArray -> �Y�������m�F�������z��
' �����Felement -> �v�f(������)
' �߂�l�F�C���f�b�N�X�ԍ�
Public Function SingleToBytes(ByRef valueArray() As Byte, ByRef value As Single)

    Call CopyMemory(valueArray(LBound(valueArray)), value, 4)

End Function

' �����FsheetName -> �V�[�g��
' �����Felement -> �v�f��
' �߂�l�F�w�肵�����O�̃V�[�g�����݂��邩
Public Function ExistsWorksheet(ByVal name As String)

    Dim ws As Worksheet
    For Each ws In Sheets
        If ws.name = name Then
            ' ���݂���
            ExistsWorksheet = True
            Exit Function
        End If
    Next
    
    ' ���݂��Ȃ�
    ExistsWorksheet = False
End Function

' �����FsheetName -> �V�[�g��
' �����Felement -> �v�f��
' �߂�l�F�C���f�b�N�X�ԍ�
Public Function SearchSheetElementId(sheetName As String, element As String)
    Dim normalizedRowCounter As Integer
    Dim normalizedCollumnCounter As Integer
    Dim maxRowCount As Integer
    
    Dim workText As String
    Dim Worksheet As Worksheet
    
    ' �V�[�g�擾
    Set Worksheet = Sheets(sheetName)

    ' �w��̃Z��������f�[�^��`
    normalizedRowCounter = Worksheet.Range(IndexStartName).Row
    normalizedCollumnCounter = Worksheet.Range(IndexStartName).Column

    ' Id�����擾����
    maxRowCount = 1
    workText = Worksheet.Cells(normalizedRowCounter + maxRowCount + 1, normalizedCollumnCounter + 1).value
    While workText <> ""
        ' ���̕����𒲂ׂ�
        maxRowCount = maxRowCount + 1
        workText = Worksheet.Cells(normalizedRowCounter + maxRowCount + 1, normalizedCollumnCounter + 1).value
    Wend

    SearchSheetElementId = 0
    workText = Worksheet.Cells(normalizedRowCounter + SearchSheetElementId + 1, normalizedCollumnCounter).value
    While workText <> element

        SearchSheetElementId = SearchSheetElementId + 1
        workText = Worksheet.Cells(normalizedRowCounter + SearchSheetElementId, normalizedCollumnCounter + 2).value
    Wend

    workText = Worksheet.Cells(normalizedRowCounter + SearchSheetElementId, normalizedCollumnCounter + 1).value
    maxRowCount = Val(workText)
    SearchSheetElementId = maxRowCount

End Function

' �����FTypeText -> �^������
' �����FElement -> �v�f������
' �����FdataTextByteArray -> �o�C�g������
' �����FDefinisionDataMap -> �f�[�^�}�b�v
' �߂�l�F�C���f�b�N�X�ԍ�
Public Function ConvertValue(TypeText As String, element As String, ByRef dataTextByteArray As Variant, ByRef DefinisionDataMap As Dictionary)
    Dim workInt As Long
    Dim workFloat As Single
    Dim byteIntArray(3) As Byte
    Dim byteFloatArray(3) As Byte
    Dim byteBooleanArray(3) As Byte
    Dim byteStringArray() As Byte
    If TypeText = "int" Then
        ' int�^
        workInt = Val(element)
        Call LongToBytes(byteIntArray, workInt)
        ConvertValue = byteIntArray
    ElseIf TypeText = "float" Then
        ' float�^
        workFloat = Val(element)
        Call SingleToBytes(byteFloatArray, workFloat)
        ConvertValue = byteFloatArray
    ElseIf TypeText = "bool" Then
        ' bool�^
        If element = "True" Then
            Call LongToBytes(byteBooleanArray, 1)
        Else
            Call LongToBytes(byteBooleanArray, 0)
        End If
        ConvertValue = byteBooleanArray
    Else
        ' enum�萔�����Ă��邷��
        ' DefinisionDataMap
        If DefinisionDataMap.Exists(TypeText) = True Then
            If element = "" Then
                workInt = 0 ' ��Ȃ�0������
            Else
                workInt = IndexOf(DefinisionDataMap.Item(TypeText), element)
            End If
            ' �C���f�b�N�X������
            Call LongToBytes(byteIntArray, workInt)
            ConvertValue = byteIntArray
        Else
            ' �^���f�[�^�V�[�g�Ƃ��đ��݂��Ă��邩�`�F�b�N
            If ExistsWorksheet(TypeText) = True Then
                ' �^����V�[�g����������Id���擾����
                workInt = SearchSheetElementId(TypeText, element)

                ' �C���f�b�N�X������
                Call LongToBytes(byteIntArray, workInt)
                ConvertValue = byteIntArray
            Else
                ' string�^��ϊ�
                byteStringArray = StrConv(element, vbFromUnicode)
    
                ' ���݂̏I�[�ʒu��Id�Ƃ���
                Dim workIndex As Long
                workIndex = UBound(dataTextByteArray) + 1
                Call LongToBytes(byteIntArray, workIndex)
                ConvertValue = byteIntArray
                
                ' ������z���ۑ�����
                ReDim Preserve dataTextByteArray(workIndex + UBound(byteStringArray) + 1)
                For workInt = 0 To UBound(byteStringArray)
                    dataTextByteArray(workIndex + workInt) = byteStringArray(workInt)
                Next workInt
                
                ' �I�[������ǉ�
                dataTextByteArray(UBound(dataTextByteArray)) = 0
            End If
        End If
    End If
End Function
