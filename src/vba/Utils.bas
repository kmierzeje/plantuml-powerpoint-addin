Attribute VB_Name = "Utils"
Private Declare PtrSafe Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long

Private Type OPENFILENAME
    lStructSize         As Long
    hwndOwner           As LongPtr
    hInstance           As LongPtr
    lpstrFilter         As String
    lpstrCustomFilter   As String
    nMaxCustFilter      As Long
    nFilterIndex        As Long
    lpstrFile           As String
    nMaxFile            As Long
    lpstrFileTitle      As String
    nMaxFileTitle       As Long
    lpstrInitialDir     As String
    lpstrTitle          As String
    flags               As Long
    nFileOffset         As Integer
    nFileExtension      As Integer
    lpstrDefExt         As String
    lCustData           As LongPtr
    lpfnHook            As LongPtr
    lpTemplateName      As String
End Type

Private Const CP_UTF8 As Long = 65001

Private Declare PtrSafe Function WideCharToMultiByte Lib "kernel32" ( _
    ByVal CodePage As Long, _
    ByVal dwFlags As Long, _
    ByVal lpWideCharStr As LongPtr, _
    ByVal cchWideChar As Long, _
    ByVal lpMultiByteStr As LongPtr, _
    ByVal cchMultiByte As Long, _
    ByVal lpDefaultChar As LongPtr, _
    ByVal lpUsedDefaultChar As Long) As Long
Public Const MB_PRECOMPOSED = &H1       'use precomposed chars

Private Declare PtrSafe Function MultiByteToWideChar Lib "kernel32" ( _
    ByVal CodePage As Long, _
    ByVal dwFlags As Long, _
    ByVal lpMultiByteStr As LongPtr, _
    ByVal cchMultiByte As Long, _
    ByVal lpWideCharStr As LongPtr, _
    ByVal cchWideChar As Long _
    ) As Long


Public Function GetSaveFileNameWinAPI(Filters() As String, Optional Title As String, Optional FileName As String) As String

    Const MAX_BUFFER As Long = 255

    Dim OFN As OPENFILENAME
    Dim FilterString As String
    Dim i As Integer
    For i = LBound(Filters) To UBound(Filters) - 1 Step 2
        FilterString = FilterString & Filters(i) & " (" & Filters(i + 1) & ")" & vbNullChar & Filters(i + 1) & vbNullChar
    Next i
    FilterString = FilterString & vbNullChar

    With OFN
        '.hwndOwner
        '.hInstance
        .lpstrFilter = FilterString
        '.lpstrCustomFilter
        '.nMaxCustFilter
        .nFilterIndex = 1
        .lpstrFile = FileName & String$(MAX_BUFFER - Len(FileName), vbNullChar)
        .nMaxFile = Len(.lpstrFile)
        .lpstrFileTitle = String$(MAX_BUFFER - 1, " ") & vbNullChar
        .nMaxFileTitle = Len(.lpstrFileTitle)
        .lpstrInitialDir = "C:\" & vbNullChar
        .lpstrTitle = Title
        .flags = OFN_FILEMUSTEXIST Or OFN_PATHMUSTEXIST
        '.nFileOffset
        '.nFileExtension
        '.lpstrDefExt
        '.lCustData
        '.lpfnHook
        '.lpTemplateName
        '.pvReserved
        '.dwReserved
        '.FlagsEx
        .lStructSize = LenB(OFN)
    End With

    Dim Result As Long
    Result = GetSaveFileName(OFN)
    If Result <> 0 Then
        GetSaveFileNameWinAPI = Trim$(Replace(OFN.lpstrFile, vbNullChar, " "))
    End If
End Function

Public Function ToUTF8(Text As String) As Byte()
    Dim lngResult As Long
    Dim UTF8() As Byte

    lngResult = WideCharToMultiByte(CP_UTF8, 0, StrPtr(Text), Len(Text), 0, 0, 0, 0)
    If lngResult > 0 Then
        ReDim UTF8(lngResult - 1)
        WideCharToMultiByte CP_UTF8, 0, StrPtr(Text), Len(Text), VarPtr(UTF8(0)), lngResult, 0, 0
    End If
    ToUTF8 = UTF8
End Function

Private Function BytesLength(abBytes() As Byte) As Long
    ' Trap error if array is uninitialized
    On Error Resume Next
    BytesLength = UBound(abBytes) - LBound(abBytes) + 1
End Function

Public Function FromUTF8(abUtf8Array() As Byte) As String
    Dim nBytes As Long
    Dim nChars As Long
    Dim strOut As String
    Utf8BytesToString = ""
    ' Catch uninitialized input array
    nBytes = BytesLength(abUtf8Array)
    If nBytes <= 0 Then Exit Function
    ' Get number of characters in output string
    nChars = MultiByteToWideChar(CP_UTF8, 0&, VarPtr(abUtf8Array(0)), nBytes, 0&, 0&)
    ' Dimension output buffer to receive string
    strOut = String(nChars, 0)
    nChars = MultiByteToWideChar(CP_UTF8, 0&, VarPtr(abUtf8Array(0)), nBytes, StrPtr(strOut), nChars)
    FromUTF8 = Left$(strOut, nChars)
End Function

Public Sub WriteUTF8File(FileName As String, Text As String)
    Dim hf As Integer: hf = FreeFile
    Open FileName For Binary Access Write As #hf
        Put #hf, , ToUTF8(Text)
    Close #hf
End Sub

Public Function ReadUTF8File(FileName As String) As String
    Dim hf As Integer: hf = FreeFile
    Dim utf8Bytes() As Byte
    Open FileName For Binary Access Read As #hf
        utf8Bytes = InputB(LOF(hf), #hf)
    Close #hf
    ReadUTF8File = FromUTF8(utf8Bytes)
End Function

