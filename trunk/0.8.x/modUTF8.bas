Attribute VB_Name = "modUTF8"
Public Const CP_UTF8 = 65001
 
Public Declare Function MultiByteToWideChar Lib "kernel32" _
(ByVal CodePage As Long, ByVal dwFlags As Long, _
ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, _
ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
 
Public Function UTFOpen(FileNameUTF As String) As String
    Dim utf8() As Byte
    Dim ucs2 As String
    Dim chars As Long
    
    Open FileNameUTF For Binary As #1   'UTF-8 문서지정
    ReDim utf8(LOF(1))
    
    Get #1, , utf8
    
    chars = MultiByteToWideChar(CP_UTF8, 0, VarPtr(utf8(0)), LOF(1), 0, 0)
    ucs2 = Space(chars)
    
    chars = MultiByteToWideChar(CP_UTF8, 0, VarPtr(utf8(0)), LOF(1), StrPtr(ucs2), chars)
    
    UTFOpen = ucs2
    Close
End Function

