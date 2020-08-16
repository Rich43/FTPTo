VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEncryption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
    '**************************************'
    '*   Copyright © T-RonX Modding 2005  *'
    '* Do only use if you have premission *'
    '*     Ask me: t_ronx@homail.com      *'
    '**************************************'

Function StrEncrypt(Str As String) As String
On Error GoTo StrError

Dim NewStr As String 'Holds the new string.
Dim ChrAsc As Long   'Holds ascii number of char.

    For i = 1 To Len(Str)
        'Convert char to ascii number.
        ChrAsc = CInt(Asc(Mid(Str, i, 1)))
        'Return new char by converting ascii code to Unicode or ANSI char. Last chat will be the first one in encrypted string.
        NewStr = ChrW((ChrAsc * 2)) + NewStr
    Next

    'Final char
    StrEncrypt = NewStr

Exit Function
StrError:
Call StrError

End Function

Function StrDecrypt(Str As String) As String
On Error GoTo StrError

Dim NewStr As String 'Holds the new string.
Dim ChrAns As Long   'Holds ANSI number of char.

    For i = 1 To Len(Str)
        'Convert char to ANSI number.
        ChrAns = CInt(AscW(Mid(Str, i, 1)))
        'Convert ANSI num to ascii char.
        NewStr = Chr((ChrAns / 2)) + NewStr
    Next

    'Final char
    StrDecrypt = NewStr
    
Exit Function
StrError:
Call StrError

End Function

Private Function StrError()
    MsgBox "An unknow string encryption error occurred.", vbCritical
End Function
