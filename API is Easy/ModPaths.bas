Attribute VB_Name = "ModPaths"
Private Type SHITEMID
    cb As Long
    abID As Byte
End Type
Private Type ITEMIDLIST
    mkid As SHITEMID
End Type
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Private Declare Function GetTempPath Lib "Kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetTempFileName Lib "Kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFilename As String) As Long
Private Declare Function GetSystemDirectory Lib "Kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetWindowsDirectory Lib "Kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetShortPathName Lib "Kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Public Declare Function GetLongPathName Lib "Kernel32" Alias "GetLongPathNameA" (ByVal lpszShortPath As String, ByVal lpszLongPath As String, ByVal cchBuffer As Long) As Long
Public Function GetLongFilename(ByVal sShortFilename As String) As String
    Dim lRet As Long
    Dim sLongFilename As String
    sLongFilename = String$(1024, " ")
    lRet = GetLongPathName(sShortFilename, sLongFilename, Len(sLongFilename))
    If lRet > Len(sLongFilename) Then
        sLongFilename = String$(lRet + 1, " ")
        lRet = GetLongPathName(sShortFilename, sLongFilename, Len(sLongFilename))
    End If
    If lRet > 0 Then
        GetLongFilename = Left$(sLongFilename, lRet)
    End If
End Function
Public Function GetDosPath(LongPath As String) As String
    Dim s As String
    Dim i As Long
    Dim PathLength As Long
    i = Len(LongPath) + 1
    s = String(i, 0)
    PathLength = GetShortPathName(LongPath, s, i)
    GetDosPath = Left$(s, PathLength)
End Function
Public Function GetTempPathName() As String
    Dim sBuffer As String
    Dim lRet As Long
    sBuffer = String$(255, vbNullChar)
    lRet = GetTempPath(255, sBuffer)
    If lRet > 0 Then
        sBuffer = Left$(sBuffer, lRet)
    End If
    GetTempPathName = sBuffer
End Function
Public Function GetTempFile(mPrefix As String) As String
    Dim lpTempFilename As String, mdir As String
    lpTempFilename = String(255, vbNullChar)
    mdir = GetTempPathName
    GetTempFileName mdir, mPrefix, 0, lpTempFilename
    GetTempFile = StripTerminator(lpTempFilename)
End Function
Public Function SpecialFolder(ByVal CSIDL As Long) As String
    Dim r As Long
    Dim sPath As String
    Dim IDL As ITEMIDLIST
    Const NOERROR = 0
    Const MAX_LENGTH = 260
    r = SHGetSpecialFolderLocation(Form1.hwnd, CSIDL, IDL)
    If r = NOERROR Then
        sPath = Space$(MAX_LENGTH)
        r = SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal sPath)
        If r Then
            SpecialFolder = Left$(sPath, InStr(sPath, Chr$(0)) - 1)
        End If
    End If
End Function
Public Function Winfolder() As String
    Dim strSave As String
    strSave = String(200, Chr$(0))
    Winfolder = Left$(strSave, GetWindowsDirectory(strSave, Len(strSave)))
End Function
Public Function Sysfolder() As String
    Dim Str As String * 128
    Dim SysWinName As Integer
    SysWinName = GetSystemDirectory(Str, 128)
    Sysfolder = Left(Str, SysWinName)
End Function
Public Function StripTerminator(ByVal strString As String) As String
    'used to remove null characters
    Dim intZeroPos As Integer
    intZeroPos = InStr(strString, Chr$(0))
    If intZeroPos > 0 Then
        StripTerminator = Left$(strString, intZeroPos - 1)
    Else
        StripTerminator = strString
    End If
End Function

