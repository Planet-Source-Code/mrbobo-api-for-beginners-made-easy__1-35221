Attribute VB_Name = "ModInfo"
Private Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Private Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Private Declare Function SHGetDiskFreeSpace Lib "shell32" Alias "SHGetDiskFreeSpaceA" (ByVal pszVolume As String, pqwFreeCaller As Currency, pqwTot As Currency, pqwFree As Currency) As Long
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Sub GetHD(mDLetter As String, ByRef mSerial As Long, ByRef mName As String, ByRef fSys As String, ByRef mTotal As String, ByRef mFree As String)
    Dim FC As Currency, TSp As Currency, FSp As Currency
    mName = String$(255, Chr$(0))
    fSys = String$(255, Chr$(0))
    GetVolumeInformation mDLetter, mName, 255, mSerial, 0, 0, fSys, 255
    mName = Left$(mName, InStr(1, mName, Chr$(0)) - 1)
    fSys = Left$(fSys, InStr(1, fSys, Chr$(0)) - 1)
    SHGetDiskFreeSpace mDLetter, FC, TSp, FSp
    mTotal = Format$(TSp * 10000, "###,###,###,##0")
    mFree = Format$(FSp * 10000, "###,###,###,##0")
End Sub
Public Function GetFixedDisks() As Variant
    Dim strReturn As String, temp As String, z As Long, mDStr() As String, cnt As Long
    strReturn = String(255, Chr$(0))
    Ret& = GetLogicalDriveStrings(255, strReturn)
    For z = 1 To 100
        If Left$(strReturn, InStr(1, strReturn, Chr$(0))) = Chr$(0) Then Exit For
        temp = Left$(strReturn, InStr(1, strReturn, Chr$(0)) - 1)
        If GetDriveType(temp) <> 2 And GetDriveType(temp) <> 5 Then
            ReDim Preserve mDStr(cnt)
            mDStr(cnt) = temp
            cnt = cnt + 1
        End If
        strReturn = Right$(strReturn, Len(strReturn) - InStr(1, strReturn, Chr$(0)))
    Next z
    GetFixedDisks = mDStr
End Function

Public Function UserName() As String
    Dim temp As String
    temp = String(100, Chr$(0))
    GetUserName temp, 100
    UserName = Left$(temp, InStr(temp, Chr$(0)) - 1)
End Function
