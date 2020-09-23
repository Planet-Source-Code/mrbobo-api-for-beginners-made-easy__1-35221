Attribute VB_Name = "ModIEversion"
'Get IE version
Private Declare Function DllGetVersion Lib "Shlwapi.dll" (dwVersion As DllVersionInfo) As Long
Private Type DllVersionInfo
    cbSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
End Type

Public Function Get_ExplorerVersion() As String
    Dim VersionInfo As DllVersionInfo
    VersionInfo.cbSize = Len(VersionInfo)
    Call DllGetVersion(VersionInfo)
    Select Case VersionInfo.dwMajorVersion
        Case 4
            Select Case VersionInfo.dwMinorVersion
                Case 40
                    Select Case VersionInfo.dwBuildNumber
                        Case 308
                            Get_ExplorerVersion = "Internet Explorer 1.0 (Plus!)"
                        Case 520
                            Get_ExplorerVersion = "Internet Explorer 2.0"
                    End Select
                Case 70
                    Select Case VersionInfo.dwBuildNumber
                        Case 1155
                            Get_ExplorerVersion = "Internet Explorer 3.0"
                        Case 1158
                            Get_ExplorerVersion = "Internet Explorer 3.0 (OSR2)"
                        Case 1215
                            Get_ExplorerVersion = "Internet Explorer 3.01"
                        Case 1300
                            Get_ExplorerVersion = "Internet Explorer 3.02 and 3.02a"
                    End Select
                Case 71
                    Select Case VersionInfo.dwBuildNumber
                        Case 544
                            Get_ExplorerVersion = "Internet Explorer 4.0 Platform Preview 1.0 (PP1)"
                        Case 1008
                            Get_ExplorerVersion = "Internet Explorer 4.0 Platform Preview 2.0 (PP2)"
                        Case 1712
                            Get_ExplorerVersion = "Internet Explorer 4.0"
                        Case 2106
                            Get_ExplorerVersion = "Internet Explorer 4.01"
                        Case 3110
                            Get_ExplorerVersion = "Internet Explorer 4.01 Service Pack 1 (SP1)"
                        Case 3612
                            Get_ExplorerVersion = "Internet Explorer 4.01 Service Pack 2 (SP2)"
                    End Select
                Case 72
                
            End Select
        Case 5
            Select Case VersionInfo.dwMinorVersion
                Case 0
                    Select Case VersionInfo.dwBuildNumber
                        Case 518
                            Get_ExplorerVersion = "Internet Explorer 5 Developer Preview (Beta 1)"
                        Case 910
                            Get_ExplorerVersion = "Internet Explorer 5 Beta (Beta 2)"
                        Case 2014
                            Get_ExplorerVersion = "Internet Explorer 5"
                        Case 2314
                            Get_ExplorerVersion = "Internet Explorer 5 (Office 2000)"
                        Case 2614
                            Get_ExplorerVersion = "Internet Explorer 5 (Windows 98 Second Edition)"
                        Case 2516
                            Get_ExplorerVersion = "Internet Explorer 5.01 (Windows 2000 Beta 3, build 5.00.2031)"
                        Case 2919.8
                            Get_ExplorerVersion = "Internet Explorer 5.01 (Windows 2000 RC1, build 5.00.2072)"
                        Case 2919.38
                            Get_ExplorerVersion = "Internet Explorer 5.01 (Windows 2000 RC2, build 5.00.2128)"
                        Case 2919.6307
                            Get_ExplorerVersion = "Internet Explorer 5.01 (Also included with Office 2000 SR-1, but not installed by default)"
                        Case 2920
                            Get_ExplorerVersion = "Internet Explorer 5.01 (Windows 2000, build 5.00.2195)"
                        Case 3103
                            Get_ExplorerVersion = "Internet Explorer 5.01 SP1 (Windows 2000)"
                        Case 3105
                            Get_ExplorerVersion = "Internet Explorer 5.01 SP1 (Windows 95/98 and Windows NT 4.0)"
                        Case 3314
                            Get_ExplorerVersion = "Internet Explorer 5.01 SP2 (Windows 95/98 and Windows NT 4.0)"
                        Case 3315
                            Get_ExplorerVersion = "Internet Explorer 5.01 SP2 (Windows 2000)"
                    End Select
                Case 5
                    Select Case VersionInfo.dwBuildNumber
                        Case 3825
                            Get_ExplorerVersion = "Internet Explorer 5.5 Developer Preview (Beta)"
                        Case 4030
                            Get_ExplorerVersion = "Internet Explorer 5.5 & Internet Tools Beta"
                        Case 4134.1
                            Get_ExplorerVersion = "Windows Me (4.90.3000)"
                        Case 4134.6
                            Get_ExplorerVersion = "Internet Explorer 5.5"
                        Case 4308
                            Get_ExplorerVersion = "Internet Explorer 5.5 Advanced Security Privacy Beta"
                        Case 4522
                            Get_ExplorerVersion = "Internet Explorer 5.5 Service Pack 1"
                        Case 4807
                            Get_ExplorerVersion = "Internet Explorer 5.5 Service Pack 2"
                    End Select
            End Select
        Case 6
            Select Case VersionInfo.dwMinorVersion
                Case 2462
                    Get_ExplorerVersion = "Internet Explorer 6 Public Preview (Beta)"
                Case 2479
                    Get_ExplorerVersion = "Internet Explorer 6 Public Preview (Beta) Refresh"
                Case 2600
                    Get_ExplorerVersion = "Internet Explorer 6"
            End Select
    End Select
End Function

Public Function IsIE4orGreater() As Boolean
    Dim VersionInfo As DllVersionInfo
    VersionInfo.cbSize = Len(VersionInfo)
    Call DllGetVersion(VersionInfo)
    Select Case VersionInfo.dwMajorVersion
        Case 4
            Select Case VersionInfo.dwMinorVersion
                Case 40
                    IsIE4orGreater = False
                Case 70
                    IsIE4orGreater = False
                Case 71
                    IsIE4orGreater = True
                Case 72
                    IsIE4orGreater = True
            End Select
        Case 5
            IsIE4orGreater = True
        Case 6
            IsIE4orGreater = True
    End Select
End Function
