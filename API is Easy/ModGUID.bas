Attribute VB_Name = "ModGUID"
Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type
Private Declare Function CoCreateGuid Lib "ole32.dll" (pguid As GUID) As Long
Private Declare Function StringFromGUID2 Lib "ole32.dll" (rguid As Any, ByVal lpstrClsId As Long, ByVal cbMax As Long) As Long
Public Function GUIDGen() As String
    Dim uGUID As GUID
    Dim sGUID As String
    Dim bGUID() As Byte
    Dim lLen As Long
    Dim retval As Long
    lLen = 40
    bGUID = String(lLen, 0)
    CoCreateGuid uGUID
    retval = StringFromGUID2(uGUID, VarPtr(bGUID(0)), lLen)
    sGUID = bGUID
    If (Asc(Mid$(sGUID, retval, 1)) = 0) Then retval = retval - 1
    GUIDGen = Left$(sGUID, retval)
End Function

