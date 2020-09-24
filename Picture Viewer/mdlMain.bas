Attribute VB_Name = "mdlMain"
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Global XRes As Integer, YRes As Integer, INCNR As Integer
Function GetToken(ByVal strVal As String, intIndex As Integer, strDelimiter As String) As String
    Dim strSubString() As String
    Dim intIndex2 As Integer
    Dim i As Integer
    Dim intDelimitLen As Integer
    intIndex2 = 1
    i = 0
    intDelimitLen = Len(strDelimiter)
    Do While intIndex2 > 0
        ReDim Preserve strSubString(i + 1)
        intIndex2 = InStr(1, strVal, strDelimiter)
        If intIndex2 > 0 Then strSubString(i) = Mid(strVal, 1, (intIndex2 - 1)): strVal = Mid(strVal, (intIndex2 + intDelimitLen), Len(strVal))
    If intIndex2 < 0 Then strSubString(i) = strVal
    i = i + 1
    Loop
    If intIndex > (i + 1) Or intIndex < 1 Then GetToken = ""
    If intIndex < (i + 1) Or intIndex > 1 Then GetToken = strSubString(intIndex - 1)
End Function


