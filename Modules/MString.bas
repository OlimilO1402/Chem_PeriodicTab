Attribute VB_Name = "MString"
Option Explicit

' String
Public Function PadLeft(StrVal As String, _
                        ByVal totalWidth As Long, _
                        Optional ByVal paddingChar As String) As String
    ' der String wird mit der angegebenen Länge zurückgegeben, der
    ' String wird nach rechts gerückt, und links mit PadChar aufgefüllt
    ' ist PadChar nicht angegeben, so wird mit RSet der String in
    ' Spaces eingefügt.
    If Len(paddingChar) Then
        If Len(StrVal) <= totalWidth Then PadLeft = String$(totalWidth - Len(StrVal), paddingChar) & StrVal
    Else
        PadLeft = Space$(totalWidth)
        RSet PadLeft = StrVal
    End If
End Function

Public Function PadRight(StrVal As String, _
                         ByVal totalWidth As Long, _
                         Optional ByVal paddingChar As String) As String
    ' der String wird mit der angegebenen Länge zurückgegeben, der
    ' String wird nach links gerückt, und rechts mit PadChar aufgefüllt
    ' ist PadChar nicht angegeben, so wird mit LSet der String in
    ' Spaces eingefügt.
    If Len(paddingChar) Then
        If Len(StrVal) <= totalWidth Then PadRight = StrVal & String$(totalWidth - Len(StrVal), paddingChar)
    Else
        PadRight = Space$(totalWidth)
        LSet PadRight = StrVal
    End If
End Function

'Bool
Public Function BoolToYesNo(ByVal b As Boolean) As String
    BoolToYesNo = IIf(b, " Ja ", "Nein")
End Function




