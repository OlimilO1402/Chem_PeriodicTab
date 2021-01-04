Attribute VB_Name = "MString"
Option Explicit

' String
Public Function PadLeft(StrVal As String, _
                        ByVal totalWidth As Long, _
                        Optional ByVal paddingChar As String) As String
    ' der String wird mit der angegebenen L�nge zur�ckgegeben, der
    ' String wird nach rechts ger�ckt, und links mit PadChar aufgef�llt
    ' ist PadChar nicht angegeben, so wird mit RSet der String in
    ' Spaces eingef�gt.
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
    ' der String wird mit der angegebenen L�nge zur�ckgegeben, der
    ' String wird nach links ger�ckt, und rechts mit PadChar aufgef�llt
    ' ist PadChar nicht angegeben, so wird mit LSet der String in
    ' Spaces eingef�gt.
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




