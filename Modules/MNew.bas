Attribute VB_Name = "MNew"
Option Explicit
Public Enum ESpin
    None = 0
    SpinUp = 1
    SpinUpDown = 2
End Enum

Public Function Atom(ByVal iOrd As Byte) As Atom
    Set Atom = New Atom: Atom.New_ iOrd
End Function

Public Function Orbital(aESpin As ESpin) As Orbital
    Set Orbital = New Orbital: Orbital.New_ aESpin
End Function

'Public Function Orbital(ByRef iElekLeft_inout As Long) As Orbital
'    Set Orbital = New Orbital
'    'Ohje, das ist ja eigentlich totaler Müll so
'    '
'    Dim e As ESpin
'    If iElekLeft_inout > 0 Then
'        e = ESpin.SpinUp
'        iElekLeft_inout = iElekLeft_inout - 1
'    End If
'    If iElekLeft_inout > 0 Then
'        e = ESpin.SpinUpDown
'        iElekLeft_inout = iElekLeft_inout - 1
'    End If
'    Orbital.New_ e
'End Function

Public Function OrbitalBase(ByVal Size As EOrbitalSize) As OrbitalBase
    Set OrbitalBase = New OrbitalBase
    OrbitalBase.New_ Size
End Function

'Public Function OrbitalS(ByVal aENiveau As Long, aOrbital As Orbital) As OrbitalS
Public Function OrbitalS(ByVal aENiveau As Long) As OrbitalS
    Set OrbitalS = New OrbitalS
    OrbitalS.New_ aENiveau ', aOrbital
End Function

'Public Function OrbitalP(ByVal aENiveau As Long, aOPx As Orbital, aOPy As Orbital, aOPz As Orbital) As OrbitalP
Public Function OrbitalP(ByVal aENiveau As Long) As OrbitalP
    Set OrbitalP = New OrbitalP
    OrbitalP.New_ aENiveau ', aOPx, aOPy, aOPz
End Function

'Public Function OrbitalD(ByVal aENiveau As Long, aO1 As Orbital, aO2 As Orbital, aO3 As Orbital, aO4 As Orbital, aO5 As Orbital) As OrbitalD
Public Function OrbitalD(ByVal aENiveau As Long) As OrbitalD
    Set OrbitalD = New OrbitalD
    OrbitalD.New_ aENiveau ', aO1, aO2, aO3, aO4, aO5
End Function

'Public Function OrbitalF(ByVal aENiveau As Long, aO1 As Orbital, aO2 As Orbital, aO3 As Orbital, aO4 As Orbital, aO5 As Orbital, aO6 As Orbital, aO7 As Orbital) As OrbitalF
Public Function OrbitalF(ByVal aENiveau As Long) As OrbitalF
    Set OrbitalF = New OrbitalF
    OrbitalF.New_ aENiveau ', aO1, aO2, aO3, aO4, aO5, aO6, aO7
End Function

Public Function Periode(ByVal aVal As Long, OS As OrbitalS, Optional OP As OrbitalP, Optional OD As OrbitalD, Optional OF As OrbitalF) As Periode
    Set Periode = New Periode
    Periode.New_ aVal, OS, OP, OD, OF
End Function

Public Function Min(V1, V2)
    If V1 < V2 Then Min = V1 Else Min = V2
End Function
Public Function Max(V1, V2)
    If V1 > V2 Then Max = V1 Else Max = V2
End Function

