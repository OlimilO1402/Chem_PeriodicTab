VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Atom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_ChemElem As ChemElement
Private m_Periods  As ListOfPeriode

Friend Sub New_(ByVal iOrd As Long)
    m_ChemElem = MChemElem.GetChemElemFromOrd(iOrd)
    Set m_Periods = MPeriods.GetPeriods(iOrd)
End Sub

Friend Function ToStr() As String
    Dim s As String
    With m_ChemElem
        s = s & PadLeft(CStr(.Ordnungszahl), 3) & " "
        s = s & PadRight(.Symbol, 2) & " "
        s = s & PadRight(.Name, 14) & " "
        s = s & PadLeft(Format(.Atomgewicht, "0.000"), 7) & "; "
        s = s & "is R.Act: " & PadRight(BoolToYesNo(.isRadioaktiv), 6) & "; "
        s = s & "is Artif: " & PadRight(BoolToYesNo(.isArtificial), 6) & "; "
        s = s & PadRight(ESerie_ToStr(.Serie), 15) & "; "
        s = s & PadRight(EStoffTyp_ToStr(.StoffTyp), 11) & "; "
        s = s & "elneg: " & Format(.ElNegativ, "0.0") & "; "
        s = s & "n-Neutr: " & PadLeft(Me.nNeutrons, 3) & "; "
        s = s & "ElKonf: " & ElKonf_ToStr(.ElKonfig)
    End With
    ToStr = s
End Function

Public Function CountElectrons() As Long
    Dim i As Long, c As Long
    For i = 0 To UBound(m_ChemElem.ElKonfig)
        c = c + m_ChemElem.ElKonfig(i)
    Next
    CountElectrons = c
End Function

Public Property Get nProtons() As Byte
    nProtons = m_ChemElem.Ordnungszahl
End Property

Public Property Get nNeutrons() As Byte
    nNeutrons = m_ChemElem.Atomgewicht - Me.nProtons
End Property

