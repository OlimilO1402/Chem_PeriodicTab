VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OrbitalF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IOrbital
'Private m_Periode As Periode
'Private m_ENiveau As EnergyNiveau
Private m_ENiveau As Long
Private m_Base As OrbitalBase
'Private m_O1      As Orbital
'Private m_O2      As Orbital
'Private m_O3      As Orbital
'Private m_O4      As Orbital
'Private m_O5      As Orbital
'Private m_O6      As Orbital
'Private m_O7      As Orbital

Private Sub Class_Initialize()
    Set m_Base = MNew.OrbitalBase(eo_F_Orbital)
End Sub

'Friend Sub New_(aPeriode As Periode, aENiveau As EnergyNiveau, aO1 As Orbital, aO2 As Orbital, aO3 As Orbital, aO4 As Orbital, aO5 As Orbital, aO6 As Orbital, aO7 As Orbital)
Friend Sub New_(ByVal aENiveau As Long) ', aO1 As Orbital, aO2 As Orbital, aO3 As Orbital, aO4 As Orbital, aO5 As Orbital, aO6 As Orbital, aO7 As Orbital)
    'Set m_Periode = aPeriode: Set m_ENiveau = aENiveau
    m_ENiveau = aENiveau
    'Set m_O1 = aO1: Set m_O2 = aO2: Set m_O3 = aO3: Set m_O4 = aO4: Set m_O5 = aO5: Set m_O6 = aO6: Set m_O7 = aO7
End Sub

Private Function IOrbital_COrbitalD() As OrbitalD
End Function

Private Function IOrbital_COrbitalF() As OrbitalF
    Set IOrbital_COrbitalF = Me
End Function

Private Function IOrbital_COrbitalP() As OrbitalP
End Function

Private Function IOrbital_COrbitalS() As OrbitalS
End Function

'Private Property Get IOrbital_Count() As Long
'    Dim c As Long
'    If Not m_O1 Is Nothing Then c = c + m_O1.Count
'    If Not m_O2 Is Nothing Then c = c + m_O2.Count
'    If Not m_O3 Is Nothing Then c = c + m_O3.Count
'    If Not m_O4 Is Nothing Then c = c + m_O4.Count
'    If Not m_O5 Is Nothing Then c = c + m_O5.Count
'    If Not m_O6 Is Nothing Then c = c + m_O6.Count
'    If Not m_O7 Is Nothing Then c = c + m_O7.Count
'    IOrbital_Count = c
'End Property
'
'Private Property Get IOrbital_ENiveau() As EnergyNiveau
'    Set IOrbital_ENiveau = m_ENiveau
'End Property
'
'Private Property Get IOrbital_Periode() As Periode
'    Set IOrbital_Periode = m_Periode
'End Property

Private Property Get IOrbital_CountElektrons() As Long
    '
End Property

Private Property Get IOrbital_EnergyNiveau() As Long
    IOrbital_EnergyNiveau = m_ENiveau
End Property

Private Property Get IOrbital_NextEnergyNiveau() As IOrbital
    '
End Property

Private Property Get IOrbital_PrevEnergyNiveau() As IOrbital
    '
End Property

Private Function IOrbital_ToStr() As String
    Dim s As String
    s = CStr(m_ENiveau) & "f" & "^" & CStr(IOrbital_CountElektrons)
    IOrbital_ToStr = s
End Function

