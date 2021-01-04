VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OrbitalP"
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
'Private m_OPx     As Orbital
'Private m_OPy     As Orbital
'Private m_OPz     As Orbital
'
'
'
'

Private Sub Class_Initialize()
    Set m_Base = MNew.OrbitalBase(eo_P_Orbital)
End Sub

'Friend Sub New_(aPeriode As Periode, aENiveau As EnergyNiveau, aOPx As Orbital, aOPy As Orbital, aOPz As Orbital)
Friend Sub New_(ByVal aENiveau As Long) ', aOPx As Orbital, aOPy As Orbital, aOPz As Orbital)
    'Set m_Periode = aPeriode: Set m_ENiveau = aENiveau
    m_ENiveau = aENiveau
    'Set m_OPx = aOPx: Set m_OPy = aOPy: Set m_OPz = aOPz
End Sub

Private Function IOrbital_COrbitalD() As OrbitalD
End Function

Private Function IOrbital_COrbitalF() As OrbitalF
End Function

Private Function IOrbital_COrbitalP() As OrbitalP
    Set IOrbital_COrbitalP = Me
End Function

Private Function IOrbital_COrbitalS() As OrbitalS
End Function

Private Property Get IOrbital_CountElektrons() As Long
'    Dim c As Long
'    If Not m_OPx Is Nothing Then c = c + m_OPx.Count
'    If Not m_OPy Is Nothing Then c = c + m_OPy.Count
'    If Not m_OPz Is Nothing Then c = c + m_OPz.Count
    IOrbital_CountElektrons = m_Base.CountElektrons 'c
End Property
'
'Private Property Get IOrbital_ENiveau() As EnergyNiveau
'    Set IOrbital_ENiveau = m_ENiveau
'End Property
'
'Private Property Get IOrbital_Periode() As Periode
'    Set IOrbital_Periode = m_Periode
'End Property
'
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
    s = CStr(m_ENiveau) & "p" & "^" & CStr(IOrbital_Count)
    IOrbital_ToStr = s
End Function

