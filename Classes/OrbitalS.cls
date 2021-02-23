VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OrbitalS"
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
'Private m_Orbital As Orbital
'
'
'
'
'
'

Private Sub Class_Initialize()
    Set m_Base = MNew.OrbitalBase(eo_S_Orbital)
End Sub

'Friend Sub New_(aPeriode As Periode, aENiveau As EnergyNiveau, aOrbital As Orbital)
Friend Sub New_(ByVal aENiveau As Long)
    m_ENiveau = aENiveau
End Sub

Private Function IOrbital_COrbitalD() As OrbitalD
End Function

Private Function IOrbital_COrbitalF() As OrbitalF
End Function

Private Function IOrbital_COrbitalP() As OrbitalP
End Function

Private Function IOrbital_COrbitalS() As OrbitalS
    Set IOrbital_COrbitalS = Me
End Function

Private Property Get IOrbital_CountElektrons() As Long
    IOrbital_CountElektrons = m_Base.CountElektrons
End Property

Private Property Get IOrbital_EnergyNiveau() As Long
    IOrbital_EnergyNiveau = m_ENiveau
End Property

Private Property Get IOrbital_NextEnergyNiveau() As IOrbital

End Property

Private Property Get IOrbital_PrevEnergyNiveau() As IOrbital

End Property

Private Function IOrbital_ToStr() As String
    Dim s As String
    s = CStr(m_ENiveau) & "s" & "^" & CStr(IOrbital_CountElektrons)
    IOrbital_ToStr = s
End Function
