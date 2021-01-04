VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Periode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_Value    As Long '1-7
Private m_OrbitalS As IOrbital
Private m_OrbitalP As IOrbital 'Optional
Private m_OrbitalD As IOrbital 'Optional
Private m_OrbitalF As IOrbital 'Optional

Friend Sub New_(ByVal aVal As Long, OS As OrbitalS, Optional OP As OrbitalP, Optional OD As OrbitalD, Optional OF As OrbitalF)
    m_Value = aVal: Set m_OrbitalS = OS: Set m_OrbitalP = OP: Set m_OrbitalD = OD: Set m_OrbitalF = OF
End Sub

Public Property Get Value() As Long
    Value = m_Value
End Property

Public Property Get CountElektrons() As Long
    Dim c As Long
    If Not m_OrbitalS Is Nothing Then c = c + m_OrbitalS.Count
    If Not m_OrbitalP Is Nothing Then c = c + m_OrbitalP.Count
    If Not m_OrbitalD Is Nothing Then c = c + m_OrbitalD.Count
    If Not m_OrbitalF Is Nothing Then c = c + m_OrbitalF.Count
    CountElektrons = c
End Property

Friend Function ToStr() As String
    Dim s As String
    's = m_Value
    If Not m_OrbitalS Is Nothing Then s = s & CStr(m_Value) & "s^" & CStr(m_OrbitalS.Count) & ", " 'IIf(Not m_OrbitalP Is Nothing, ", ", "")
    If Not m_OrbitalP Is Nothing Then s = s & CStr(m_Value) & "p^" & CStr(m_OrbitalP.Count) & ", " 'IIf(Not m_OrbitalD Is Nothing, ", ", "")
    If Not m_OrbitalD Is Nothing Then s = s & CStr(m_Value) & "d^" & CStr(m_OrbitalD.Count) & ", " 'IIf(Not m_OrbitalF Is Nothing, ", ", "")
    If Not m_OrbitalF Is Nothing Then s = s & CStr(m_Value) & "f^" & CStr(m_OrbitalF.Count) '& ", "
    ToStr = s
End Function
