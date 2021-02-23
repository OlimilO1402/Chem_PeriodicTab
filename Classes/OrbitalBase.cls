VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OrbitalBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Enum EOrbitalSize
    eo_S_Orbital = 1  ' 2
    eo_P_Orbital = 3  ' 6
    eo_D_Orbital = 5  '10
    eo_F_Orbital = 7  '14
End Enum
Private m_Arr() As Orbital
Private m_Count As Long

Friend Sub New_(ByVal Size As EOrbitalSize)
    m_Count = Size
    Dim i As Long
    ReDim m_Arr(0 To m_Count - 1)
    For i = 0 To UBound(m_Arr)
        Set m_Arr(i) = New Orbital
    Next
End Sub

Public Property Get Count() As Long
    Count = m_Count
End Property

Public Property Get CountElektrons() As Long
    Dim c As Long, i As Long
    For i = 0 To UBound(m_Arr)
        c = c + m_Arr(i).CountElektrons
    Next
End Property

Public Function AddElektrons(ByVal nElektronsLeft As Long) As Long
    If nElektronsLeft = 0 Then Exit Function
    Dim i As Long
    For i = 0 To Min(UBound(m_Arr), nElektronsLeft)
        If m_Arr(i).SpinInc Then nElektronsLeft = nElektronsLeft - 1
        If nElektronsLeft = 0 Then Exit Function
    Next
    For i = 0 To Min(UBound(m_Arr), nElektronsLeft)
        If m_Arr(i).SpinInc Then nElektronsLeft = nElektronsLeft - 1
        If nElektronsLeft = 0 Then Exit Function
    Next
    AddElektrons = nElektronsLeft
End Function

Public Property Get Item(ByVal Index As Long) As Orbital
    If Index < 0 And m_Count <= Index Then Exit Property
    Set Item = m_Arr(Index)
End Property

Public Sub ShareOrbital(aOrbital As Orbital)
    ReDim Preserve m_Arr(0 To m_Count)
    Set m_Arr(m_Count) = aOrbital
    m_Count = m_Count + 1
End Sub
