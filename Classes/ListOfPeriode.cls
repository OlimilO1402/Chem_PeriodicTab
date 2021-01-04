VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ListOfPeriode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_Arr() As Periode
Private m_Count As Long

Private Sub Class_Initialize()
    ReDim m_Arr(0 To 3)
End Sub

Public Function Add(aObj As Periode) As Periode
    If UBound(m_Arr) < m_Count Then ReDim Preserve m_Arr(0 To 2 * m_Count - 1)
    Set Add = aObj
    Set m_Arr(m_Count) = Add
    m_Count = m_Count + 1
End Function

Public Property Get Count() As Long
    Count = m_Count
End Property

Public Property Get CountElektrons() As Long
    Dim p As Periode
    Dim i As Long, c As Long
    For i = 0 To m_Count - 1
        Set p = m_Arr(i)
        c = c + p.CountElektrons
    Next
    CountElektrons = c
End Property

Public Property Get Item(ByVal Index As Long) As Periode
    If Index >= m_Count Then Exit Property
    Set Item = m_Arr(Index)
End Property

Public Property Get Valenz() As Periode
    Set Valenz = m_Arr(m_Count - 1)
End Property

Friend Function ToStr() As String
    Dim p As Periode: Set p = m_Arr(0)
    If p Is Nothing Then Exit Function
    Dim s As String: s = p.ToStr ' p.Value & ": " & CStr(p.CountElektrons)
    Dim i As Long
    For i = 1 To m_Count - 1
        Set p = m_Arr(i)
        s = s & " / " & p.ToStr ' p.Value & ": " & CStr(p.CountElektrons)
    Next
    ToStr = s
End Function
