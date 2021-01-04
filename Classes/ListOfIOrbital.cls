VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ListOfIOrbital"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_Arr() As IOrbital
Private m_Count As Long

Private Sub Class_Initialize()
    ReDim m_Arr(0 To 3)
End Sub

Public Sub Add(aObj As IOrbital)
    If UBound(m_Arr) < m_Count Then ReDim Preserve m_Arr(0 To 2 * m_Count - 1)
    'Set Add = aObj:
    Set m_Arr(m_Count) = aObj: m_Count = m_Count + 1
End Sub

Public Property Get Count() As Long
    Count = m_Count
End Property

Public Property Get Item(ByVal Index As Long) As IOrbital
    Index = Index - 1
    If 0 <= Index And Index < m_Count Then Set Item = m_Arr(Index)
End Property

Friend Function ToStr() As String
    Dim i As Long
    Dim o As IOrbital
    Dim s As String
    For i = 0 To m_Count - 1
        Set o = m_Arr(i)
        s = s & o.ToStr & ", "
    Next
    ToStr = s
End Function
