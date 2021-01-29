VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ListOfAtom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Arr() As Atom
Private m_Count As Long

Private Sub Class_Initialize()
    m_Count = 118
    ReDim m_Arr(0 To m_Count)
    Dim i As Long
    For i = 1 To m_Count
        Set m_Arr(i) = MNew.Atom(i)
    Next
End Sub

Public Property Get Item(ByVal Index As Long) As Atom
    Set Item = m_Arr(Index)
End Property

Public Sub ToListBox(aLB As ListBox)
    aLB.Clear
    Dim i As Long
    For i = 1 To m_Count
        aLB.AddItem m_Arr(i).ToStr
    Next
End Sub
