VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Orbital"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Implements IOrbital
Private m_Spin    As ESpin

Friend Sub New_(aESpin As ESpin)
    m_Spin = aESpin
End Sub

Public Property Get ESpin() As ESpin
    ESpin = m_Spin
End Property

Public Function ESpinInc() As Boolean
    If ESpin.SpinUpDown <= m_Spin Then Exit Function
    If ESpin.none <= m_Spin And m_Spin < ESpin.SpinUpDown Then m_Spin = m_Spin + 1
    ESpinInc = True
End Function

Public Function ESpinDec() As Boolean
    If m_Spin <= ESpin.none Then Exit Function
    If m_Spin <= ESpin.SpinUpDown Then m_Spin = m_Spin - 1
    ESpinDec = True
End Function

Public Property Get CountElektrons() As Long
    CountElektrons = CLng(m_Spin)
End Property
'
'Private Property Get IOrbital_Count() As Long
'    IOrbital_Count = CLng(m_Spin)
'End Property
'
'Private Property Get IOrbital_EnergyNiveau() As Long
'
'End Property
