VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IOrbital"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Property Get CountElektrons() As Long
    'gibt die Anzahl der Elektronen in diesem Orbital zurück
End Property

'Warum nochmal um ales in der Welt muss das Orbiotal etwas von seiner Periode wissen?
'Public Property Get Periode() As Periode
'End Property

'Public Property Get ENiveau() As EnergyNiveau
'End Property

Public Property Get EnergyNiveau() As Long
    'gibt das Energie-Niveau in diesem Orbital zurück
End Property

Public Property Get PrevEnergyNiveau() As IOrbital
    'gibt das Orbital mit dem nächst niedrigeren Energieniveau zurück
End Property
Public Property Get NextEnergyNiveau() As IOrbital
    'gibt das Orbital mit dem nächst höheren Energieniveau zurück
End Property

Public Function COrbitalS() As OrbitalS
    
End Function
Public Function COrbitalP() As OrbitalP
    
End Function
Public Function COrbitalD() As OrbitalD
    
End Function
Public Function COrbitalF() As OrbitalF
    
End Function

Public Function ToStr() As String

End Function
