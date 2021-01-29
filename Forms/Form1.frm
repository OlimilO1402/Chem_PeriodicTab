VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8415
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   20775
   LinkTopic       =   "Form1"
   ScaleHeight     =   561
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1385
   StartUpPosition =   3  'Windows-Standard
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'Kein
      Height          =   615
      Left            =   120
      ScaleHeight     =   615
      ScaleWidth      =   18855
      TabIndex        =   1
      Top             =   120
      Width           =   18855
      Begin VB.ListBox List2 
         Height          =   510
         Left            =   3240
         Style           =   1  'Kontrollkästchen
         TabIndex        =   2
         Top             =   0
         Width           =   2655
      End
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7485
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   20535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Atoms As ListOfAtom

'OK wo wollen wir eigentlich hin?
'das Ganze soll ein Periodensystem werden mit dem man rechnen kann,
'alle Werte der Atome sollen im Periodensystem angezeigt werden,
'd.h. die Werte sollen beim Erstellen des Periodensystem automatisch berechnet/ermittelt werden
'
'
Private Sub Form_Load()
    InitChemElements
    'Set Atoms = New ListOfAtom
    'Atoms.ToListBox List1
End Sub

Private Sub Form_Resize()
    Dim brdr As Single: brdr = 8
    Dim L As Single: L = List1.Left
    Dim T As Single: T = List1.Top
    Dim W As Single: W = Me.ScaleWidth - L - brdr
    Dim H As Single: H = Me.ScaleHeight - T - brdr
    If W > 0 And H > 0 Then List1.Move L, T, W, H
End Sub

Private Sub List1_Click()
    Dim a As Atom
    Dim i As Long: i = List1.ListIndex + 1
    If i > 0 Then
        Set a = Atoms.Item(i)
        'MsgBox a.CountElectrons
    End If
End Sub
