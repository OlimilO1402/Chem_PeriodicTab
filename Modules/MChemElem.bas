Attribute VB_Name = "MChemElem"
Option Explicit

Public Enum ESerie
    none = 0                ' R ,  G ,  B
    Nichtmetall = &H1       '228, 255, 228
    Edelgas = &H2           '237, 255, 255
    Alkalimetall = &H4      '255, 213, 213
    Erdalkalimetall = &H8   '255, 245, 232
    Halbmetall = &H10       '241, 241, 227
    Halogen = &H20          '255, 255, 227
    Metall = &H40           '241, 241, 241
    Übergangsmetall = &H80  '255, 237, 237
    Lanthanoid = &H100      '255, 237, 255
    Actinoid = &H200        '255, 227, 241
End Enum

Public Enum EStoffTyp 'bei normaler Temperatur 20°C
    Feststoff   'fest
    Gas         'gasförmig
    Flüssigkeit 'flüssig
End Enum

'Bohrsche-Atommodell
'Sommerfeld'sche Atommodell
'Rutherfordsche-Atommodell

'Orbitalmodell nach Pauling, Erwin Schrödinger und Werner Heisenberg
'Schrödingergleichung: Lösung gibt die Bahn von Elektronen an
'Heisenbergsche Unschärferelation:
'Man kann nicht Ort und Impuls gleichzeitig bestimmen
'delta_x_delta_p >= h / (4*Pi)
'
'Energieprinzip
'  Energieärmere Orbitale werden zuerst besetzt
'Pauli-Verbot:
'  jedes Orbital kann max. 2 Elektronen unterschiedlichen Spins aufnehmen
'Pauliprinzip: Ein Orbital wird mit maximal 2 Elektronen besetzt.
'  Diese beiden Elektronen unterscheiden sich in ihrem intrinsichen Drehimpuls, dem Spin
'Hund'sche Regel:
'  Erergiegleiche Orbitale werden erst einzeln, dann unter korrekter Spin-Paarung doppelt besetzt
'
''-Spin 'ESpin
Public Enum OrbitalSpin
    none = 0
    SpinUp = 1
    SpinUpDown = 2
End Enum
'Public Enum ESpin
'    None = 0
'    SpinUp = 1
'    SpinUpDown = 2
'End Enum
'Public Type Orbital
'    E12 As ESpin
'End Type
'Public Type OrbitalS
'    E0102 As ESpin 'Orbital
'End Type
'Public Type OrbitalP
'    E0102 As ESpin 'X
'    E0304 As ESpin 'Y
'    E0506 As ESpin 'Z
'End Type
'Public Type OrbitalD
'    E0102 As ESpin
'    E0304 As ESpin
'    E0506 As ESpin
'    E0708 As ESpin
'    E0910 As ESpin
'End Type
'Public Type OrbitalF
'    E0102 As ESpin
'    E0304 As ESpin
'    E0506 As ESpin
'    E0708 As ESpin
'    E0910 As ESpin
'    E1112 As ESpin
'    E1314 As ESpin
'End Type
'Public Type OrbitalSP
'    OrbS As OrbitalS
'    OrbP As OrbitalP
'End Type
'Public Type OrbitalSPD
'    OrbS As OrbitalS
'    OrbD As OrbitalD
'    OrbP As OrbitalP
'End Type
'Public Type OrbitalSPDF
'    OrbS As OrbitalS
'    OrbP As OrbitalP
'    OrbD As OrbitalD
'    OrbF As OrbitalF
'End Type
'Public Type ElKonfiguration
'    Orb1 As OrbitalS     '1 s-Orbital
'    Orb2 As OritalSP     '
'    Orb3 As OritalSP
'    Orb4 As OrbitalSPD
'    Orb5 As OrbitalSPD
'    Orb6 As OrbitalSPDF
'    Orb7 As OrbitalSPDF
'End Type
Public Type ChemElement
    Ordnungszahl As Long
    Atomgewicht  As Double
    Symbol       As String 'H, He, O, S, Ca, etc
    Name         As String 'Deutscher, englischer oder lateinischer Name
    Serie        As ESerie
    isRadioaktiv As Boolean
    isArtificial As Boolean
    StoffTyp     As EStoffTyp
    'ElKonfig     As ElKonfiguration ' Elektronenkonfiguration
    'ElKonfig()   As Byte            ' Elektronenkonfiguration
    ElKonfig()   As Long             ' Elektronenkonfiguration
    ElNegativ    As Double    ' Elektronegativität
End Type

Private m_ChemElements() As ChemElement


Public Sub InitChemElements()
    ReDim m_ChemElements(0 To 118)
    CreateChemElements
End Sub

Public Sub ChemElements_ToListbox(aLB As ListBox)
    With aLB
        Dim i As Long
        'For i = LBound(m_ChemElements) To UBound(m_ChemElements)
        For i = 1 To UBound(m_ChemElements)
            .AddItem ChemElement_ToStr(m_ChemElements(i))
        Next
    End With
End Sub

Public Function ChemElement_ToStr(this As ChemElement) As String
    Dim s As String: s = "Atom{"
    With this
        s = s & "Ordz: " & PadLeft(CStr(.Ordnungszahl), 3) & "; "
        s = s & "Symb: " & PadRight(.Symbol, 2) & "; "
        s = s & "Name: " & PadRight(.Name, 14) & "; "
        s = s & "AtomW: " & PadLeft(Format(.Atomgewicht, "0.0000"), 7) & "; "
        s = s & "isRAct: " & PadRight(BoolToYesNo(.isRadioaktiv), 6) & "; "
        s = s & "isArtif: " & PadRight(BoolToYesNo(.isArtificial), 6) & "; "
        s = s & "Serie: " & PadRight(ESerie_ToStr(.Serie), 15) & "; "
        s = s & "Stoff: " & PadRight(EStoffTyp_ToStr(.StoffTyp), 11) & "; "
        s = s & "elneg: " & Format(CStr(.ElNegativ), "0.00") & "; "
        's = s & "n-Neutr: " & PadLeft(.nNeutrons, 3) & "; "
        s = s & "ElKonf:  " & ElKonf_ToStr(.ElKonfig) & ";"
    End With
    ChemElement_ToStr = s & "}"
End Function

Public Function GetChemElemFromOrd(ByVal iOrd As Long) As ChemElement
    GetChemElemFromOrd = m_ChemElements(iOrd)
End Function

Private Function GetESerieFromOrd(ByVal iOrd As Long) As ESerie
    Dim e As ESerie
    Select Case iOrd
    Case 0:                                                e = ESerie.none
    Case 1, 6, 7, 8, 15, 16, 34:                           e = ESerie.Nichtmetall
    Case 2, 10, 18, 36, 54, 86:                            e = ESerie.Edelgas
    Case 3, 11, 19, 37, 55, 87:                            e = ESerie.Alkalimetall
    Case 4, 12, 20, 38, 56, 88:                            e = ESerie.Erdalkalimetall
    Case 5, 14, 32, 33, 34, 51, 52, 84, 85:                e = ESerie.Halbmetall
    Case 9, 17, 35, 53, 85:                                e = ESerie.Halogen
    Case 13, 31, 32, 49 To 51, 81 To 84, 113 To 118:       e = ESerie.Metall
    Case 21 To 30, 39 To 48, 57, 72 To 80, 89, 104 To 112: e = ESerie.Übergangsmetall
    Case 58 To 71:                                         e = ESerie.Lanthanoid
    Case 90 To 103:                                        e = ESerie.Actinoid
    End Select
    GetESerieFromOrd = e
End Function

Public Function ESerie_ToStr(e As ESerie) As String
    Dim s As String
    Select Case e
    Case ESerie.none:            s = "None"
    Case ESerie.Nichtmetall:     s = "Nichtmetall"
    Case ESerie.Edelgas:         s = "Edelgas"
    Case ESerie.Alkalimetall:    s = "Alkalimetall"
    Case ESerie.Erdalkalimetall: s = "Erdalkalimetall"
    Case ESerie.Halbmetall:      s = "Halbmetall"
    Case ESerie.Halogen:         s = "Halogen"
    Case ESerie.Metall:          s = "Metall"
    Case ESerie.Übergangsmetall: s = "Übergangsmetall"
    Case ESerie.Lanthanoid:      s = "Lanthanoid"
    Case ESerie.Actinoid:        s = "Actinoid"
    End Select
    ESerie_ToStr = s
End Function

Private Function GetRadioactivFromOrd(ByVal iOrd As Long) As Boolean
    Select Case iOrd
    Case 43, 61, 83 To 118: GetRadioactivFromOrd = True
    End Select
End Function

Private Function GetArtificialFromOrd(ByVal iOrd As Long) As Boolean
    Select Case iOrd
    Case 61, 93 To 118: GetArtificialFromOrd = True
    End Select
End Function

Private Function GetEStoffTypFromOrd(ByVal iOrd As Long) As EStoffTyp
    Dim e As EStoffTyp
    Select Case iOrd
    Case 2, 7 To 10, 17, 18, 36, 54, 86, 118: e = Gas
    Case 35, 80, 112:                e = Flüssigkeit
    Case Else:                       e = Feststoff
    End Select
    GetEStoffTypFromOrd = e
End Function

Public Function EStoffTyp_ToStr(ByVal e As EStoffTyp) As String
    Dim s As String
    Select Case e
    Case Feststoff:   s = "Feststoff"
    Case Flüssigkeit: s = "Flüssigkeit"
    Case Gas:         s = "Gas"
    End Select
    EStoffTyp_ToStr = s
End Function

Public Function ElKonf_ToStr(elkonf() As Long) As String
    Dim i As Long
    Dim s As String: s = elkonf(i)
    Dim u As Long: u = UBound(elkonf)
    If u > 0 Then
        For i = 1 To u
            s = s & "/" & elkonf(i)
        Next
    End If
    ElKonf_ToStr = s
End Function

Private Function ESpin_Inc(ByRef iElLeft_inout As Long, ByVal e As ESpin) As ESpin
    ESpin_Inc = CLng(e) + 1
    iElLeft_inout = iElLeft_inout - 1
End Function

Private Function ESpin_Dec(ByRef iElLeft_inout As Long, ByVal e As ESpin) As ESpin
    ESpin_Dec = CLng(e) - 1
    iElLeft_inout = iElLeft_inout + 1
End Function

'Private Function GetOrbitalS(ByRef iElLeft_inout As Long) As OrbitalS
'    Dim i As Long: i = iElLeft_inout
'    With GetOrbitalS
'        If i > 0 Then .E0102 = ESpin_Inc(i, .E0102)
'        If i > 0 Then .E0102 = ESpin_Inc(i, .E0102)
'    End With
'    iElLeft_inout = i
'End Function
'
'Private Function GetOrbitalP(ByRef iElLeft_inout As Long) As OrbitalP
'    Dim i As Long: i = iElLeft_inout
'    With GetOrbitalP
'        If i > 0 Then .E0102 = ESpin_Inc(i, .E0102)
'        If i > 0 Then .E0304 = ESpin_Inc(i, .E0304)
'        If i > 0 Then .E0506 = ESpin_Inc(i, .E0506)
'        If i > 0 Then .E0102 = ESpin_Inc(i, .E0102)
'        If i > 0 Then .E0304 = ESpin_Inc(i, .E0304)
'        If i > 0 Then .E0506 = ESpin_Inc(i, .E0506)
'    End With
'    iElLeft_inout = i
'End Function
'
'Private Function GetOrbitalD(ByRef iElLeft_inout As Long) As OrbitalD
'    Dim i As Long: i = iElLeft_inout
'    With GetOrbitalD
'        If i > 0 Then .E0102 = ESpin_Inc(i, .E0102)
'        If i > 0 Then .E0304 = ESpin_Inc(i, .E0304)
'        If i > 0 Then .E0506 = ESpin_Inc(i, .E0506)
'        If i > 0 Then .E0102 = ESpin_Inc(i, .E0102)
'        If i > 0 Then .E0304 = ESpin_Inc(i, .E0304)
'        If i > 0 Then .E0506 = ESpin_Inc(i, .E0506)
'    End With
'    iElLeft_inout = i
'End Function
'
'Private Function GetElKonfig(ByVal iORd As Long) As ElKonfiguration
'    Dim ek As ElKonfiguration
'    Dim i As Long: i = iORd
'    'For i = iORd To 1 Step -1
'    Do While i > 0
'        With ek
'            .Orb1 = GetOrbitalS(i)
'        End With
'    Loop
'    GetElKonfig = ek
'End Function

Private Function New_ChemElement(ByVal iOrd As Long, ByVal aSymbol As String, ByVal AtomWeight As String, ByVal elneg As String, ByVal aName As String, ParamArray elkonf()) As ChemElement
    Dim d As Double
    With New_ChemElement
        .Ordnungszahl = iOrd
        If Double_TryParse(AtomWeight, d) Then
            .Atomgewicht = d
        End If
        .Symbol = Trim$(aSymbol)
        .Name = Trim$(aName)
        .Serie = GetESerieFromOrd(.Ordnungszahl)
        .isRadioaktiv = GetRadioactivFromOrd(.Ordnungszahl)
        .isArtificial = GetArtificialFromOrd(.Ordnungszahl)
        .StoffTyp = GetEStoffTypFromOrd(.Ordnungszahl)
        If Double_TryParse(elneg, d) Then
            .ElNegativ = d
        End If
        '.ElKonfig = GetElKonfig(.Ordnungszahl)
        ReDim .ElKonfig(0 To UBound(elkonf))
        Dim i As Long
        For i = 0 To UBound(elkonf)
            .ElKonfig(i) = elkonf(i)
        Next
    End With
End Function

'Private Function GetElKonfig(ByVal iORd As Long) As Long()
'    Dim b() As Long
'    Select Case iORd
'    Case Is <= 2:  ReDim b(0):      b(0) = iORd
'    Case Is <= 10: ReDim b(0 To 1): b(0) = 2:    b(1) = iORd - 2
'    Case Is <= 18: ReDim b(0 To 2): b(0) = 2:    b(1) = 8:        b(2) = iORd - 10
'    Case Is <= 36: ReDim b(0 To 3): b(0) = 2:    b(1) = 8:        b(2) = 8:        b(3) = iORd
'    ' . . .
'    End Select
'    GetElKonfig = b
'End Function

Private Sub AddChemElement(ByRef iOrd_inout As Long, aChemElement As ChemElement)
    Dim i As Long: i = iOrd_inout
    m_ChemElements(i) = aChemElement
    iOrd_inout = iOrd_inout + 1
End Sub

'zur Elektronenkonfiguration
'maximale Anzahl der Schale/Periode
'[A] | [B] | [C=WENN(B<5;B;9-B)] | [D=2*C^2] |
' K  |  1  |  1                  |   2       |
' L  |  2  |  2                  |   8       |
' M  |  3  |  3                  |  18       |
' N  |  4  |  4                  |  32       |
' O  |  5  |  4                  |  32       |
' P  |  6  |  3                  |  18       |
' Q  |  7  |  2                  |   8       |

' 1  2  3  4  5  6  7
'
' 2  8 18 32 32 18  8
'=

Private Sub CreateChemElements()
    Dim i As Long: i = 1
     
    AddChemElement i, New_ChemElement(i, "H ", "1.007940", "2.02", "Wasserstoff  ", 1)                        '  1
    AddChemElement i, New_ChemElement(i, "He", "4.002602", "0.00", "Helium       ", 2)                        '  2
    
    AddChemElement i, New_ChemElement(i, "Li", "6.941000", "0.98", "Lithium      ", 2, 1)                     '  3
    AddChemElement i, New_ChemElement(i, "Be", "9.012182", "1.57", "Beryllium    ", 2, 2)                     '  4
    AddChemElement i, New_ChemElement(i, "B ", "10.81100", "2.04", "Bor          ", 2, 3)                     '  5
    AddChemElement i, New_ChemElement(i, "C ", "12.01070", "2.55", "Kohlenstoff  ", 2, 4)                     '  6
    AddChemElement i, New_ChemElement(i, "N ", "14.00674", "3.04", "Stickstoff   ", 2, 5)                     '  7
    AddChemElement i, New_ChemElement(i, "O ", "15.99400", "3.44", "Sauerstoff   ", 2, 6)                     '  8
    AddChemElement i, New_ChemElement(i, "F ", "18.99840", "3.98", "Fluor        ", 2, 7)                     '  9
    AddChemElement i, New_ChemElement(i, "Ne", "20.17970", "0.00", "Neon         ", 2, 8)                     ' 10
    
    AddChemElement i, New_ChemElement(i, "Na", "22.98977", "0.93", "Natrium      ", 2, 8, 1)                  ' 11
    AddChemElement i, New_ChemElement(i, "Mg", "24.30500", "1.31", "Magnesium    ", 2, 8, 2)                  ' 12
    AddChemElement i, New_ChemElement(i, "Al", "26.981538", "1.61", "Aluminium    ", 2, 8, 3)                 ' 13
    AddChemElement i, New_ChemElement(i, "Si", "28.085500", "1.90", "Silicium     ", 2, 8, 4)                 ' 14
    AddChemElement i, New_ChemElement(i, "P ", "30.973761", "2.19", "Phosphor     ", 2, 8, 5)                 ' 15
    AddChemElement i, New_ChemElement(i, "S ", "32.066000", "2.58", "Schwefel     ", 2, 8, 6)                 ' 16
    AddChemElement i, New_ChemElement(i, "Cl", "35.452700", "3.16", "Chlor        ", 2, 8, 7)                 ' 17
    AddChemElement i, New_ChemElement(i, "Ar", "39.948000", "0.82", "Argon        ", 2, 8, 8)                 ' 18
    
    AddChemElement i, New_ChemElement(i, "K ", "39.098300", "0.82", "Kalium       ", 2, 8, 8, 1)              ' 19
    AddChemElement i, New_ChemElement(i, "Ca", "40.078000", "1.00", "Calcium      ", 2, 8, 8, 2)              ' 20
    AddChemElement i, New_ChemElement(i, "Sc", "44.955910", "1.36", "Scandium     ", 2, 8, 9, 2)              ' 21
    AddChemElement i, New_ChemElement(i, "Ti", "47.867000", "1.54", "Titan        ", 2, 8, 10, 2)             ' 22
    AddChemElement i, New_ChemElement(i, "V ", "50.941500", "1.63", "Vanadium     ", 2, 8, 11, 2)             ' 23
    AddChemElement i, New_ChemElement(i, "Cr", "51.996100", "1.66", "Chrom        ", 2, 8, 13, 1)             ' 24
    AddChemElement i, New_ChemElement(i, "Mn", "54.938049", "1.55", "Mangan       ", 2, 8, 13, 2)             ' 25
    AddChemElement i, New_ChemElement(i, "Fe", "55.845000", "1.83", "Eisen        ", 2, 8, 14, 2)             ' 26
    AddChemElement i, New_ChemElement(i, "Co", "58.933200", "1.91", "Cobalt       ", 2, 8, 15, 2)             ' 27
    AddChemElement i, New_ChemElement(i, "Ni", "58.693400", "1.88", "Nickel       ", 2, 8, 16, 2)             ' 28
    AddChemElement i, New_ChemElement(i, "Cu", "63.546000", "1.90", "Kupfer       ", 2, 8, 18, 1)             ' 29
    AddChemElement i, New_ChemElement(i, "Zn", "65.389000", "1.65", "Zink         ", 2, 8, 18, 2)             ' 30
    AddChemElement i, New_ChemElement(i, "Ga", "69.723000", "1.81", "Gallium      ", 2, 8, 18, 3)             ' 31
    AddChemElement i, New_ChemElement(i, "Ge", "72.641000", "2.01", "Germanium    ", 2, 8, 18, 4)             ' 32
    AddChemElement i, New_ChemElement(i, "As", "74.921600", "2.18", "Arsen        ", 2, 8, 18, 5)             ' 33
    AddChemElement i, New_ChemElement(i, "Se", "78.960000", "2.55", "Selen        ", 2, 8, 18, 6)             ' 34
    AddChemElement i, New_ChemElement(i, "Br", "79.904000", "2.96", "Brom         ", 2, 8, 18, 7)             ' 35
    AddChemElement i, New_ChemElement(i, "Kr", "83.798000", "0.00", "Krypton      ", 2, 8, 18, 8)             ' 36
    
    AddChemElement i, New_ChemElement(i, "Rb", "85.467800", "0.82", "Rubidium     ", 2, 8, 18, 8, 1)          ' 37
    AddChemElement i, New_ChemElement(i, "Sr", "87.620000", "0.95", "Strontium    ", 2, 8, 18, 8, 2)          ' 38
    AddChemElement i, New_ChemElement(i, "Y ", "88.905850", "1.22", "Yttrium      ", 2, 8, 18, 9, 2)          ' 39
    AddChemElement i, New_ChemElement(i, "Z ", "91.224000", "1.33", "Zirconium    ", 2, 8, 18, 10, 2)         ' 40
    AddChemElement i, New_ChemElement(i, "Nb", "92.906380", "1.60", "Niob         ", 2, 8, 18, 12, 1)         ' 41
    AddChemElement i, New_ChemElement(i, "Mo", "95.964000", "2.16", "Molybdän     ", 2, 8, 18, 13, 1)         ' 42
    AddChemElement i, New_ChemElement(i, "Tc", "98.910000", "1.90", "Technetium   ", 2, 8, 18, 13, 2)         ' 43
    AddChemElement i, New_ChemElement(i, "Ru", "101.07000", "2.20", "Ruthenium    ", 2, 8, 18, 15, 1)         ' 44
    AddChemElement i, New_ChemElement(i, "Rh", "102.90550", "2.28", "Rhodium      ", 2, 8, 18, 16, 1)         ' 45
    AddChemElement i, New_ChemElement(i, "Pd", "106.42000", "2.20", "Palladium    ", 2, 8, 18, 18)            ' 46
    AddChemElement i, New_ChemElement(i, "Ag", "107.86820", "1.93", "Silber       ", 2, 8, 18, 18, 1)         ' 47
    AddChemElement i, New_ChemElement(i, "Cd", "112.41100", "1.69", "Cadmium      ", 2, 8, 18, 18, 2)         ' 48
    AddChemElement i, New_ChemElement(i, "In", "114.81800", "1.78", "Indium       ", 2, 8, 18, 18, 3)         ' 49
    AddChemElement i, New_ChemElement(i, "Sn", "118.71000", "1.96", "Zinn         ", 2, 8, 18, 18, 4)         ' 50
    AddChemElement i, New_ChemElement(i, "Sb", "121.76000", "2.05", "Antimon      ", 2, 8, 18, 18, 5)         ' 51
    AddChemElement i, New_ChemElement(i, "Te", "127.60000", "2.66", "Tellur       ", 2, 8, 18, 18, 6)         ' 52
    AddChemElement i, New_ChemElement(i, "I ", "126.90447", "2.10", "Iod          ", 2, 8, 18, 18, 7)         ' 53
    AddChemElement i, New_ChemElement(i, "Xe", "131.29000", "2.60", "Xenon        ", 2, 8, 18, 18, 8)         ' 54
    
    AddChemElement i, New_ChemElement(i, "Cs", "132.90545", "0.79", "Caesium      ", 2, 8, 18, 18, 8, 1)      ' 55
    AddChemElement i, New_ChemElement(i, "Ba", "137.32700", "0.89", "Barium       ", 2, 8, 18, 18, 8, 2)      ' 56
    
    AddChemElement i, New_ChemElement(i, "La", "138.90550", "1.10", "Lanthan      ", 2, 8, 18, 18, 9, 2)      ' 57
    AddChemElement i, New_ChemElement(i, "Ce", "140.11600", "1.12", "Cer          ", 2, 8, 18, 19, 9, 2)      ' 58
    AddChemElement i, New_ChemElement(i, "Pr", "140.90765", "1.13", "Praseodym    ", 2, 8, 18, 21, 8, 2)      ' 59
    AddChemElement i, New_ChemElement(i, "Nd", "144.24000", "1.14", "Neodym       ", 2, 8, 18, 22, 8, 2)      ' 60
    AddChemElement i, New_ChemElement(i, "Pm", "145.00000", "1.13", "Promethium   ", 2, 8, 18, 23, 8, 2)      ' 61
    AddChemElement i, New_ChemElement(i, "Sm", "150.36000", "1.17", "Samarium     ", 2, 8, 18, 24, 8, 2)      ' 62
    AddChemElement i, New_ChemElement(i, "Eu", "151.96400", "1.20", "Europium     ", 2, 8, 18, 25, 8, 2)      ' 63
    AddChemElement i, New_ChemElement(i, "Gd", "157.25000", "1.20", "Gadolinium   ", 2, 8, 18, 25, 9, 2)      ' 64
    AddChemElement i, New_ChemElement(i, "Tb", "158.92534", "1.10", "Terbium      ", 2, 8, 18, 27, 8, 2)      ' 65
    AddChemElement i, New_ChemElement(i, "Dy", "162.50000", "1.22", "Dysprosium   ", 2, 8, 18, 28, 8, 2)      ' 66
    AddChemElement i, New_ChemElement(i, "Ho", "164.93032", "1.23", "Holmium      ", 2, 8, 18, 29, 8, 2)      ' 67
    AddChemElement i, New_ChemElement(i, "Er", "167.26000", "1.24", "Erbium       ", 2, 8, 18, 30, 8, 2)      ' 68
    AddChemElement i, New_ChemElement(i, "Tm", "168.93421", "1.25", "Thulium      ", 2, 8, 18, 31, 8, 2)      ' 69
    AddChemElement i, New_ChemElement(i, "Yb", "173.04000", "1.10", "Ytterbium    ", 2, 8, 18, 32, 8, 2)      ' 70
    AddChemElement i, New_ChemElement(i, "Lu", "174.96700", "1.27", "Lutetium     ", 2, 8, 18, 32, 9, 2)      ' 71
    
    AddChemElement i, New_ChemElement(i, "Hf", "178.49000", "1.30", "Hafnium      ", 2, 8, 18, 32, 10, 2)     ' 72
    AddChemElement i, New_ChemElement(i, "Ta", "180.94790", "1.50", "Tantal       ", 2, 8, 18, 32, 11, 2)     ' 73
    AddChemElement i, New_ChemElement(i, "W ", "183.84000", "2.36", "Wolfram      ", 2, 8, 18, 32, 12, 2)     ' 74
    AddChemElement i, New_ChemElement(i, "Re", "186.20700", "1.90", "Rhenium      ", 2, 8, 18, 32, 13, 2)     ' 75
    AddChemElement i, New_ChemElement(i, "Os", "190.23000", "2.20", "Osmium       ", 2, 8, 18, 32, 14, 2)     ' 76
    AddChemElement i, New_ChemElement(i, "Ir", "192.21700", "2.20", "Iridium      ", 2, 8, 18, 32, 15, 2)     ' 77
    AddChemElement i, New_ChemElement(i, "Pt", "195.07800", "2.28", "Platin       ", 2, 8, 18, 32, 17, 1)     ' 78
    AddChemElement i, New_ChemElement(i, "Au", "196.96655", "2.54", "Gold         ", 2, 8, 18, 32, 18, 1)     ' 79
    AddChemElement i, New_ChemElement(i, "Hg", "200.59000", "1.90", "Quecksilber  ", 2, 8, 18, 32, 18, 2)     ' 80
    AddChemElement i, New_ChemElement(i, "Tl", "204.38330", "1.62", "Thallium     ", 2, 8, 18, 32, 18, 3)     ' 81
    AddChemElement i, New_ChemElement(i, "Pb", "207.20000", "2.33", "Blei         ", 2, 8, 18, 32, 18, 4)     ' 82
    AddChemElement i, New_ChemElement(i, "Bi", "208.98038", "2.02", "Bismut       ", 2, 8, 18, 32, 18, 5)     ' 83
    AddChemElement i, New_ChemElement(i, "Po", "209.98000", "2.00", "Polonium     ", 2, 8, 18, 32, 18, 6)     ' 84
    AddChemElement i, New_ChemElement(i, "At", "210.00000", "2.20", "Astat        ", 2, 8, 18, 32, 18, 7)     ' 85
    AddChemElement i, New_ChemElement(i, "Rn", "222.00000", "0.00", "Radon        ", 2, 8, 18, 32, 18, 8)     ' 86
    
    AddChemElement i, New_ChemElement(i, "Fr", "223.00000", "0.70", "Francium     ", 2, 8, 18, 32, 18, 8, 1)  ' 87
    AddChemElement i, New_ChemElement(i, "Ra", "226.03000", "0.89", "Radium       ", 2, 8, 18, 32, 18, 8, 2)  ' 88
    AddChemElement i, New_ChemElement(i, "Ac", "227.00", "1.10", "Actinium     ", 2, 8, 18, 32, 18, 9, 2)     ' 89
    
    AddChemElement i, New_ChemElement(i, "Th", "232.04", "1.50", "Thorium      ", 2, 8, 18, 32, 18, 10, 2)    ' 90
    AddChemElement i, New_ChemElement(i, "Pa", "231.04", "1.30", "Protactinium ", 2, 8, 18, 32, 20, 9, 2)     ' 91
    AddChemElement i, New_ChemElement(i, "U ", "238.03", "1.36", "Uran         ", 2, 8, 18, 32, 21, 9, 2)     ' 92
    AddChemElement i, New_ChemElement(i, "Np", "237.05", "1.38", "Neptunium    ", 2, 8, 18, 32, 22, 9, 2)     ' 93
    AddChemElement i, New_ChemElement(i, "Pu", "244.10", "1.30", "Plutonium    ", 2, 8, 18, 32, 24, 8, 2)     ' 94
    AddChemElement i, New_ChemElement(i, "Am", "243.10", "1.28", "Americium    ", 2, 8, 18, 32, 25, 8, 2)     ' 95
    AddChemElement i, New_ChemElement(i, "Cm", "247.10", "1.30", "Curium       ", 2, 8, 18, 32, 25, 9, 2)     ' 96
    AddChemElement i, New_ChemElement(i, "Bk", "247.10", "1.30", "Berkelium    ", 2, 8, 18, 32, 25, 10, 2)    ' 97
    AddChemElement i, New_ChemElement(i, "Cf", "251.10", "1.30", "Californium  ", 2, 8, 18, 32, 28, 8, 2)     ' 98
    AddChemElement i, New_ChemElement(i, "Es", "254.10", "1.30", "Einsteinium  ", 2, 8, 18, 32, 29, 8, 2)     ' 99
    AddChemElement i, New_ChemElement(i, "Fm", "257.10", "1.30", "Fermium      ", 2, 8, 18, 32, 30, 8, 2)     '100
    AddChemElement i, New_ChemElement(i, "Md", "258.00", "1.30", "Mendelevium  ", 2, 8, 18, 32, 31, 8, 2)     '101
    AddChemElement i, New_ChemElement(i, "No", "259.00", "1.30", "Nobelium     ", 2, 8, 18, 32, 32, 8, 2)     '102
    AddChemElement i, New_ChemElement(i, "Lr", "260.00", "1.30", "Lawrencium   ", 2, 8, 18, 32, 32, 9, 2)     '103
    
    AddChemElement i, New_ChemElement(i, "Rf", "261.00", "0.00", "Rutherfordium", 2, 8, 18, 32, 32, 10, 2)    '104
    AddChemElement i, New_ChemElement(i, "Db", "262.00", "0.00", "Dubnium      ", 2, 8, 18, 32, 32, 11, 2)    '105
    AddChemElement i, New_ChemElement(i, "Sg", "263.00", "0.00", "Seaborgium   ", 2, 8, 18, 32, 32, 12, 2)    '106
    AddChemElement i, New_ChemElement(i, "Bh", "262.00", "0.00", "Bohrium      ", 2, 8, 18, 32, 32, 13, 2)    '107
    AddChemElement i, New_ChemElement(i, "Hs", "265.00", "0.00", "Hassium      ", 2, 8, 18, 32, 32, 14, 2)    '108
    AddChemElement i, New_ChemElement(i, "Mt", "266.00", "0.00", "Meitnerium   ", 2, 8, 18, 32, 32, 15, 2)    '109
    AddChemElement i, New_ChemElement(i, "Ds", "269.00", "0.00", "Darmstadtium ", 2, 8, 18, 32, 32, 17, 1)    '110
    AddChemElement i, New_ChemElement(i, "Rg", "272.00", "0.00", "Roentgenium  ", 2, 8, 18, 32, 32, 18, 1)    '111
    AddChemElement i, New_ChemElement(i, "Cn", "277.00", "0.00", "Copernicium  ", 2, 8, 18, 32, 32, 18, 2)    '112
    AddChemElement i, New_ChemElement(i, "Nh", "287.00", "0.00", "Nihonium     ", 2, 8, 18, 32, 32, 18, 3)    '113
    AddChemElement i, New_ChemElement(i, "Fl", "289.00", "0.00", "Flerovium    ", 2, 8, 18, 32, 32, 18, 4)    '114
    AddChemElement i, New_ChemElement(i, "Mc", "288.00", "0.00", "Moscovium    ", 2, 8, 18, 32, 32, 18, 5)    '115
    AddChemElement i, New_ChemElement(i, "Lv", "289.00", "0.00", "Livermorium  ", 2, 8, 18, 32, 32, 18, 6)    '116
    AddChemElement i, New_ChemElement(i, "Ts", "293.00", "0.00", "Tenness      ", 2, 8, 18, 32, 32, 18, 7)    '117
    AddChemElement i, New_ChemElement(i, "Og", "294.00", "0.00", "Oganesson    ", 2, 8, 18, 32, 32, 18, 8)    '118
    
    'ce(i) = New_ChemElement(i, "", 0, "")
    
    '#CreateChemElements = ce
End Sub
'    AddChemElement i, New_ChemElement(i, "H", 1.0079, 2.1, "Wasserstoff  ", 1)                        '  1
'    AddChemElement i, New_ChemElement(i, "He", 4.0026, 0#, "Helium       ", 2)                        '  2
'
'    AddChemElement i, New_ChemElement(i, "Li", 6.941, 1#, "Lithium       ", 2, 1)                     '  3
'    AddChemElement i, New_ChemElement(i, "Be", 9.0122, 1.5, "Beryllium   ", 2, 2)                     '  4
'    AddChemElement i, New_ChemElement(i, "B", 10.811, 2#, "Bor           ", 2, 3)                     '  5
'    AddChemElement i, New_ChemElement(i, "C", 12.011, 2.5, "Kohlenstoff  ", 2, 4)                     '  6
'    AddChemElement i, New_ChemElement(i, "N", 14.007, 3#, "Stickstoff    ", 2, 5)                     '  7
'    AddChemElement i, New_ChemElement(i, "O", 15.999, 3.5, "Sauerstoff   ", 2, 6)                     '  8
'    AddChemElement i, New_ChemElement(i, "F", 18.988, 4#, "Fluor         ", 2, 7)                     '  9
'    AddChemElement i, New_ChemElement(i, "Ne", 20.18, 0#, "Neon          ", 2, 8)                     ' 10
'
'    AddChemElement i, New_ChemElement(i, "Na", 22.99, 0.9, "Natrium      ", 2, 8, 1)                  ' 11
'    AddChemElement i, New_ChemElement(i, "Mg", 24.305, 1.2, "Magnesium   ", 2, 8, 2)                  ' 12
'    AddChemElement i, New_ChemElement(i, "Al", 26.982, 1.5, "Aluminium   ", 2, 8, 3)                  ' 13
'    AddChemElement i, New_ChemElement(i, "Si", 28.086, 1.8, "Silicium    ", 2, 8, 4)                  ' 14
'    AddChemElement i, New_ChemElement(i, "P", 30.974, 2.1, "Phosphor     ", 2, 8, 5)                  ' 15
'    AddChemElement i, New_ChemElement(i, "S", 32.065, 2.5, "Schwefel     ", 2, 8, 6)                  ' 16
'    AddChemElement i, New_ChemElement(i, "Cl", 35.453, 3#, "Chlor        ", 2, 8, 7)                  ' 17
'    AddChemElement i, New_ChemElement(i, "Ar", 39.948, 0#, "Argon        ", 2, 8, 8)                  ' 18
'
'    AddChemElement i, New_ChemElement(i, "K", 39.098, 0.8, "Kalium       ", 2, 8, 8, 1)               ' 19
'    AddChemElement i, New_ChemElement(i, "Ca", 40.078, 1#, "Calcium      ", 2, 8, 8, 2)               ' 20
'    AddChemElement i, New_ChemElement(i, "Sc", 44.956, 1.3, "Scandium    ", 2, 8, 9, 2)               ' 21
'    AddChemElement i, New_ChemElement(i, "Ti", 47.867, 1.5, "Titan       ", 2, 8, 10, 2)              ' 22
'    AddChemElement i, New_ChemElement(i, "V", 50.942, 1.6, "Vanadium     ", 2, 8, 11, 2)              ' 23
'    AddChemElement i, New_ChemElement(i, "Cr", 51.996, 1.6, "Chrom       ", 2, 8, 13, 1)              ' 24
'    AddChemElement i, New_ChemElement(i, "Mn", 54.938, 1.5, "Mangan      ", 2, 8, 13, 2)              ' 25
'    AddChemElement i, New_ChemElement(i, "Fe", 55.845, 1.8, "Eisen       ", 2, 8, 14, 2)              ' 26
'    AddChemElement i, New_ChemElement(i, "Co", 58.933, 1.8, "Cobalt      ", 2, 8, 15, 2)              ' 27
'    AddChemElement i, New_ChemElement(i, "Ni", 58.693, 1.8, "Nickel      ", 2, 8, 16, 2)              ' 28
'    AddChemElement i, New_ChemElement(i, "Cu", 63.546, 1.9, "Kupfer      ", 2, 8, 18, 1)              ' 29
'    AddChemElement i, New_ChemElement(i, "Zn", 65.38, 1.6, "Zink         ", 2, 8, 18, 2)              ' 30
'    AddChemElement i, New_ChemElement(i, "Ga", 69.723, 1.6, "Gallium     ", 2, 8, 18, 3)              ' 31
'    AddChemElement i, New_ChemElement(i, "Ge", 72.64, 1.8, "Germanium    ", 2, 8, 18, 4)              ' 32
'    AddChemElement i, New_ChemElement(i, "As", 74.922, 2#, "Arsen        ", 2, 8, 18, 5)              ' 33
'    AddChemElement i, New_ChemElement(i, "Se", 78.96, 2.4, "Selen        ", 2, 8, 18, 6)              ' 34
'    AddChemElement i, New_ChemElement(i, "Br", 79.904, 2.8, "Brom        ", 2, 8, 18, 7)              ' 35
'    AddChemElement i, New_ChemElement(i, "Kr", 83.798, 0#, "Krypton      ", 2, 8, 18, 8)              ' 36
'
'    AddChemElement i, New_ChemElement(i, "Rb", 85.468, 0.8, "Rubidium    ", 2, 8, 18, 8, 1)           ' 37
'    AddChemElement i, New_ChemElement(i, "Sr", 87.62, 1#, "Strontium     ", 2, 8, 18, 8, 2)           ' 38
'    AddChemElement i, New_ChemElement(i, "Y", 88.906, 1.3, "Yttrium      ", 2, 8, 18, 9, 2)           ' 39
'    AddChemElement i, New_ChemElement(i, "Z", 91.224, 1.4, "Zirconium    ", 2, 8, 18, 10, 2)          ' 40
'    AddChemElement i, New_ChemElement(i, "Nb", 92.906, 1.6, "Niob        ", 2, 8, 18, 12, 1)          ' 41
'    AddChemElement i, New_ChemElement(i, "Mo", 95.96, 1.8, "Molybdän     ", 2, 8, 18, 13, 1)          ' 42
'    AddChemElement i, New_ChemElement(i, "Tc", 98.91, 1.9, "Technetium   ", 2, 8, 18, 13, 2)          ' 43
'    AddChemElement i, New_ChemElement(i, "Ru", 101.07, 2.2, "Ruthenium   ", 2, 8, 18, 15, 1)          ' 44
'    AddChemElement i, New_ChemElement(i, "Rh", 102.91, 2.2, "Rhodium     ", 2, 8, 18, 16, 1)          ' 45
'    AddChemElement i, New_ChemElement(i, "Pd", 106.42, 2.2, "Palladium   ", 2, 8, 18, 18)             ' 46
'    AddChemElement i, New_ChemElement(i, "Ag", 107.87, 1.9, "Silber      ", 2, 8, 18, 18, 1)          ' 47
'    AddChemElement i, New_ChemElement(i, "Cd", 112.41, 1.7, "Cadmium     ", 2, 8, 18, 18, 2)          ' 48
'    AddChemElement i, New_ChemElement(i, "In", 114.82, 1.7, "Indium      ", 2, 8, 18, 18, 3)          ' 49
'    AddChemElement i, New_ChemElement(i, "Sn", 118.71, 1.8, "Zinn        ", 2, 8, 18, 18, 4)          ' 50
'    AddChemElement i, New_ChemElement(i, "Sb", 121.76, 1.9, "Antimon     ", 2, 8, 18, 18, 5)          ' 51
'    AddChemElement i, New_ChemElement(i, "Te", 127.6, 2.1, "Tellur       ", 2, 8, 18, 18, 6)          ' 52
'    AddChemElement i, New_ChemElement(i, "I", 126.9, 2.5, "Iod           ", 2, 8, 18, 18, 7)          ' 53
'    AddChemElement i, New_ChemElement(i, "Xe", 131.29, 0#, "Xenon        ", 2, 8, 18, 18, 8)          ' 54
'
'    AddChemElement i, New_ChemElement(i, "Cs", 132.91, 0.7, "Caesium     ", 2, 8, 18, 18, 8, 1)       ' 55
'    AddChemElement i, New_ChemElement(i, "Ba", 137.33, 0.9, "Barium      ", 2, 8, 18, 18, 8, 2)       ' 56
'    AddChemElement i, New_ChemElement(i, "La", 138.91, 1.1, "Lanthan     ", 2, 8, 18, 18, 9, 2)       ' 57
'    AddChemElement i, New_ChemElement(i, "Ce", 140.12, 1.1, "Cer         ", 2, 8, 18, 19, 9, 2)       ' 58
'    AddChemElement i, New_ChemElement(i, "Pr", 140.91, 1.1, "Praseodym   ", 2, 8, 18, 21, 8, 2)       ' 59
'    AddChemElement i, New_ChemElement(i, "Nd", 144.24, 1.1, "Neodym      ", 2, 8, 18, 22, 8, 2)       ' 60
'    AddChemElement i, New_ChemElement(i, "Pm", 146.9, 1.1, "Promethium   ", 2, 8, 18, 23, 8, 2)       ' 61
'    AddChemElement i, New_ChemElement(i, "Sm", 150.36, 1.2, "Samarium    ", 2, 8, 18, 24, 8, 2)       ' 62
'    AddChemElement i, New_ChemElement(i, "Eu", 151.96, 1.2, "Europium    ", 2, 8, 18, 25, 8, 2)       ' 63
'    AddChemElement i, New_ChemElement(i, "Gd", 157.25, 1.2, "Gadolinium  ", 2, 8, 18, 25, 9, 2)       ' 64
'    AddChemElement i, New_ChemElement(i, "Tb", 158.93, 1.2, "Terbium     ", 2, 8, 18, 27, 8, 2)       ' 65
'    AddChemElement i, New_ChemElement(i, "Dy", 162.5, 1.2, "Dysprosium   ", 2, 8, 18, 28, 8, 2)       ' 66
'    AddChemElement i, New_ChemElement(i, "Ho", 164.93, 1.2, "Holmium     ", 2, 8, 18, 29, 8, 2)       ' 67
'    AddChemElement i, New_ChemElement(i, "Er", 167.26, 1.2, "Erbium      ", 2, 8, 18, 30, 8, 2)       ' 68
'    AddChemElement i, New_ChemElement(i, "Tm", 168.93, 1.2, "Thulium     ", 2, 8, 18, 31, 8, 2)       ' 69
'    AddChemElement i, New_ChemElement(i, "Yb", 173.05, 1.2, "Ytterbium   ", 2, 8, 18, 32, 8, 2)       ' 70
'    AddChemElement i, New_ChemElement(i, "Lu", 174.97, 1.2, "Lutetium    ", 2, 8, 18, 32, 9, 2)       ' 71
'    AddChemElement i, New_ChemElement(i, "Hf", 178.49, 1.3, "Hafnium     ", 2, 8, 18, 32, 10, 2)      ' 72
'    AddChemElement i, New_ChemElement(i, "Ta", 180.95, 1.5, "Tantal      ", 2, 8, 18, 32, 11, 2)      ' 73
'    AddChemElement i, New_ChemElement(i, "W", 183.84, 1.7, "Wolfram      ", 2, 8, 18, 32, 12, 2)      ' 74
'    AddChemElement i, New_ChemElement(i, "Re", 186.21, 1.9, "Rhenium     ", 2, 8, 18, 32, 13, 2)      ' 75
'    AddChemElement i, New_ChemElement(i, "Os", 190.23, 2.2, "Osmium      ", 2, 8, 18, 32, 14, 2)      ' 76
'    AddChemElement i, New_ChemElement(i, "Ir", 192.22, 2.2, "Iridium     ", 2, 8, 18, 32, 15, 2)      ' 77
'    AddChemElement i, New_ChemElement(i, "Pt", 195.08, 2.2, "Platin      ", 2, 8, 18, 32, 17, 1)      ' 78
'    AddChemElement i, New_ChemElement(i, "Au", 196.97, 2.4, "Gold        ", 2, 8, 18, 32, 18, 1)      ' 79
'    AddChemElement i, New_ChemElement(i, "Hg", 200.59, 1.9, "Quecksilber ", 2, 8, 18, 32, 18, 2)      ' 80
'    AddChemElement i, New_ChemElement(i, "Tl", 204.38, 1.8, "Thallium    ", 2, 8, 18, 32, 18, 3)      ' 81
'    AddChemElement i, New_ChemElement(i, "Pb", 207.2, 1.8, "Blei         ", 2, 8, 18, 32, 18, 4)      ' 82
'    AddChemElement i, New_ChemElement(i, "Bi", 208.98, 1.9, "Bismut      ", 2, 8, 18, 32, 18, 5)      ' 83
'    AddChemElement i, New_ChemElement(i, "Po", 209.98, 2#, "Polonium     ", 2, 8, 18, 32, 18, 6)      ' 84
'    AddChemElement i, New_ChemElement(i, "At", 210, 2.2, "Astat          ", 2, 8, 18, 32, 18, 7)      ' 85
'    AddChemElement i, New_ChemElement(i, "Rn", 222, 0#, "Radon           ", 2, 8, 18, 32, 18, 8)      ' 86
'
'    AddChemElement i, New_ChemElement(i, "Fr", 223, 0.7, "Francium       ", 2, 8, 18, 32, 18, 8, 1)   ' 87
'    AddChemElement i, New_ChemElement(i, "Ra", 226.03, 0.9, "Radium      ", 2, 8, 18, 32, 18, 8, 2)   ' 88
'    AddChemElement i, New_ChemElement(i, "Ac", 227, 1.1, "Actinium       ", 2, 8, 18, 32, 18, 9, 2)   ' 89
'    AddChemElement i, New_ChemElement(i, "Th", 232.04, 1.3, "Thorium     ", 2, 8, 18, 32, 18, 10, 2)  ' 90
'    AddChemElement i, New_ChemElement(i, "Pa", 231.04, 1.5, "Protactinium", 2, 8, 18, 32, 20, 9, 2)   ' 91
'    AddChemElement i, New_ChemElement(i, "U", 238.03, 1.4, "Uran         ", 2, 8, 18, 32, 21, 9, 2)   ' 92
'    AddChemElement i, New_ChemElement(i, "Np", 237.05, 1.3, "Neptunium   ", 2, 8, 18, 32, 22, 9, 2)   ' 93
'    AddChemElement i, New_ChemElement(i, "Pu", 244.1, 1.3, "Plutonium    ", 2, 8, 18, 32, 24, 8, 2)   ' 94
'    AddChemElement i, New_ChemElement(i, "Am", 243.1, 1.3, "Americium    ", 2, 8, 18, 32, 25, 8, 2)   ' 95
'    AddChemElement i, New_ChemElement(i, "Cm", 247.1, 1.3, "Curium       ", 2, 8, 18, 32, 25, 9, 2)   ' 96
'    AddChemElement i, New_ChemElement(i, "Bk", 247.1, 1.3, "Berkelium    ", 2, 8, 18, 32, 25, 10, 2)  ' 97
'    AddChemElement i, New_ChemElement(i, "Cf", 251.1, 1.3, "Californium  ", 2, 8, 18, 32, 28, 8, 2)   ' 98
'    AddChemElement i, New_ChemElement(i, "Es", 254.1, 1.3, "Einsteinium  ", 2, 8, 18, 32, 29, 8, 2)   ' 99
'    AddChemElement i, New_ChemElement(i, "Fm", 257.1, 1.3, "Fermium      ", 2, 8, 18, 32, 30, 8, 2)   '100
'    AddChemElement i, New_ChemElement(i, "Md", 258, 1.3, "Mendelevium    ", 2, 8, 18, 32, 31, 8, 2)   '101
'    AddChemElement i, New_ChemElement(i, "No", 259, 1.3, "Nobelium       ", 2, 8, 18, 32, 32, 8, 2)   '102
'    AddChemElement i, New_ChemElement(i, "Lr", 260, 1.3, "Lawrencium     ", 2, 8, 18, 32, 32, 9, 2)   '103
'    AddChemElement i, New_ChemElement(i, "Rf", 261, 0#, "Rutherfordium   ", 2, 8, 18, 32, 32, 10, 2)  '104
'    AddChemElement i, New_ChemElement(i, "Db", 262, 0#, "Dubnium         ", 2, 8, 18, 32, 32, 11, 2)  '105
'    AddChemElement i, New_ChemElement(i, "Sg", 263, 0#, "Seaborgium      ", 2, 8, 18, 32, 32, 12, 2)  '106
'    AddChemElement i, New_ChemElement(i, "Bh", 262, 0#, "Bohrium         ", 2, 8, 18, 32, 32, 13, 2)  '107
'    AddChemElement i, New_ChemElement(i, "Hs", 265, 0#, "Hassium         ", 2, 8, 18, 32, 32, 14, 2)  '108
'    AddChemElement i, New_ChemElement(i, "Mt", 266, 0#, "Meitnerium      ", 2, 8, 18, 32, 32, 15, 2)  '109
'    AddChemElement i, New_ChemElement(i, "Ds", 269, 0#, "Darmstadtium    ", 2, 8, 18, 32, 32, 17, 1)  '110
'    AddChemElement i, New_ChemElement(i, "Rg", 272, 0#, "Roentgenium     ", 2, 8, 18, 32, 32, 18, 1)  '111
'    AddChemElement i, New_ChemElement(i, "Cn", 277, 0#, "Copernicium     ", 2, 8, 18, 32, 32, 18, 2)  '112
'    AddChemElement i, New_ChemElement(i, "Nh", 287, 0#, "Nihonium        ", 2, 8, 18, 32, 32, 18, 3)  '113
'    AddChemElement i, New_ChemElement(i, "Fl", 289, 0#, "Flerovium       ", 2, 8, 18, 32, 32, 18, 4)  '114
'    AddChemElement i, New_ChemElement(i, "Mc", 288, 0#, "Moscovium       ", 2, 8, 18, 32, 32, 18, 5)  '115
'    AddChemElement i, New_ChemElement(i, "Lv", 289, 0#, "Livermorium     ", 2, 8, 18, 32, 32, 18, 6)  '116
'    AddChemElement i, New_ChemElement(i, "Ts", 293, 0#, "Tenness         ", 2, 8, 18, 32, 32, 18, 7)  '117
'    AddChemElement i, New_ChemElement(i, "Og", 294, 0#, "Oganesson       ", 2, 8, 18, 32, 32, 18, 8)  '118

