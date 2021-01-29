Attribute VB_Name = "MPeriods"
Option Explicit


'Energieniveau ->
'2      8       18                      32
' 1   2   3   4   5   6   7   8   9  10  11  12  13  14  15  16  17  18  19  20  21  22
'1s, 2s, 2p, 3s, 3p, 4s, 3d, 4p, 5s, 4d, 5p, 6s, 4f, 5d, 6p, 7s, 5f, 6d, 7p, 6f, 7d, 7f

'Public Function GetNextEnergyNiveau(ByRef iEn_inout As Long, ByRef iOrd_inout As Long) As IOrbital
Public Function GetNextEnergyNiveau(ByVal iENiveau As Long, ByRef iOrd_inout As Long) As IOrbital
    Dim o1 As Orbital, o2 As Orbital, o3 As Orbital, o4 As Orbital, o5 As Orbital, o6 As Orbital, o7 As Orbital
    Dim o As IOrbital
    Select Case iENiveau
    Case 1, 2, 4, 6, 9, 12, 16:
        'Oh Mann so ist das ja sowieso totaler Käse,
        'kein Wunder dass das nicht klappt, weil der code einfach von vornherein Bockmist ist
        'das Auffüllen mit Elektronen darf entweder nur die Periode machen oder ein S-, P-, D-, oder F-Orbital
        'und nicht das kleine 2er-Orbital selber, nein das muss von außen geschehen,
        'um überhaupt nach der Pauli-Regel vorgehen zu können
        '
        'irgendwie könnte man eine Elektronen-Counter-Iterator-Klasse EIterator ganz gut gebrauchen
        'EIterator wird an den Konstruktor von S-P-D-F-Orbital übergeben, EITerator könnten zum nächsten Energie-Niveau führen,
        'EIteratoir hat die Ordnungszahl, und einen Counter der nach oben zählt und nach unten zählt, also wieviele noch übrig bleiben
        'der EIterator könnte überhaupt die Arbeit mit den Orbitalen übernehmen, d.h. er kennt das Pauli-Prinzip für alle Orbitale etc.
        'die Orbitale sind dann nur noch Eletronen-Datenträger
        '
        '
        '
        '
        'der Aufbau sollte auch anders sein,
        'man könnte eine doppelt verlinkte Liste machen, sowohl die Perioden als auch das EnergyNiveau
        'man könnte auch nochimmer mit UDTypes arbeiten, oder?
        '
        
                                Set o1 = MNew.Orbital(iOrd_inout)
                                'Set o = MNew.OrbitalS(iEn_inout, MNew.Orbital(iOrd_inout))
                                Set o = MNew.OrbitalS(iENiveau) ', o1)
                                'If iOrd_inout <= 0 Then Exit Function
    Case 3, 5, 8, 11, 15, 19:
                                Set o1 = MNew.Orbital(iOrd_inout)
                                Set o2 = MNew.Orbital(iOrd_inout)
                                Set o3 = MNew.Orbital(iOrd_inout)
                                'Set o = MNew.OrbitalP(iEn_inout, MNew.Orbital(iOrd_inout), MNew.Orbital(iOrd_inout), MNew.Orbital(iOrd_inout))
                                Set o = MNew.OrbitalP(iENiveau) ', o1, o2, o3)
                                'If iOrd_inout <= 0 Then Exit Function
    Case 7, 10, 14, 18, 21:
                                Set o1 = MNew.Orbital(iOrd_inout)
                                Set o2 = MNew.Orbital(iOrd_inout)
                                Set o3 = MNew.Orbital(iOrd_inout)
                                Set o4 = MNew.Orbital(iOrd_inout)
                                Set o5 = MNew.Orbital(iOrd_inout)
                                'Set o = MNew.OrbitalD(iEn_inout, MNew.Orbital(iOrd_inout), MNew.Orbital(iOrd_inout), MNew.Orbital(iOrd_inout), MNew.Orbital(iOrd_inout), MNew.Orbital(iOrd_inout))
                                Set o = MNew.OrbitalD(iENiveau) ', o1, o2, o3, o4, o5)
                                'If iOrd_inout <= 0 Then Exit Function
    Case 13, 17, 20, 22:
                                Set o1 = MNew.Orbital(iOrd_inout)
                                Set o2 = MNew.Orbital(iOrd_inout)
                                Set o3 = MNew.Orbital(iOrd_inout)
                                Set o4 = MNew.Orbital(iOrd_inout)
                                Set o5 = MNew.Orbital(iOrd_inout)
                                Set o6 = MNew.Orbital(iOrd_inout)
                                Set o7 = MNew.Orbital(iOrd_inout)
                                'Set o = MNew.OrbitalF(iEn_inout, MNew.Orbital(iOrd_inout), MNew.Orbital(iOrd_inout), MNew.Orbital(iOrd_inout), MNew.Orbital(iOrd_inout), MNew.Orbital(iOrd_inout), MNew.Orbital(iOrd_inout), MNew.Orbital(iOrd_inout))
                                Set o = MNew.OrbitalF(iENiveau) ', o1, o2, o3, o4, o5, o6, o7)
                                'If iOrd_inout <= 0 Then Exit Function
    End Select
    Set GetNextEnergyNiveau = o
End Function

Public Function CreateListOfIOrbital(ByVal iOrd As Long) As ListOfIOrbital
    'iOrd wird runtergezählt
    Dim lo  As New ListOfIOrbital
    Dim orb As IOrbital
    Dim enc As Long ': enc = 1 'Energieniveau-Counter wird hochgezählt
    For enc = 1 To 22
    'Do
        'If iOrd <= 0 Then Exit Do 'For
        Set orb = GetNextEnergyNiveau(enc, iOrd)
        lo.Add orb
        If iOrd <= 0 Then
            'Debug.Print iOrd
            Exit For
        End If

        'enc = enc + 1
    'Loop
    Next
    Set CreateListOfIOrbital = lo
End Function

Sub MessO(o As IOrbital)
    If o Is Nothing Then Debug.Print "Orbital is nothing"
End Sub

Public Function GetPeriods(ByVal iOrd As Long) As ListOfPeriode
    If iOrd = 90 Then
        Debug.Print iOrd & vbNewLine
    End If
    Dim lo       As ListOfIOrbital: Set lo = CreateListOfIOrbital(iOrd)
    'Debug.Print lo.ToStr
    Dim Periods  As New ListOfPeriode
    Dim pc As Long: pc = 1 ' Periode-Counter wird hochgezählt
    Dim OS As OrbitalS
    Dim OP As OrbitalP
    Dim OD As OrbitalD
    Dim OF As OrbitalF

'Energieniveau ->
'2      8       18                      32
' 1   2   3   4   5   6   7   8   9  10  11  12  13  14  15  16  17  18  19  20  21  22
'1s, 2s, 2p, 3s, 3p, 4s, 3d, 4p, 5s, 4d, 5p, 6s, 4f, 5d, 6p, 7s, 5f, 6d, 7p, 6f, 7d, 7f
'1
    If lo.Count > 0 Then Set OS = lo.Item(1).COrbitalS: MessO OS
    If lo.Count > 0 Then Periods.Add MNew.Periode(pc, OS)
    pc = pc + 1: Set OS = Nothing: Set OP = Nothing: Set OD = Nothing: Set OF = Nothing
    
    
' 1   2   3   4   5   6   7   8   9  10  11  12  13  14  15  16  17  18  19  20  21  22
'1s, 2s, 2p, 3s, 3p, 4s, 3d, 4p, 5s, 4d, 5p, 6s, 4f, 5d, 6p, 7s, 5f, 6d, 7p, 6f, 7d, 7f
'2
    If lo.Count > 1 Then Set OS = lo.Item(2).COrbitalS: MessO OS
    If lo.Count > 2 Then Set OP = lo.Item(3).COrbitalP: MessO OP
    If lo.Count > 1 Then Periods.Add MNew.Periode(pc, OS, OP)
    pc = pc + 1: Set OS = Nothing: Set OP = Nothing: Set OD = Nothing: Set OF = Nothing
    
    
' 1   2   3   4   5   6   7   8   9  10  11  12  13  14  15  16  17  18  19  20  21  22
'1s, 2s, 2p, 3s, 3p, 4s, 3d, 4p, 5s, 4d, 5p, 6s, 4f, 5d, 6p, 7s, 5f, 6d, 7p, 6f, 7d, 7f
'3
    If lo.Count > 3 Then Set OS = lo.Item(4).COrbitalS: MessO OS
    If lo.Count > 4 Then Set OP = lo.Item(5).COrbitalP: MessO OP
    If lo.Count > 6 Then Set OD = lo.Item(7).COrbitalD: MessO OD
    If lo.Count > 3 Then Periods.Add MNew.Periode(pc, OS, OP, OD)
    pc = pc + 1: Set OS = Nothing: Set OP = Nothing: Set OD = Nothing: Set OF = Nothing
    
    
' 1   2   3   4   5   6   7   8   9  10  11  12  13  14  15  16  17  18  19  20  21  22
'1s, 2s, 2p, 3s, 3p, 4s, 3d, 4p, 5s, 4d, 5p, 6s, 4f, 5d, 6p, 7s, 5f, 6d, 7p, 6f, 7d, 7f
'4
    If lo.Count > 5 Then Set OS = lo.Item(6).COrbitalS: MessO OS
    If lo.Count > 7 Then Set OP = lo.Item(8).COrbitalP: MessO OP
    If lo.Count > 9 Then Set OD = lo.Item(10).COrbitalD: MessO OD
    If lo.Count > 5 Then Periods.Add MNew.Periode(pc, OS, OP, OD)
    pc = pc + 1: Set OS = Nothing: Set OP = Nothing: Set OD = Nothing: Set OF = Nothing
    
    
' 1   2   3   4   5   6   7   8   9  10  11  12  13  14  15  16  17  18  19  20  21  22
'1s, 2s, 2p, 3s, 3p, 4s, 3d, 4p, 5s, 4d, 5p, 6s, 4f, 5d, 6p, 7s, 5f, 6d, 7p, 6f, 7d, 7f
'5
    If lo.Count > 8 Then Set OS = lo.Item(9).COrbitalS: MessO OS
    If lo.Count > 10 Then Set OP = lo.Item(11).COrbitalP: MessO OP
    If lo.Count > 13 Then Set OD = lo.Item(14).COrbitalD: MessO OD
    If lo.Count > 16 Then Set OF = lo.Item(17).COrbitalF: MessO OF
    If lo.Count > 8 Then Periods.Add MNew.Periode(pc, OS, OP, OD, OF)
    pc = pc + 1: Set OS = Nothing: Set OP = Nothing: Set OD = Nothing: Set OF = Nothing
    
    
' 1   2   3   4   5   6   7   8   9  10  11  12  13  14  15  16  17  18  19  20  21  22
'1s, 2s, 2p, 3s, 3p, 4s, 3d, 4p, 5s, 4d, 5p, 6s, 4f, 5d, 6p, 7s, 5f, 6d, 7p, 6f, 7d, 7f
'6
    If lo.Count > 11 Then Set OS = lo.Item(12).COrbitalS: MessO OS
    If lo.Count > 14 Then Set OP = lo.Item(15).COrbitalP: MessO OP
    If lo.Count > 17 Then Set OD = lo.Item(18).COrbitalD: MessO OD
    If lo.Count > 19 Then Set OF = lo.Item(20).COrbitalF: MessO OF
    If lo.Count > 11 Then Periods.Add MNew.Periode(pc, OS, OP, OD, OF)
    pc = pc + 1: Set OS = Nothing: Set OP = Nothing: Set OD = Nothing: Set OF = Nothing
    
    
' 1   2   3   4   5   6   7   8   9  10  11  12  13  14  15  16  17  18  19  20  21  22
'1s, 2s, 2p, 3s, 3p, 4s, 3d, 4p, 5s, 4d, 5p, 6s, 4f, 5d, 6p, 7s, 5f, 6d, 7p, 6f, 7d, 7f
'7
    If lo.Count > 15 Then Set OS = lo.Item(16).COrbitalS: MessO OS
    If lo.Count > 18 Then Set OP = lo.Item(19).COrbitalP: MessO OP
    If lo.Count > 20 Then Set OP = lo.Item(21).COrbitalD: MessO OD
    If lo.Count > 21 Then Set OF = lo.Item(22).COrbitalF: MessO OF
    If lo.Count > 15 Then Periods.Add MNew.Periode(pc, OS, OP, OD, OF)
    pc = pc + 1: Set OS = Nothing: Set OP = Nothing: Set OD = Nothing: Set OF = Nothing
    
    
    If iOrd = 90 Then
        '90: Thorium : 2/8/18/32/18/10/2
        Debug.Print iOrd & " : " & Periods.Count & " : " & Periods.CountElektrons & "  " & Periods.ToStr
    End If
    Set GetPeriods = Periods
End Function


