' 1. Menghitung Ketidakpastian dengan banyak data n
    ' Ngitung delta kalau datanya selain n = 3 atau n = 5
Public Function DELTA_n(n As Integer, SIGMA As Double, SIGMA_KUADRAT As Double) As Double

    DELTA_n = (1 / n) * (((n * SIGMA_KUADRAT - SIGMA ^ 2) / (n - 1)) ^ 0.5)

End Function

' 2. Menghitung Ketidakpastian dengan banyak data 3
Public Function DELTA_3(SIGMA As Double, SIGMA_KUADRAT As Double) As Double

    DELTA_3 = (1 / 3) * (((3 * SIGMA_KUADRAT - SIGMA ^ 2) / (2)) ^ 0.5)
    
End Function

' 3. Menghitung Ketidakpastian dengan banyak data 5
Public Function DELTA_5(SIGMA As Double, SIGMA_KUADRAT As Double) As Double

    DELTA_5 = (1 / 5) * (((5 * SIGMA_KUADRAT - SIGMA ^ 2) / (4)) ^ 0.5)
    
End Function

' 4. Menghitung Ketidakpastian Relatif (KSR) dalam %
Public Function KSR(RATA_RATA As Double, DELTA As Double) As String
    Dim Tahap_1 As Double
    Dim Tahap_2 As String
    
    Tahap_1 = (DELTA / RATA_RATA)
    Tahap_2 = Format(Tahap_1, "0.00 %")
    KSR = Replace(Tahap_2, ".", ",")

End Function

' 5. Menghitung Angka Penting (AP)
Public Function ANGKA_PENTING(KSR As String) As String
    Dim Tahap_1 As String
    Dim Tahap_2 As Double
    Dim Tahap_3 As Double
    
    Tahap_1 = Replace(KSR, " %", "")
    Tahap_2 = Replace(Tahap_1, ",", ".")
    Tahap_3 = CDbl(Tahap_2)
    
    If Tahap_3 <= 0.1 Then
        ANGKA_PENTING = "4 AP"
        
    ElseIf Tahap_3 > 0.1 And Tahap_3 < 1 Then
        ANGKA_PENTING = "3 AP"
        
    ElseIf Tahap_3 > 1 And Tahap_3 <= 10 Then
        ANGKA_PENTING = "2 AP"
        
    ElseIf Tahap_3 > 10 And Tahap_3 <= 100 Then
        ANGKA_PENTING = "1 AP / ERROR"
        
    Else
        ANGKA_PENTING = "LAH GEDE AMAT, KOCAK NI ORANG, CEK LAGI"
        
    End If
    
End Function

' 6. Pembulatan Ketidapastian sesuai AP
    ' Pembulatan yg berdasarkan Angka Penting
    ' Misalnya NILAI = 0,0000035726278, AP = 4 AP, maka PEMBULATAN = 0,000003573
    ' Misalnya NILAI = 0,0000035726278, AP = 3 AP, maka PEMBULATAN = 0,00000357
    ' Misalnya NILAI = 0,0000035726278, AP = 2 AP, maka PEMBULATAN = 0,0000036
    ' Misalnya NILAI = 0,0000035726278, AP = 1 AP / ERROR, maka PEMBULATAN = 0,000004
Public Function PEMBULATAN(NILAI As Double, AP As String) As Double
    Dim Tahap_1 As Double
    Dim Tahap_2 As Double
    Dim Tahap_3 As Double
        
    If AP = "4 AP" Then
        Tahap_1 = Abs(NILAI)
        Tahap_2 = Application.WorksheetFunction.Log10(Tahap_1)
        Tahap_3 = 1 + (Int(Tahap_2))
        PEMBULATAN = Application.WorksheetFunction.Round(NILAI, 4 - Tahap_3)
        
    ElseIf AP = "3 AP" Then
        Tahap_1 = Abs(NILAI)
        Tahap_2 = Application.WorksheetFunction.Log10(Tahap_1)
        Tahap_3 = 1 + (Int(Tahap_2))
        PEMBULATAN = Application.WorksheetFunction.Round(NILAI, 3 - Tahap_3)
    
    ElseIf AP = "2 AP" Then
        Tahap_1 = Abs(NILAI)
        Tahap_2 = Application.WorksheetFunction.Log10(Tahap_1)
        Tahap_3 = 1 + (Int(Tahap_2))
        PEMBULATAN = Application.WorksheetFunction.Round(NILAI, 2 - Tahap_3)
    
    ElseIf AP = "1 AP / ERROR" Then
        Tahap_1 = Abs(NILAI)
        Tahap_2 = Application.WorksheetFunction.Log10(Tahap_1)
        Tahap_3 = 1 + (Int(Tahap_2))
        PEMBULATAN = Application.WorksheetFunction.Round(NILAI, 1 - Tahap_3)
        
    End If

End Function

'7. Menghitung a atau b (GRAFIK)
    'Dipake khusus buat ngitung hasil a di grafik
Public Function a(n As Integer, x As Double, Y As Double, X_KUADRAT As Double, XY As Double) As Double
    Dim atas As Double
    Dim bawah As Double
    
    atas = (Y * X_KUADRAT) - (x * XY)
    bawah = (n * X_KUADRAT) - (x ^ 2)
    a = atas / bawah

End Function

'8. Menghitung b atau m atau gradien (GRAFIK)
    'Dipake khusus buat ngitung hasil b di grafik
Public Function b(n As Integer, x As Double, Y As Double, X_KUADRAT As Double, XY As Double) As Double
    Dim atas As Double
    Dim bawah As Double
    
    atas = (n * XY) - (x * Y)
    bawah = (n * X_KUADRAT) - (x ^ 2)
    b = atas / bawah

End Function

'9. Menghitung y (GRAFIK)
    'Dipake khusus buat ngitung hasil y di grafik
Public Function Y(a As Double, b As Double, x As Double) As Double

    Y = a + (b * x)

End Function

'10. Menghitung HASIL
    'Buat ngitung hasil akhir dengan ngegabungin Rata-rata dan delta
    'Formatnya sesuai dengan notasi ilmiah (*10^)
Public Function HASIL(RATA_RATA As Double, DELTA As Double, AP As String) As String
    Dim a1 As String
    Dim a2 As String
    
    If AP = "4 AP" Then
        a1 = Format(RATA_RATA, "0,000E+0")
        a1 = Replace(a1, "E", "*10^")
        a2 = Format(DELTA, "0,000E+0")
        a2 = Replace(a2, "E", "*10^")
        HASIL = Application.WorksheetFunction.Concat("(", a1, " ± ", a2, ")")
        
    ElseIf AP = "3 AP" Then
        a1 = Format(RATA_RATA, "0,00E+0")
        a1 = Replace(a1, "E", "*10^")
        a2 = Format(DELTA, "0,00E+0")
        a2 = Replace(a2, "E", "*10^")
        HASIL = Application.WorksheetFunction.Concat("(", a1, " ± ", a2, ")")
    
    ElseIf AP = "2 AP" Then
        a1 = Format(RATA_RATA, "0,0E+0")
        a1 = Replace(a1, "E", "*10^")
        a2 = Format(DELTA, "0,0E+0")
        a2 = Replace(a2, "E", "*10^")
        HASIL = Application.WorksheetFunction.Concat("(", a1, " ± ", a2, ")")
    
    ElseIf AP = "1 AP / ERROR" Then
        a1 = Format(RATA_RATA, "0E+0")
        a1 = Replace(a1, "E", "*10^")
        a2 = Format(DELTA, "0E+0")
        a2 = Replace(a2, "E", "*10^")
        HASIL = Application.WorksheetFunction.Concat("(", a1, " ± ", a2, ")")
    
    End If
    
End Function

'11. Menghitung HASIL (sesuai format equation di Microsoft Word)
    'Buat ngitung hasil akhir dengan ngegabungin Rata-rata dan delta
    'Formatnya sesuai dengan equation di Microsoft Word (tinggal diconvert)
Public Function HASIL_EQU(RATA_RATA As Double, DELTA As Double, AP As String) As String
    Dim a1 As String
    Dim a2 As String
    
    If AP = "4 AP" Then
        a1 = Format(RATA_RATA, "0.000E+0")
        a1 = Replace(a1, ".", ",")
        a1 = Replace(a1, "E", "\bullet10^")
        a1 = Replace(a1, "+", "")
        a2 = Format(DELTA, "0.000E+0")
        a2 = Replace(a2, ".", ",")
        a2 = Replace(a2, "E", "\bullet10^")
        a2 = Replace(a2, "+", "")
        HASIL_EQU = Application.WorksheetFunction.Concat("(", a1, "±", a2, ")")
        
    ElseIf AP = "3 AP" Then
        a1 = Format(RATA_RATA, "0.00E+0")
        a1 = Replace(a1, ".", ",")
        a1 = Replace(a1, "E", "\bullet10^")
        a1 = Replace(a1, "+", "")
        a2 = Format(DELTA, "0.00E+0")
        a2 = Replace(a2, ".", ",")
        a2 = Replace(a2, "E", "\bullet10^")
        a2 = Replace(a2, "+", "")
        HASIL_EQU = Application.WorksheetFunction.Concat("(", a1, "±", a2, ")")
    
    ElseIf AP = "2 AP" Then
        a1 = Format(RATA_RATA, "0.0E+0")
        a1 = Replace(a1, ".", ",")
        a1 = Replace(a1, "E", "\bullet10^")
        a1 = Replace(a1, "+", "")
        a2 = Format(DELTA, "0.0E+0")
        a2 = Replace(a2, ".", ",")
        a2 = Replace(a2, "E", "\bullet10^")
        a2 = Replace(a2, "+", "")
        HASIL_EQU = Application.WorksheetFunction.Concat("(", a1, "±", a2, ")")
    
    ElseIf AP = "1 AP / ERROR" Then
        a1 = Format(RATA_RATA, "0E+0")
        a1 = Replace(a1, "E", "\bullet10^")
        a1 = Replace(a1, "+", "")
        a2 = Format(DELTA, "0E+0")
        a2 = Replace(a2, "E", "\bullet10^")
        a2 = Replace(a2, "+", "")
        HASIL_EQU = Application.WorksheetFunction.Concat("(", a1, "±", a2, ")")
    
    End If
    
End Function

