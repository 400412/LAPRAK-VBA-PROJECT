Attribute VB_Name = "FUNCTION_MODULE"
' 1. Menghitung Ketidakpastian dengan banyak data n
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
Public Function KSR(RATA_RATA As Double, DELTA As Double) As Double
    
    KSR = Application.WorksheetFunction.Round((DELTA / RATA_RATA * 100), 4)

End Function
' 5. Menghitung Angka Penting (AP)
Public Function ANGKA_PENTING(KSR As Double) As String
    
    If KSR <= 0.1 Then
        ANGKA_PENTING = "4 AP"
        
    ElseIf KSR > 0.1 And KSR < 1 Then
        ANGKA_PENTING = "3 AP"
        
    ElseIf KSR > 1 And KSR <= 10 Then
        ANGKA_PENTING = "2 AP"
        
    ElseIf KSR > 10 And KSR <= 100 Then
        ANGKA_PENTING = "1 AP / ERROR"
        
    Else
        ANGKA_PENTING = "LAH GEDE AMAT, KOCAK NI ORANG, CEK LAGI"
        
    End If
    
End Function
' 6. Pembulatan Ketidapastian sesuai AP
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
'7. Menghitung a (GRAFIK)
Public Function a(n As Integer, X As Double, Y As Double, X_KUADRAT As Double, XY As Double) As Double
    Dim atas As Double
    Dim bawah As Double
    
    atas = (Y * X_KUADRAT) - (X * XY)
    bawah = (n * X_KUADRAT) - (X ^ 2)
    a = atas / bawah

End Function
'8. Menghitung b (GRAFIK)
Public Function b(n As Integer, X As Double, Y As Double, X_KUADRAT As Double, XY As Double) As Double
    Dim atas As Double
    Dim bawah As Double
    
    atas = (n * XY) - (X * Y)
    bawah = (n * X_KUADRAT) - (X ^ 2)
    b = atas / bawah

End Function
'9. Menghitung y (GRAFIK)
Public Function Y(a As Double, b As Double, X As Double) As Double

    Y = a + (b * X)

End Function
'10. Menghitung HASIL
Public Function HASIL(RATA_RATA As Double, DELTA As Double, AP As String) As String
    Dim a1 As String
    Dim a2 As String
    
    If AP = "4 AP" Then
        a1 = Application.WorksheetFunction.Text(RATA_RATA, "0.000E+00")
        a2 = Application.WorksheetFunction.Text(DELTA, "0.000E+00")
        HASIL = Application.WorksheetFunction.Concat(a1, " ± ", a2)
        
    ElseIf AP = "3 AP" Then
        a1 = Application.WorksheetFunction.Text(RATA_RATA, "0.00E+00")
        a2 = Application.WorksheetFunction.Text(DELTA, "0.00E+00")
        HASIL = Application.WorksheetFunction.Concat(a1, " ± ", a2)
    
    ElseIf AP = "2 AP" Then
        a1 = Application.WorksheetFunction.Text(RATA_RATA, "0.0E+00")
        a2 = Application.WorksheetFunction.Text(DELTA, "0.0E+00")
        HASIL = Application.WorksheetFunction.Concat(a1, " ± ", a2)
    
    ElseIf AP = "1 AP / ERROR" Then
        a1 = Application.WorksheetFunction.Text(RATA_RATA, "0E+00")
        a2 = Application.WorksheetFunction.Text(DELTA, "0E+00")
        HASIL = Application.WorksheetFunction.Concat(a1, " ± ", a2)
    
    End If
    
End Function


