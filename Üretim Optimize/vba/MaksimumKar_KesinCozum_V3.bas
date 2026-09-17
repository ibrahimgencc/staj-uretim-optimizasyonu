Attribute VB_Name = "MaksimumKar_KesinCozum_V3"
'Attribute VB_Name = "MaksimumKarSolver_V4"
Option Explicit

Public Sub MaksimumKar_EnAz2_V4()
    Dim ws As Worksheet
    Dim sonSatir As Long
    Dim r As Long
    Dim c As Long
    Dim sonuc As Variant
    Dim degiskenAdres As String
    Dim toplamSaat As Double
    Dim asgariGerekliSaat As Double
    Dim hucre As Range

    On Error GoTo HataYakala

    Set ws = ThisWorkbook.Worksheets("Sayfa1")
    ws.Activate

    sonSatir = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If sonSatir > 202 Then sonSatir = 202
    If sonSatir < 3 Then
        MsgBox "A sutununda optimize edilecek urun bulunamadi.", vbExclamation
        Exit Sub
    End If

    ' Kapasitelerin tamami sayisal ve sifir/pozitif olmali.
    For Each hucre In ws.Range("E203:R203").Cells
        If Len(Trim$(CStr(hucre.Value2))) = 0 Or Not IsNumeric(hucre.Value2) Then
            MsgBox hucre.Address(False, False) & " hucresine sayisal kapasite girin.", vbExclamation, "Eksik Kapasite"
            Exit Sub
        End If
        If CDbl(hucre.Value2) < 0 Then
            MsgBox hucre.Address(False, False) & " kapasitesi negatif olamaz.", vbExclamation, "Hatali Kapasite"
            Exit Sub
        End If
    Next hucre

    ' Fiyati olan her urun en az bir makinede islem gormeli; aksi halde model sinirsiz olur.
    For r = 3 To sonSatir
        If Len(Trim$(CStr(ws.Cells(r, "A").Value2))) > 0 Then
            If Not IsNumeric(ws.Cells(r, "C").Value2) Then
                MsgBox "C" & r & " birim satis fiyati sayisal degil.", vbExclamation, "Hatali Fiyat"
                Exit Sub
            End If
            toplamSaat = Application.Sum(ws.Range("E" & r & ":R" & r))
            If CDbl(ws.Cells(r, "C").Value2) > 0 And toplamSaat <= 0 Then
                MsgBox "Satir " & r & " icin fiyat var ancak makine islem saati yok. Model sinirsiz olur.", vbExclamation, "Eksik Islem Saati"
                Exit Sub
            End If
        End If
    Next r

    ' Tum malzemelerden en az 2 adet icin kapasitenin yeterli oldugunu onceden kontrol et.
    For c = 5 To 18
        asgariGerekliSaat = 2 * Application.Sum(ws.Range(ws.Cells(3, c), ws.Cells(sonSatir, c)))
        If asgariGerekliSaat > CDbl(ws.Cells(203, c).Value2) + 0.0000001 Then
            MsgBox ws.Cells(1, c).Value & " makinesi icin kapasite yetersiz." & vbCrLf & _
                   "En az gerekli saat: " & Format(asgariGerekliSaat, "0.00") & vbCrLf & _
                   "Girilen kapasite: " & Format(ws.Cells(203, c).Value2, "0.00"), _
                   vbExclamation, "En Az 2 Adet Uretilemiyor"
            Exit Sub
        End If
    Next c

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationAutomatic

    ws.Range("D3:D" & sonSatir).Value2 = 2
    ws.Range("U2").Value = "MAKSIMUM KAR"
    ws.Range("U3").Formula = "=SUMPRODUCT($C$3:$C$" & sonSatir & ",$D$3:$D$" & sonSatir & ")-SUMPRODUCT($E$204:$R$204,$E$2:$R$2)"
    ws.Range("E204").Formula = "=SUMPRODUCT(E$3:E$" & sonSatir & ",$D$3:$D$" & sonSatir & ")"
    ws.Range("E204:R204").FillRight
    ws.Range("A203").Value = "MAKINE KAPASITESI"
    ws.Range("A204").Value = "KULLANILAN KAPASITE"
    Application.CalculateFull

    degiskenAdres = "$D$3:$D$" & sonSatir

    ' En uyumlu Solver cagrilari: aktif sayfa ve metin hucre adresleri kullanilir.
    Application.Run "Solver.xlam!SolverReset"
    Application.Run "Solver.xlam!SolverOk", "$U$3", 1, 0, degiskenAdres, 2
    ' Kapasite kisitlarini tek tek eklemek Solver'in aralik-aralik hatasini onler.
    For c = 5 To 18
        Application.Run "Solver.xlam!SolverAdd", _
                        ws.Cells(204, c).Address, 1, _
                        ws.Cells(203, c).Address
    Next c
    Application.Run "Solver.xlam!SolverAdd", degiskenAdres, 4
    Application.Run "Solver.xlam!SolverAdd", degiskenAdres, 3, 2

    Application.CalculateFull
    sonuc = Application.Run("Solver.xlam!SolverSolve", True)
    Application.CalculateFull
    Application.ScreenUpdating = True

    If IsError(sonuc) Then
        MsgBox "Solver modeli okuyamadi. Veri hucrelerinde hata degeri olup olmadigini kontrol edin.", vbExclamation, "Solver Hatasi"
    ElseIf CLng(sonuc) = 0 Or CLng(sonuc) = 1 Or CLng(sonuc) = 2 Then
        MsgBox "Cozum tamamlandi." & vbCrLf & _
               "Optimum miktarlar (her urun en az 2): D3:D" & sonSatir & vbCrLf & _
               "Maksimum kar: U3", vbInformation, "Basarili"
    Else
        MsgBox "Solver sonuc kodu: " & CStr(sonuc), vbExclamation, "Solver Sonucu"
    End If
    Exit Sub

HataYakala:
    Application.ScreenUpdating = True
    MsgBox "Makro hatasi " & Err.Number & ": " & Err.Description & vbCrLf & _
           "Solver Eklentisi etkin olmali.", vbCritical, "Maksimum Kar V4"
End Sub


