Private Sub InsiasiModulOutlook()
    PasangReferensi
End Sub
'-------------------------------------------------------------------------------------------------------------

Private Sub Contoh_SalinBerkasOutlookX()
    
    OutlookX_SalinBerkas "Arsip Email_Vincent", "Vincentius.Budhyanto@generali.co.id", "Deleted Items"
    
End Sub

Function OutlookX_SalinBerkas(ByVal BerkasSumber As String, ByVal BerkasTujuan As String, ByVal Berkas As String)

    BerkasAwal = OutlookX_BuatDaftarBerkas(BerkasSumber, Berkas)
    On Error Resume Next
    For i = 0 To UBound(BerkasAwal)
        OutlookX_TambahBerkas BerkasAwal(i), BerkasTujuan: Debug.Print BerkasAwal(i)
    Next

End Function

'-------------------------------------------------------------------------------------------------------------
Private Sub Contoh_KirimSurelTeksOutlook()
    Outlook_KirimSurelTeks "vincentius.budhyanto@generali.co.id", "Test", "Test", False
End Sub

Function Outlook_KirimSurelTeks(ByVal Tujuan As String, ByVal Judul As String, ByVal IsiSurel As String, ByVal LangsungKirim As Boolean, _
    Optional ByVal NamaLokasiLampiran As String, Optional ByVal Tembusan As String, Optional ByVal TembusanTersembunyi As String)

    Set ProOutlook = CreateObject("Outlook.Application"): ProOutlook.Session.Logon
    Set Surel = ProOutlook.CreateItem(olMailItem)
    
    On Error GoTo Galat
    With Surel
        .To = Tujuan: .Subject = Judul: .CC = Tembusan: .BCC = TembusanTersembunyi
        .Body = IsiSurel
        If NamaLokasiLampiran = "" Then
        Else: .Attachments.Add NamaLokasiLampiran
        End If
        If LangsungKirim = False Then
            .Display
        ElseIf LangsungKirim = True Then
            .Send
        End If
    End With
    
    Debug.Print "BERHASIL TERBENTUK_" & Tujuan & "_" & Judul & "."
    Exit Function

Galat:
    Debug.Print "TERJADI GALAT_" & Tujuan & "_" & Judul & "."
    Exit Function
End Function
'-------------------------------------------------------------------------------------------------------------
        
Private Sub Contoh_TambahBerkasOutlookX()
    OutlookX_TambahBerkas "Inbox/Sapi", "vincentius.budhyanto@generali.co.id"
    OutlookX_TambahBerkas "Inbox/Sapi/Kodok", "Vincentius.Budhyanto@generali.co.id"
    OutlookX_TambahBerkas "Inbox/Sapi/Kodok/Kambing", "vincentius.budhyanto@generali.co.id"
End Sub

Function OutlookX_TambahBerkas(ByVal NamaBerkasBaru As String, ByVal NamaAkunAtauDokumenOST As String)
    On Error GoTo Galat
    
    Set AplOutlook = CreateObject("Outlook.Application")
    Set Penyimpanan = AplOutlook.Session.Folders.Item(NamaAkunAtauDokumenOST)
    
    If InStr(NamaBerkasBaru, "/") > 0 Then GoTo Alur1
    Penyimpanan.Folders.Add NamaBerkasBaru
  
    Debug.Print "BERHASIL MENAMBAH BERKAS: " & NamaBerkasBaru
Exit Function
    
Alur1:
    BerkasBerkas = Split(NamaBerkasBaru, "/")
    Set KotakMasuk = Penyimpanan
    
    For i = 0 To UBound(BerkasBerkas) - 1
        Set KotakMasuk = KotakMasuk.Folders(BerkasBerkas(i))
    Next
    
    Set BerkasDalam = KotakMasuk.Folders
    BerkasDalam.Add BerkasBerkas(UBound(BerkasBerkas))
    Debug.Print "BERHASIL MENAMBAH BERKAS DALAM: " & NamaBerkasBaru
Exit Function

Galat:

    Debug.Print "GALAT FUNGSI: Outlook_TambahBerkas_" & NamaBerkasBaru
Exit Function

End Function
'-------------------------------------------------------------------------------------------------------------

Private Sub Contoh_PeriksaBerkasOutlookX()
    DaftarBerkasSurel = OutlookX_BuatDaftarBerkas("Arsip Email_Vincent", "Inbox")
End Sub

Function OutlookX_DaftarBerkasDalam(ByRef BerkasBerkas As Variant, ByVal NamaBerkasDiperiksa As String, Optional ByRef Hasil, _
    Optional ByRef Lokasi)
    On Error Resume Next
    
    Set BerkasDiperiksa = BerkasBerkas.Folders(NamaBerkasDiperiksa)
    Kembalikan = 1: i = 1: Do Until Kembalikan = 0
        NamaBerkas = BerkasDiperiksa.Folders.Item(i)
        If NamaBerkas = "" Then
            Kembalikan = 1: Exit Do
        Else
            Hasil = Hasil & "<PembatasBerkas>" & Lokasi & "/" & NamaBerkas
            LokasiDiperiksa = Lokasi & "/" & NamaBerkas
        End If: 'Debug.Print Lokasi & "/" & NamaBerkas
        
        Periksa = OutlookX_DaftarBerkasDalam(BerkasDiperiksa, NamaBerkas, Hasil, LokasiDiperiksa)
        If Periksa <> "" Then Hasil = Hasil & "<PembatasBerkas>" & Periksa
        NamaBerkas = ""
    i = i + 1: Loop
    
End Function

Function OutlookX_BuatDaftarBerkas(ByVal NamaAkunAtauDokumenOST As String, Optional ByVal BerkasYangInginDiperiksa As String)
    On Error GoTo Galat

    Set AplOutlook = CreateObject("Outlook.Application")
    Set Penyimpanan = AplOutlook.Session.Folders.Item(NamaAkunAtauDokumenOST)

    If BerkasYangInginDiperiksa = "" Then
        Set BerkasDiperiksa = Penyimpanan
    Else
        Set BerkasBerkas = Penyimpanan.Folders(BerkasYangInginDiperiksa): Set BerkasDiperiksa = BerkasBerkas
    End If

    On Error Resume Next
    Kembalikan = 1: i = 1: Do Until Kembalikan = 0
        
        If BerkasYangInginDiperiksa = "" Then Lokasi = NamaAkunAtauDokumenOST Else Lokasi = BerkasYangInginDiperiksa
        
        NamaBerkas = BerkasDiperiksa.Folders.Item(i)
        If NamaBerkas = "" Then
            Kembalikan = 1: Exit Do
        End If
        
        If i = 1 Then
            Hasil = Lokasi & "/" & NamaBerkas
        ElseIf i > 1 And InStr(Hasil, NamaBerkas) = 0 Then
            Hasil = Hasil & "<PembatasBerkas>" & Lokasi & "/" & NamaBerkas
        End If: 'Debug.Print Lokasi & "/" & NamaBerkas
        Lokasi = Lokasi & "/" & NamaBerkas
        
        Periksa = OutlookX_DaftarBerkasDalam(BerkasDiperiksa, NamaBerkas, Hasil, Lokasi)
        
        NamaBerkas = ""
    i = i + 1: Loop
    On Error GoTo Galat
       
    Hasil = Split(Hasil, "<PembatasBerkas>")
    'For i = 0 To UBound(Hasil): Debug.Print Hasil(i): Next
    
    OutlookX_BuatDaftarBerkas = Hasil
    
Exit Function

Galat:

    Debug.Print "GALAT FUNGSI: OutlookX_PeriksaBerkas"
Exit Function

End Function

