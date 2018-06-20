Private Sub InsiasiModulOutlook()
    PasangReferensi
End Sub
'-------------------------------------------------------------------------------------------------------------

Private Sub Contoh_KirimSurelTeksOutlook()
    Outlook_KirimSurelTeks "vincentius.budhyanto@generali.co.id", "Test", "test", False
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
        
Private Sub Contoh_TambahBerkasOutlook()
    Outlook_TambahBerkas "Sapi": Outlook_TambahBerkas "Sapi/Kodok": Outlook_TambahBerkas "Sapi/Kodok/Kambing"
End Sub

Function Outlook_TambahBerkas(ByVal NamaBerkasBaru As String)
    On Error GoTo Galat
    
    Set AplOutlook = CreateObject("Outlook.Application"): AplOutlook.Session.Logon
    Set Penyimpanan = AplOutlook.GetNamespace("MAPI")
    Set KotakMasukUtama = Penyimpanan.GetDefaultFolder(olFolderInbox)
    Set SurelSurel = Penyimpanan.GetDefaultFolder(olFolderInbox).Items
    Set KotakKotakMasuk = Penyimpanan.GetDefaultFolder(olFolderInbox).Folders
    
    If InStr(NamaBerkasBaru, "/") > 0 Then GoTo Alur1
    KotakMasukUtama.Folders.Add NamaBerkasBaru
  
    Debug.Print "BERHASIL MENAMBAH BERKAS: " & NamaBerkasBaru
Exit Function
    
Alur1:
    BerkasBerkas = Split(NamaBerkasBaru, "/")
    Set KotakMasuk = KotakMasukUtama
    
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

