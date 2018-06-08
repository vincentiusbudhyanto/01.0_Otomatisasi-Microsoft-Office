Function PemicuWaktuPasti(ByVal Mode As String, ByVal KondisiWaktu As String, ByVal KapanMulai As String, _
    ByVal ProsedurApa As String, ByVal KataKunci As String)

    'Mode:
    'Harian = Dijalankan tiap hari
    'Mingguan = Dijalankan tiap minggu
    'Bulanan = Dijalankan tiap bulan
    'Tahunan = Dijalankan tiap tahun
    'PerJam = Dijalankan tiap jam

    Dim PatokanMulai As Date
    Dim WaktuPatokan As Date
    Dim Sekarang As Date

    Dim Status As Boolean

    Dim Jam As Integer: Jam = Hour(Now)
    Dim Menit As Integer: Menit = Minute(Now)
    Dim Detik As Integer: Detik = Second(Now)

    Dim WaktuTerakhirJalanstr As String
    Dim WaktuTerakhirJalan As Date

    WaktuTerakhirJalanstr = AmbilData(Ambil("Prosedur", ProsedurApa, KataKunci, 0, 0), 3, "|")

    If WaktuTerakhirJalanstr = "" Then
        WaktuTerakhirJalan = 0
    '    PemicuWaktuPasti = True
    '    Exit Function
    Else
        WaktuTerakhirJalanstr = Left(WaktuTerakhirJalanstr, Len(WaktuTerakhirJalanstr) - 3)
    End If

    WaktuTerakhirJalan = CDate(WaktuTerakhirJalanstr)

    PatokanMulai = CDate(KapanMulai)
    WaktuPatokan = CDate(KondisiWaktu)
    Sekarang = CDate(Jam & ":" & Menit & ":" & Detik)

    Select Case Mode
    Case "PerJam"
        If Now > PatokanMulai And Sekarang > WaktuPatokan And (Hour(WaktuTerakhirJalan) + (1 / (24))) < Hour(Now) Then
            PemicuWaktuPasti = True
            Exit Function
        End If

    Case "Harian"
        If Now > PatokanMulai And Sekarang > WaktuPatokan And Day(WaktuTerakhirJalan) <> Day(Now) Then
            PemicuWaktuPasti = True
            Exit Function
        End If

    Case "Mingguan"
        If WaktuTerakhirJalan = 0 Then WaktuTerakhirJalanstr = -7
        If Now > PatokanMulai And Now >= (WaktuTerakhirJalan + 7) And Sekarang > WaktuPatokan And (WaktuTerakhirJalan + 6) <= Now _
            And Weekday(WaktuTerakhirJalan) = Weekday(Now) Then
            PemicuWaktuPasti = True
            Exit Function
        End If
    Case "Bulanan"
        If Now > PatokanMulai And Sekarang > WaktuPatokan And Month(WaktuTerakhirJalan) <> Month(Now) And DatePart("d", WaktuTerakhirJalanstr) = DatePart("d", Now) Then
            PemicuWaktuPasti = True
            Exit Function
        End If

    Case "Tahunan"
        If Now > PatokanMulai And Sekarang > WaktuPatokan And (Month(WaktuTerakhirJalan) + 1) <= Month(Now) And _
            Day(WaktuTerakhirJalan) <= Day(Now) And (Year(WaktuTerakhirJalan) + 1) <= Year(Now) Then
            PemicuWaktuPasti = True
            Exit Function
        End If

    End Select

    Debug.Print "Prosedur tidak dijalankan, pemicu tidak aktif."
    PemicuWaktuPasti = False

End Function

Private Function PemicuSurel(ByVal NamaAkun As String, ByVal Pengirim As String, ByVal KataKunci As String, _
    ByVal BerapaHariKebelakang As Double)

    Dim ProgramOutlook As Outlook.Application
    Dim Akun As Object
    Dim PakaiAkun As Object
    Dim Ruang As Namespace
    Dim BerkasUtama As Outlook.MAPIFolder
    Dim Berkas As Object

    Dim IsiBerkas(0 To 1000) As Variant
    Dim i As Integer

    Set ProgramOutlook = CreateObject("Outlook.Application")
    Set Akun = ProgramOutlook.Session.Accounts
    Set Ruang = GetNamespace("MAPI")

    For Each Item In ProgramOutlook.Session.Accounts
        Set PakaiAkun = Item
        If PakaiAkun = NamaAkun Then
            Set BerkasUtama = Ruang.GetDefaultFolder(olFolderInbox)
        End If
    Next

    For Each Item In BerkasUtama.Folders
        IsiBerkas(i) = Item.Name
        i = i + 1
    Next

    i = 1

    Do While IsiBerkas(i) <> ""
        Set Berkas = BerkasUtama.Folders(IsiBerkas(i))
        For Each Item In Berkas.Items
            If (TypeOf Item Is Outlook.MailItem) Then
                Set Surel = Item
                If Surel.SenderEmailAddress = Pengirim And InStr(Surel.Subject, KataKunci) > 0 And (Now - Surel.ReceivedTime < BerapaHariKebelakang) Then

                   Debug.Print Surel.SenderEmailAddress & "|" & Surel.Subject & "|" & Surel.ReceivedTime 'Masukkan kode ekspor disini

                End If
            End If
        Next
        i = i + 1
    Loop

    Set ProgramOutlook = Nothing
    Set Akun = Nothing
    Set PakaiAkun = Nothing
    Set Ruang = Nothing
    Set BerkasUtama = Nothing
    Set Berkas = Nothing

    Erase IsiBerkas
    i = 0

    End Function

Function PemicuWaktuKadaluarsa(ByVal NamLokDokumen As String, ByVal BatasWaktu As Integer)

    Dim Sistem As Object
    Dim Dokumen As Object
    Dim TanggalModif As Date

    Set Sistem = CreateObject("Scripting.FileSystemObject")
    Set Dokumen = Sistem.GetFile(NamLokDokumen)
    TanggalModif = Dokumen.DateLastModified

    If (TanggalModif + BatasWaktu) < Now Then
        PemicuWaktuKadaluarsa = True
    Else
        PemicuWaktuKadaluarsa = False
    End If

    Set Sistem = Nothing
    Set Dokumen = Nothing
    TanggalModif = 0
    BatasWaktu = 0

End Function

'----------------------------------------------------------------------------------------------------------

Function PemicuWaktuTertentu(ByVal NamLokDokumen As String, ByVal BatasWaktu As String)

    BatasWaktu = CDate(BatasWaktu)

    If BatasWaktu < Now Then
        PemicuWaktuTertentu = True
    Else
        PemicuWaktuTertentu = False
    End If

    NamLokDokumen = vbNullString
    BatasWaktu = 0

End Function

'----------------------------------------------------------------------------------------------------------

Function PemicuWaktuRutin(ByVal NamLokDokumen As String, ByVal BerapaLama As Double, _
    ByVal KapanMulai As String)

    Dim Sistem As Object
    Dim Dokumen As Object
    Dim TanggalModif As Date
    Dim BatasMulai As Double

    Set Sistem = CreateObject("Scripting.FileSystemObject")
    Set Dokumen = Sistem.GetFile(NamLokDokumen)
    TanggalModif = Dokumen.DateLastModified

    BatasMulai = TanggalModif + BerapaLama
    BatasMulai = CDate(BatasMulai)
    KapanMulai = CDate(KapanMulai)

    If BatasMulai < Now And KapanMulai < Now Then
        PemicuWaktuRutin = True
    Else
        PemicuWaktuRutin = False
    End If

    Set Sistem = Nothing
    Set Dokumen = Nothing
    TanggalModif = 0
    BatasMulai = 0
    BerapaLama = 0

End Function

'----------------------------------------------------------------------------------------------------------

Function PeriksaBerkas(ByVal KodeAPI As String, ByVal Ekstensi As String, ByVal SumberBerkas As String)

    Dim Sistem As Object
    Dim Berkas As Object
    Dim Dokumen As Object

    Set Sistem = CreateObject("Scripting.FileSystemObject")
    Debug.Print SumberBerkas
    Set Berkas = Sistem.GetFolder(SumberBerkas)

    For Each Dokumen In Berkas.Files

        If InStr(Dokumen.Name, KodeAPI) And (Ekstensi = Right(Dokumen.Name, Len(Ekstensi))) And InStr(Dokumen.Name, "Selesai") = 0 Then
            PeriksaBerkas = True
            Exit Function
        End If

    Next Dokumen

    Set Sistem = Nothing
    Set BukaBerkas = Nothing

    PeriksaBerkas = False

End Function
