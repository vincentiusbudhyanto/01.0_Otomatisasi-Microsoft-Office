Function PenghilangNilaiGandaTumpuk2(ByRef BarisanBertumpukDua As Variant) 'BarisanBertumpukDua itu jagged array 2 tingkat
    For i = 0 To UBound(BarisanBertumpukDua): Data = ""
        For i1 = 0 To UBound(BarisanBertumpukDua(i))
            If Data = "" Then Data = BarisanBertumpukDua(i)(i1) Else Data = Data & "|" & BarisanBertumpukDua(i)(i1)
        Next
        If Hasil = "" Then
            Hasil = Data
        ElseIf InStr(Hasil, Data) = 0 Then Hasil = Hasil & "<Hasil>" & Data
        End If: 'Debug.Print Hasil
    Next: Hasil = Split(Hasil, "<Hasil>"): PenghilangNilaiGandaTumpuk2 = Hasil
End Function

Function AmbilAngkaAja(ByVal Tulisan As String)
    For i = 1 To Len(Tulisan)
        If IsNumeric(Mid(Tulisan, i, 1)) Then
            Hasil = Hasil & Mid(Tulisan, i, 1): Aktif = True
        ElseIf Len(Hasil) > 0 And Aktif = True Then AmbilAngkaAja = Hasil: Exit Function
    End If: Next
End Function

Function Tunda(ByVal DurasiDetik As Long)
    Sleep DurasiDetik * 1000
End Function

Function TungguBentar(ByVal Durasi As Integer)
    Dim BerapaLama As Date: BerapaLama = Now + TimeValue("0:00:" & Durasi)
    While Now < BerapaLama: DoEvents: Wend
End Function

Function TungguSiap()
    Do Until Application.Ready: DoEvents: Loop
End Function

Function Bicara(ByVal ApaYangMauDikatakan As String)
    Application.Speech.Speak ApaYangMauDikatakan
End Function

Function ApakahAdaHuruf(ByVal Tulisan As String) As Boolean
    For i = 1 To Len(Tulisan)
        Select Case Asc(Mid(Tulisan, i, 1))
            Case 65 To 90, 97 To 122: ApakahAdaHuruf = True: Exit For
            Case Else: ApakahAdaHuruf = False
        End Select
    Next
End Function

Function HitungDalamTulisan(ByVal YangDicari As String, ByVal Tulisan As String, ByVal Mode As String): i = 0
    Do Until i > Len(Tulisan): Lokasi = 0
        If YangDicari = "" Then Hitung = 0: Exit Do
        If i = 0 Then Lokasi = InStr(Tulisan, YangDicari) Else Lokasi = InStr(i, Tulisan, YangDicari)
        If Lokasi > 0 Then
            Hitung = Hitung + 1: i = Lokasi + Len(YangDicari)
            If CatatLokasi = "" Then CatatLokasi = Lokasi Else CatatLokasi = CatatLokasi & "," & Lokasi
        ElseIf Lokasi = 0 Then
            If Hitung = "" Then Hitung = 0
            Select Case UCase(Mode)
                Case "HITUNG AJA": HitungDalamTulisan = Hitung: Exit Function
                Case "HITUNG LENGKAP": HitungDalamTulisan = Hitung & "|" & CatatLokasi: Exit Function
            End Select
        End If
    Loop
    Select Case UCase(Mode)
        Case "HITUNG AJA": HitungDalamTulisan = Hitung: Exit Function
        Case "HITUNG LENGKAP": HitungDalamTulisan = Hitung & "|" & CatatLokasi: Exit Function
    End Select
End Function

Function UrutanDanJumlah(ByVal KataKataDicari As String, ByRef BarisanData As Variant)
    For i = LBound(BarisanData) To UBound(BarisanData)
        If InStr(BarisanData(i), KataKataDicari) > 0 Then
            i1 = i1 + 1: Jumlah = i1: If Urutan = "" Then Urutan = i Else Urutan = Urutan & "," & i
        End If
    Next: UrutanDanJumlah = Jumlah & "|" & Urutan
End Function

Function SebagianUrutanKeBerapaDiBarisan(ByVal KataKataDicari As String, ByRef BarisanData As Variant)
    For i = LBound(BarisanData) To UBound(BarisanData)
        If InStr(BarisanData(i), KataKataDicari) > 0 Then
            SebagianUrutanKeBerapaDiBarisan = i & "|" & True
            Exit For
        ElseIf BarisanData(i) = Empty Then
            SebagianUrutanKeBerapaDiBarisan = i & "|" & False
            Exit For
        End If
    Next i
End Function

Function UbahJikaDitemukan(ByVal IsiTulisan As String, ByRef BarisanTulisanAwal As Variant, ByRef BarisanTulisanAkhir As Variant)
    For i = 0 To UBound(BarisanTulisanAwal)
        If UCase(IsiTulisan) = UCase(BarisanTulisanAwal(i)) Then
            IsiTulisan = BarisanTulisanAkhir(i): Exit For
        End If
    Next: UbahJikaDitemukan = IsiTulisan
End Function

Function AmbilAngkaDiDepanTulisan(ByVal Tulisan As String) As Integer
    Tulisan = BersihkanAwalAkhirTulisan(Tulisan, " ")
    i = 1: Do While IsNumeric(Mid(Tulisan, i, 1)): Angka = Angka & Mid(Tulisan, i, 1): i = i + 1: Loop: Angka = CInt(Angka)
    AmbilAngkaDiDepanTulisan = Angka: 'Debug.Print Angka
End Function

Function BersihkanAwalAkhirTulisan(ByVal Tulisan As String, ByVal KarakterYangInginDisingkirkan As String)
    Kar = KarakterYangInginDisingkirkan: JumlahKar = Len(Kar): If JumlahKar > 1 Then GoTo Galat
    For i = 1 To Len(Tulisan)
        If Mid(Tulisan, i, 1) = KarakterYangInginDisingkirkan Then TitikMulai = i + 1
        If TitikMulai = "" Then Exit For
        If Mid(Tulisan, TitikMulai, 1) <> KarakterYangInginDisingkirkan Then Exit For
    Next: If TitikMulai <> "" Then Tulisan = Mid(Tulisan, TitikMulai, Len(Tulisan))
    For i = 1 To Len(Tulisan)
        If Mid(Tulisan, Len(Tulisan) - i + 1, 1) = KarakterYangInginDisingkirkan Then TitikAkhir = Len(Tulisan) - i
        If TitikAkhir = "" Then Exit For
        If Mid(Tulisan, TitikAkhir, 1) <> KarakterYangInginDisingkirkan Then Exit For
    Next: If TitikAkhir <> "" Then Tulisan = Mid(Tulisan, 1, TitikAkhir)
    BersihkanAwalAkhirTulisan = Tulisan
Exit Function
Galat: Debug.Print "Terjadi kesalahan pada fungsi -BersihkanAwalAkhirTulisan-."
End Function

Function PeriksaBarisan(ByVal YangDicari As String, ByRef Barisan As Variant) As Boolean
    On Error Resume Next: PeriksaBarisan = (UBound(Filter(Barisan, YangDicari)) > -1): Err.Clear: On Error GoTo 0
End Function

Function UrutanKeBerapaDiBarisan(ByVal KataKataDicari As String, ByRef BarisanData As Variant, Optional ByVal CariKeseluruhanKata As String)
    If CariKeseluruhanKata <> "Tidak" Then
        For i = LBound(BarisanData) To UBound(BarisanData)
            If UCase(KataKataDicari) = UCase(BarisanData(i)) Then
                UrutanKeBerapaDiBarisan = i
                Exit For
            ElseIf BarisanData(i) = Empty Then Exit For
            End If
        Next i
    ElseIf CariKeseluruhanKata = "Tidak" Then
        For i = LBound(BarisanData) To UBound(BarisanData)
            If InStr(UCase(BarisanData(i)), UCase(KataKataDicari)) > 0 Then
                UrutanKeBerapaDiBarisan = i
                Exit For
            ElseIf BarisanData(i) = Empty Then Exit For
            End If
        Next i
    End If
End Function

Function Ambil(ByVal Berkas As String, ByVal NamaLog As String, ByVal KataKunci As String, ByVal Mode As String, Optional ByVal Kondisi)
    'Mode:
    '0 = Akan mengambil yang terakhir, variabel Kondisi disarankan untuk 0
    '1 = Akan mengambil urutan log sesuai variabel Kondisi, Kondisi dalam integer
    '2 = Akan mengambil seluruh log, variabel Kondisi disarankan untuk 0
    '3 = Akan mengambil sesuai dengan rentang, Kondisi dalam string
    '4 = Akan mengambil berdasarkan berapa setelah tanggal, Kondisi dalam string
    '5 = Akan mengambil berdasarkan rentang tanggal sesuai variabel Kondisi, Kondisi dalam string

    Dim LokasiLog As String: Dim MemoriIsiLog As Integer: MemoriIsiLog = FreeFile: Dim BarisIsiLog As String
    Dim UrutanBaris As Integer: Dim PenghitungBaris As Integer: Dim LokasiDitemukan As Integer: Dim Hasil As String
    Dim AmbilUrutan As String: Dim Urutan As Integer: Dim Waktu As String: Dim WaktuBeneran As Date: Dim WaktuPembanding As Date
    Dim Rentang As String: Dim Pembanding As Integer
    Dim FSO As Object: Dim FileBaru As Object

    If Mode = "0" Then Mode = "Terakhir" Else Mode = Mode
    BuatFolderJikaTidakAda ThisWorkbook.Path & "\" & Berkas
    LokasiLog = ThisWorkbook.Path & "\" & Berkas & "\" & NamaLog & ".txt"
    Waktu = TunjukkanWaktu: 'Debug.Print LokasiLog: Debug.Print Dir(LokasiLog)
                                                    
    If Dir(LokasiLog) <> "" Then
        Open LokasiLog For Input As #MemoriIsiLog
    Else
        Set FSO = CreateObject("Scripting.FileSystemObject")
        Set FileBaru = FSO.CreateTextFile(LokasiLog, True)
        FileBaru.WriteLine (Environ("username") & Chr(124) & "Inisiasi Log" & Chr(124) & Waktu)
        FileBaru.Close
        Open LokasiLog For Input As #MemoriIsiLog
    End If

    Waktu = 0: UrutanBaris = 0

    If TypeName(Kondisi) = "Integer" Then Urutan = Kondisi

    Do Until EOF(1)
        Line Input #1, BarisIsiLog: UrutanBaris = UrutanBaris + 1
        If InStr(BarisIsiLog, KataKunci) > 0 And (Mode = "0" Or Mode = "1" Or Mode = "Terakhir") Then
            LokasiDitemukan = UrutanBaris: PenghitungBaris = PenghitungBaris + 1

            If Mode = "1" And Urutan = PenghitungBaris Then AmbilUrutan = BarisIsiLog & "|" & UrutanBaris
            Hasil = BarisIsiLog & "|" & UrutanBaris

        ElseIf InStr(BarisIsiLog, KataKunci) > 0 And Mode = "2" Then
            LokasiDitemukan = UrutanBaris: PenghitungBaris = PenghitungBaris + 1
            Hasil = Hasil & vbCr & BarisIsiLog & "|" & UrutanBaris

        ElseIf InStr(BarisIsiLog, KataKunci) > 0 And Mode = "3" Then
            LokasiDitemukan = UrutanBaris: PenghitungBaris = PenghitungBaris + 1
            Pembanding = CInt(AmbilData(Kondisi, 1, "-"))
            If PenghitungBaris >= Pembanding Then
                Pembanding = CInt(AmbilData(Kondisi, 2, "-"))
                If PenghitungBaris <= Pembanding Then Hasil = Hasil & vbCr & BarisIsiLog & "|" & UrutanBaris
            End If

        ElseIf InStr(BarisIsiLog, KataKunci) > 0 And Mode = "4" Then
            Waktu = AmbilData(BarisIsiLog, 3, Chr(124)): WaktuBeneran = CDate(Left(Waktu, Len(Waktu) - 3))
            WaktuPembanding = CDate(Kondisi)
            If WaktuBeneran >= WaktuPembanding Then Hasil = Hasil & vbCr & BarisIsiLog & "|" & UrutanBaris

        ElseIf InStr(BarisIsiLog, KataKunci) > 0 And Mode = "5" Then
            Waktu = AmbilData(BarisIsiLog, 3, Chr(124)): WaktuBeneran = CDate(Left(Waktu, Len(Waktu) - 3))
            WaktuPembanding = CDate(AmbilData(Kondisi, 1, "-"))
            If WaktuBeneran >= WaktuPembanding Then
                WaktuPembanding = CDate(AmbilData(Kondisi, 2, "-"))
                If WaktuBeneran <= WaktuPembanding Then Hasil = Hasil & vbCr & BarisIsiLog & "|" & UrutanBaris
            End If
                                                                                        
        Else PemicuLogTerakhir = False
        End If
    Loop

    If LokasiDitemukan > 0 Then AmbilUrutan = LokasiDitemukan: Close #MemoriIsiLog
    
    On Error GoTo Galat1
    Select Case Mode
        Case "Terakhir": Ambil = Hasil
        Case "1": Ambil = AmbilUrutan
        Case "2": Ambil = Hasil
        Case "3": Ambil = Hasil
        Case "4": Ambil = Mid(Hasil, 2, Len(Hasil) - 1)
        Case "5": Ambil = Mid(Hasil, 2, Len(Hasil) - 1)
    End Select

    Exit Function

Galat1:
    Debug.Print "Kegalatan pada saat mencetak hasil, mungkin hasil tidak ada."
End Function

Function AmbilDataLog(ByVal NamaLog As String, ByVal KataKunci As String, ByVal Mode As String, ByVal Kondisi)

    'Mode:
    '0 = Akan mengambil yang terakhir, variabel Kondisi disarankan untuk 0
    '1 = Akan mengambil urutan log sesuai variabel Kondisi, Kondisi dalam integer
    '2 = Akan mengambil seluruh log, variabel Kondisi disarankan untuk 0
    '3 = Akan mengambil sesuai dengan rentang, Kondisi dalam string
    '4 = Akan mengambil berdasarkan berapa setelah tanggal, Kondisi dalam string
    '5 = Akan mengambil berdasarkan rentang tanggal sesuai variabel Kondisi, Kondisi dalam string

    Dim LokasiLog As String: Dim MemoriIsiLog As Integer: MemoriIsiLog = FreeFile: Dim BarisIsiLog As String
    Dim UrutanBaris As Integer: Dim PenghitungBaris As Integer: Dim LokasiDitemukan As Integer:
    Dim Hasil As String: Dim AmbilUrutan As String: Dim Urutan As Integer: Dim Waktu As String:
    Dim WaktuBeneran As Date: Dim WaktuPembanding As Date: Dim Rentang As String: Dim Pembanding As Integer

    If Mode = "0" Then Mode = "Terakhir" Else Mode = Mode
    LokasiLog = ThisWorkbook.Path & "\Log\" & NamaLog & ".txt"
    Open LokasiLog For Input As #MemoriIsiLog: UrutanBaris = 0
    If TypeName(Kondisi) = "Integer" Then Urutan = Kondisi

    Do Until EOF(1)
        Line Input #1, BarisIsiLog: UrutanBaris = UrutanBaris + 1
        If InStr(BarisIsiLog, KataKunci) > 0 And (Mode = "0" Or Mode = "1" Or Mode = "Terakhir") Then
            PemicuLogTerakhir = True: LokasiDitemukan = UrutanBaris: PenghitungBaris = PenghitungBaris + 1
            If Mode = "1" And Urutan = PenghitungBaris Then
                AmbilUrutan = BarisIsiLog & "|" & UrutanBaris
            End If: Hasil = BarisIsiLog & "|" & UrutanBaris

        ElseIf InStr(BarisIsiLog, KataKunci) > 0 And Mode = "2" Then
            PemicuLogTerakhir = True: LokasiDitemukan = UrutanBaris: PenghitungBaris = PenghitungBaris + 1
            Hasil = Hasil & vbCr & BarisIsiLog & "|" & UrutanBaris

        ElseIf InStr(BarisIsiLog, KataKunci) > 0 And Mode = "3" Then
            PemicuLogTerakhir = True: LokasiDitemukan = UrutanBaris: PenghitungBaris = PenghitungBaris + 1
            Pembanding = CInt(AmbilData(Kondisi, 1, "-"))
            If PenghitungBaris >= Pembanding Then
                Pembanding = CInt(AmbilData(Kondisi, 2, "-"))
                If PenghitungBaris <= Pembanding Then Hasil = Hasil & vbCr & BarisIsiLog & "|" & UrutanBaris
            End If

        ElseIf InStr(BarisIsiLog, KataKunci) > 0 And Mode = "4" Then
            PemicuLogTerakhir = True: Waktu = AmbilData(BarisIsiLog, 3, Chr(124)): WaktuBeneran = CDate(Left(Waktu, Len(Waktu) - 3))
            WaktuPembanding = CDate(Kondisi)
            If WaktuBeneran >= WaktuPembanding Then Hasil = Hasil & vbCr & BarisIsiLog & "|" & UrutanBaris

        ElseIf InStr(BarisIsiLog, KataKunci) > 0 And Mode = "5" Then
            PemicuLogTerakhir = True: Waktu = AmbilData(BarisIsiLog, 3, Chr(124)): WaktuBeneran = CDate(Left(Waktu, Len(Waktu) - 3))
            WaktuPembanding = CDate(AmbilData(Kondisi, 1, "-"))
            If WaktuBeneran >= WaktuPembanding Then
                WaktuPembanding = CDate(AmbilData(Kondisi, 2, "-"))
                If WaktuBeneran <= WaktuPembanding Then Hasil = Hasil & vbCr & BarisIsiLog & "|" & UrutanBaris
            End If
                                                                                                                    
        Else: PemicuLogTerakhir = False
        End If
    Loop: Close #MemoriIsiLog

    Select Case Mode
        Case "Terakhir": AmbilDataLog = Hasil
        Case "1": AmbilDataLog = AmbilUrutan
        Case "2": AmbilDataLog = Hasil
        Case "3": AmbilDataLog = Hasil
        Case "4": AmbilDataLog = Mid(Hasil, 2, Len(Hasil) - 1)
        Case "5": AmbilDataLog = Mid(Hasil, 2, Len(Hasil) - 1)
    End Select
End Function

Function AmbilData(ByVal Data As String, ByVal UrutanKe As Integer, ByVal Pembagi As String)
    Dim Membagi() As String: On Error GoTo Galat1: Membagi() = Split(Data, Pembagi)
                                                                                                            
    If Data = "" Then Debug.Print "Data tidak ditemukan":Exit Function
    UrutanKe = UrutanKe - 1: AmbilData = Membagi(UrutanKe)        
                                                                                                                    
    Exit Function
    Galat1: Debug.Print "Terjadi kesalahan pada fungsi AmbilData"
End Function

Function CatatLokal(ByVal YangDiCatat As String, ByVal NamaLog As String, ByVal NamaBerkas As String)
    Dim FSO As Object: Dim LokasiKerja As String: Dim NamaFile As String: Dim LokFile As String: Dim FileBaru As Object
    Dim RekamJejak As String
    Set FSO = CreateObject("Scripting.FileSystemObject")

    NamaFile = NamaLog: LokasiKerja = ThisWorkbook.Path & "\": CekFolder = LokasiKerja & NamaBerkas: BuatFolderJikaTidakAda (CekFolder)
    RekamJejak = CekFolder & "\" & NamaFile & ".txt": Waktu = TunjukkanWaktu

    If Dir(RekamJejak) <> "" Then
        Set FileBaru = FSO.OpenTextFile(RekamJejak, ForAppending)
        FileBaru.WriteLine (Environ("username") & Chr(124) & YangDiCatat & Chr(124) & Waktu): FileBaru.Close
    Else
        Set FileBaru = FSO.CreateTextFile(RekamJejak, True)
        FileBaru.WriteLine (Environ("username") & Chr(124) & "Inisiasi Log" & Chr(124) & Waktu): Waktu = TunjukkanWaktu
        FileBaru.WriteLine (Environ("username") & Chr(124) & YangDiCatat & Chr(124) & Waktu): FileBaru.Close
    End If
End Function

Function CatatData(ByVal YangDiCatat As String, ByVal NamaLog As String)
    Dim FSO As Object: Dim LokasiKerja As String: Dim NamaFile As String: Dim LokFile As String: Dim FileBaru As Object:
    Dim RekamJejak As String: Set FSO = CreateObject("Scripting.FileSystemObject")

    NamaFile = NamaLog: LokasiKerja = ThisWorkbook.Path & "\": CekFolder = LokasiKerja & "Data"
    BuatFolderJikaTidakAda (CekFolder): RekamJejak = CekFolder & "\" & NamaFile & ".txt": Waktu = TunjukkanWaktu

    If Dir(RekamJejak) <> "" Then
        Set FileBaru = FSO.OpenTextFile(RekamJejak, ForAppending)
        FileBaru.WriteLine (Environ("username") & Chr(124) & YangDiCatat & Chr(124) & Waktu): FileBaru.Close
    Else
        Set FileBaru = FSO.CreateTextFile(RekamJejak, True)
        FileBaru.WriteLine (Environ("username") & Chr(124) & "Inisiasi Log" & Chr(124) & Waktu): Waktu = TunjukkanWaktu
        FileBaru.WriteLine (Environ("username") & Chr(124) & YangDiCatat & Chr(124) & Waktu): FileBaru.Close
    End If
End Function

Function CatatProsedur(ByVal YangDiCatat As String, ByVal NamaLog As String)
    Dim FSO As Object: Dim LokasiKerja As String: Dim NamaFile As String: Dim LokFile As String: Dim FileBaru As Object:
    Dim RekamJejak As String: Set FSO = CreateObject("Scripting.FileSystemObject")

    NamaFile = NamaLog: LokasiKerja = ThisWorkbook.Path & "\": CekFolder = LokasiKerja & "Prosedur"
    BuatFolderJikaTidakAda (CekFolder)
    RekamJejak = CekFolder & "\" & NamaFile & ".txt": Waktu = TunjukkanWaktu

    If Dir(RekamJejak) <> "" Then
        Set FileBaru = FSO.OpenTextFile(RekamJejak, ForAppending)
        FileBaru.WriteLine (Environ("username") & Chr(124) & YangDiCatat & Chr(124) & Waktu): FileBaru.Close
    Else
        Set FileBaru = FSO.CreateTextFile(RekamJejak, True)
        FileBaru.WriteLine (Environ("username") & Chr(124) & "Inisiasi Log" & Chr(124) & Waktu): Waktu = TunjukkanWaktu
        FileBaru.WriteLine (Environ("username") & Chr(124) & YangDiCatat & Chr(124) & Waktu): FileBaru.Close
    End If
End Function


Function Catat(ByVal YangDiCatat As String, ByVal NamaFile As String)                                                                    
    Dim FSO As Object: Dim LokasiKerja As String: Dim LokFile As String: Dim FileBaru As Object: Dim RekamJejak As String
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    LokasiKerja = ThisWorkbook.Path & "\": CekFolder = LokasiKerja & "Log": BuatFolderJikaTidakAda (CekFolder)
    RekamJejak = CekFolder & "\" & NamaFile & "_" & Format(Now, "yyyymm") & ".txt": Waktu = TunjukkanWaktu

    If Dir(RekamJejak) <> "" Then
        Set FileBaru = FSO.OpenTextFile(RekamJejak, ForAppending)
        FileBaru.WriteLine (Environ("username") & Chr(124) & YangDiCatat & Chr(124) & Waktu)
        FileBaru.Close
    Else
        Set FileBaru = FSO.CreateTextFile(RekamJejak, True)
        FileBaru.WriteLine (Environ("username") & Chr(124) & "Inisiasi " & NamaFile & Chr(124) & Waktu)
        Waktu = TunjukkanWaktu
        FileBaru.WriteLine (Environ("username") & Chr(124) & YangDiCatat & Chr(124) & Waktu)
        FileBaru.Close
    End If
End Function

Function BuatFolderJikaTidakAda(ByVal LokasiFolder As String)
    Dim FSO As Object: Set FSO = CreateObject("Scripting.FileSystemObject")
    If Not (FSO.FolderExists(LokasiFolder)) Then MkDir (LokasiFolder)
    BuatFolderJikaTidakAda = "Sudah diproses"
End Function

Function ApakahAdaAngka(ByVal Data As String) As Boolean
    BarisanAngka = Array("0", "1", "2", "3", "4", "5", "6", "7", "8", "9"): ApakahAdaAngka = False
    For i = 0 To UBound(BarisanAngka)
        If InStr(Data, BarisanAngka(i)) > 0 Then ApakahAdaAngka = True: Exit Function
    Next
End Function

Function TunjukkanWaktu()
    Waktu = Strings.Format(Now, "dd-MMM-yyyy hh:mm:ss") & "." & Strings.Right(Strings.Format(Timer, "#0.00"), 2)
    TunjukkanWaktu = Waktu: 'Debug.Print Waktu
End Function

Function PenghilangNilaiGandaTumpuk2(ByRef BarisanBertumpukDua As Variant)
    For i = 0 To UBound(BarisanBertumpukDua): Data = ""
        For i1 = 0 To UBound(BarisanBertumpukDua(i))
            If Data = "" Then Data = BarisanBertumpukDua(i)(i1) Else Data = Data & "|" & BarisanBertumpukDua(i)(i1)
        Next
        If Hasil = "" Then
            Hasil = Data
        ElseIf InStr(Hasil, Data) = 0 Then Hasil = Hasil & "<Hasil>" & Data
        End If: 'Debug.Print Hasil
    Next: Hasil = Split(Hasil, "<Hasil>"): PenghilangNilaiGandaTumpuk2 = Hasil
End Function                                                                                                                                                  
                                                                                                                                                    
