'Panggil referensi yang ingin di pasang. 
'Contohnya sebagai berikut:

'------------------
Sub PasangReferensi()
    'Call PasangRefScripting: PasangRefOutlook
End Sub
'------------------


Private Sub PasangRefScripting()
    LokasiSistem = AmbilLokasiSistem: If Not AdakahRef("Scripting", "Nama") Then TambahRef LokasiSistem & "\scrrun.dll"
End Sub

Private Sub PasangRefOutlook()
    NamaAplikasi = Application.Name: DataRef = BuatDaftarReferensi
    For i = 0 To UBound(DataRef)
        If InStr(DataRef(i)(1), NamaAplikasi) > 0 Then
            LokasiProgram = DataRef(i)(2): BatasBerkas = LokTerakhirKarakter("\", LokasiProgram)
            LokasiProgram = Left(LokasiProgram, BatasBerkas - 1): 'Debug.Print LokasiProgram
    End If: Next
    If Not AdakahRef("Microsoft Outlook", "Deskripsi") Then TambahRef LokasiProgram & "\MSOUTL.OLB": GoTo CobaVersiLain
    Exit Sub
CobaVersiLain:
    Debug.Print "Mencoba pada versi yang lebih rendah."
    If InStr(UCase(LokasiProgram), "OFFICE15") > 0 Then
        LokasiProgram = Replace(UCase(LokasiProgram), "OFFICE15", "OFFICE14")
    ElseIf InStr(UCase(LokasiProgram), "OFFICE16") > 0 Then LokasiProgram = Replace(UCase(LokasiProgram), "OFFICE16", "OFFICE15")
    End If
    If Not AdakahRef("Microsoft Outlook", "Deskripsi") Then TambahRef LokasiProgram & "\MSOUTL.OLB"
End Sub

Private Function LokTerakhirKarakter(ByVal Karakter As String, ByVal Tulisan As String)
    Karakter = Left(Karakter, 1)
    For i = 1 To Len(Tulisan)
        If Mid(Tulisan, i, 1) = Karakter Then Hasil = i
    Next: LokTerakhirKarakter = Hasil
End Function

Private Function BuatDaftarReferensi()
    For Each Ref In Application.VBE.ActiveVBProject.References
        If BarisRef = "" Then BarisRef = Ref.Name & "<>" & Ref.Description & "<>" & Ref.FullPath Else BarisRef = BarisRef & "<BarisRef>" & _
            Ref.Name & "<>" & Ref.Description & "<>" & Ref.FullPath
        i = i + 1: 'Debug.Print "Referensi " & i & " -> " & Ref.Name & " / " & Ref.Description & " / " & Ref.FullPath
    Next Ref: BarisRef = Split(BarisRef, "<BarisRef>")
    ReDim Hasil(UBound(BarisRef)): For i = 0 To UBound(Hasil): Hasil(i) = Split(BarisRef(i), "<>"): Next: BuatDaftarReferensi = Hasil
End Function

Private Function AdakahRef(ByVal NamaReferensi As String, ByVal NamaAtauDeskripsi As String) As Boolean
    RefTerdaftar = BuatDaftarReferensi: If InStr(UCase(NamaAtauDeskripsi), "NAMA") > 0 Then j = 0 Else j = 1
    For i = 0 To UBound(RefTerdaftar)
        If j = 0 Then
            If RefTerdaftar(i)(j) = NamaReferensi Then
                AdakahRef = True: Exit Function
            Else: AdakahRef = False
            End If
        ElseIf j = 1 Then
            If InStr(RefTerdaftar(i)(j), NamaReferensi) > 0 Then
                AdakahRef = True: Exit Function
            Else: AdakahRef = False
    End If: End If: Next
End Function

Private Function TambahRef(ByVal NamLokRef As String)
    On Error GoTo TidakBisa: Application.VBE.ActiveVBProject.References.AddFromFile NamLokRef:
    Debug.Print "Berhasil menambahkan Referensi " & NamLokRef: Exit Function
TidakBisa: Debug.Print "Tidak bisa menambahkan Referensi " & NamLokRef
End Function

Private Function AmbilLokasiSistem() As String
    Set JenisSistem = GetObject("winmgmts:root\cimv2:Win32_Processor='cpu0'"): Sistem = JenisSistem.Architecture
    If Sistem = 9 Then
        PeriksaJenisSistem = "x64"
    ElseIf Sistem = 0 Then PeriksaJenisSistem = "x32"
    End If: 'Debug.Print PeriksaJenisSistem
    If PeriksaJenisSistem = "x64" Then
        LokasiSistem = "C:\Windows\SysWOW64"
    ElseIf PeriksaJenisSistem = "x32" Then LokasiSistem = "C:\Windows\system32"
    End If: AmbilLokasiSistem = LokasiSistem
End Function
