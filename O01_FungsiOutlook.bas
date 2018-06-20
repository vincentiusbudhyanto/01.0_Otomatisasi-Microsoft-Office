Private Sub ContohPenggunaanKirimSurelTeks()
    KirimSurelTeks "vincentius.budhyanto@generali.co.id", "Test", "test", False
End Sub

Private Function KirimSurelTeks(ByVal Tujuan As String, ByVal Judul As String, ByVal IsiSurel As String, ByVal LangsungKirim As Boolean, _
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
