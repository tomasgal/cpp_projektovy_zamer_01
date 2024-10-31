Private Sub TextBox1_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Dim dateRange As String
    dateRange = Trim(TextBox1.Value)
    
    ' Skontrolujeme, či text zodpovedá formátu ##.##.####-##.##.####
    If dateRange Like "##.##.####-##.##.####" Then
        ' Preformátujeme na ##. ##. #### - ##. ##. ####
        TextBox1.Value = Left(dateRange, 2) & ". " & Mid(dateRange, 4, 2) & ". " & Mid(dateRange, 7, 4) & " - " & _
                         Mid(dateRange, 12, 2) & ". " & Mid(dateRange, 15, 2) & ". " & Right(dateRange, 4)
    Else
        MsgBox "Zadajte dátumové rozpätie vo formáte ##.##.####-##.##.#### bez medzier", vbExclamation
        Cancel = True ' Zostaneme v poli, ak formát nesúhlasí
    End If
End Sub
