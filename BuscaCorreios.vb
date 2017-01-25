Sub BuscaCorreio(repeat As Long)
    Dim Crast As Worksheet
    Set Crast = Sheets("Rastreios Correios")
    Dim Cres As Worksheet
    Set Cres = Sheets("Resultados")
    Dim cod As Worksheet
    Set cod = Sheets("Codigos Correios")
    Dim IE As Object
    Dim i As Long
    Dim j As String
    Dim k As Long
    Dim l As Long
    Dim m As Long
    Dim resp As String
    Dim stat As String
    Dim objCollection As Object
    ' Create InternetExplorer Object
    Set IE = CreateObject("InternetExplorer.Application")
    lr = Cres.Cells(Rows.Count, 1).End(xlUp).Row
    resp = MsgBox("Você quer limpar os resultados?", vbYesNoCancel)
    If resp = vbYes Then
        Cres.Range("A2:E" & lr).Clear
        lr = 1
    ElseIf resp = vbCancel Then
        Exit Sub
    End If
For m = 1 To repeat
    rastreio = Crast.Cells(m, 1).Value
    ' You can uncoment Next line To see form results
    IE.Visible = False

    ' URL to get data from
    IE.Navigate "http://websro.correios.com.br/sro_bin/txect01$.startup?P_LINGUA=001&P_TIPO=001"
    ' Statusbar
    Application.StatusBar = "Carregando"
    ' Wait while IE loading...
    On Error Resume Next
    Do While IE.Busy
        DoEvents
        Application.Wait DateAdd("s", 1, Now)
        i = i + 1
        If i = 1 Then
        j = "/"
        ElseIf i = 2 Then
        j = "-"
        ElseIf i = 3 Then
        j = "\"
        ElseIf i = 4 Then
        j = "|"
        ElseIf i = 5 Then
        i = 1
        j = "/"
        End If
        Application.StatusBar = "Carregando " & j
    Loop
    'Rastreio
    Set objCollection = IE.document.getElementsByTagName("input")
    i = 0
    While i < objCollection.Length
        If objCollection(i).Name = "P_COD_UNI" Then
            objCollection(i).Value = rastreio
        ElseIf objCollection(i).Name = "done" Then
            objCollection(i).Click
            GoTo leave0
        End If
        i = i + 1
    Wend
leave0:
    Do While IE.Busy
        DoEvents
        Application.Wait DateAdd("s", 1, Now)
        i = i + 1
        If i = 1 Then
        j = "/"
        ElseIf i = 2 Then
        j = "-"
        ElseIf i = 3 Then
        j = "\"
        ElseIf i = 4 Then
        j = "|"
        ElseIf i = 5 Then
        i = 1
        j = "/"
        End If
        Application.StatusBar = "Carregando " & j
    Loop
    'Extrai informação
    lr = lr + 1
    Cres.Cells(lr, 1).Value = rastreio
    Set objCollection = IE.document.getElementsByTagName("tr")
    For i = 0 To objCollection.Length - 1
    DoEvents
            tam = IE.document.getElementsByTagName("tr")(i).getElementsByTagName("td").Length
            If tam = 3 Then
                If ult = 1 Then
                    ult = 0
                    Exit For
                End If
                For c = 0 To tam
                DoEvents
                    stat = IE.document.getElementsByTagName("tr")(i).getElementsByTagName("td")(c).innertext
                    For k = 1 To cod.Cells(Rows.Count, 1).End(xlUp).Row
                    DoEvents
                        If stat = cod.Cells(k, 1).Value Then
                            l = 4
                            Cres.Cells(lr, 2).Value = IE.document.getElementsByTagName("tr")(i).getElementsByTagName("td")(c - 2).innertext
                            Cres.Cells(lr, 3).Value = stat
                            Cres.Cells(lr, 4).Value = IE.document.getElementsByTagName("tr")(i).getElementsByTagName("td")(c - 1).innertext
                            ult = 1
                            stat = vbNullString
                            Exit For
                        End If
                    Next k
                Next c
            ElseIf tam = 1 Then
                l = l + 1
                Cres.Cells(lr, l).Value = IE.document.getElementsByTagName("tr")(i).getElementsByTagName("td")(0).innertext
            End If
    Next i
Next m
    

    ' Show IE
    IE.Visible = True
    IE.Quit
    ' Clean up
    Set IE = Nothing
    Application.StatusBar = ""

End Sub


