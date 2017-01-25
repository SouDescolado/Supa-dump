Sub BuscaCorreio(repeat As Long)
    '----- Desenvolvido por /u/SouDescolado ------------------------------
    '----- Dúvidas ou sugestões por email wizardoffate@gmail.com ---------
    
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
    ' Cria uma janela no internet explorer
    Set IE = CreateObject("InternetExplorer.Application")
    lr = Cres.Cells(Rows.Count, 1).End(xlUp).Row
    resp = MsgBox("Você quer limpar os resultados?", vbYesNoCancel)
    If resp = vbYes Then
        Cres.Range("A2:E" & lr).Clear
        lr = 1
    ElseIf resp = vbCancel Then
        Exit Sub
    End If
For m = 2 To repeat
    rastreio = Crast.Cells(m, 1).Value
    'Se botar a linha abaixo como comentário dá pra ver o que tá acontecendo no internet explorer
    IE.Visible = False

    'Qual o url do site dos correios que to usando
    IE.Navigate "http://websro.correios.com.br/sro_bin/txect01$.startup?P_LINGUA=001&P_TIPO=001"
    'Coloca situação carregando na barrinha do excel, lá embaixo
    Application.StatusBar = "Carregando"
    'Espera o Internet Explorer carregar (Com negocinho que roda pra ficar bonitinho)
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
    'Põe o rastreio no site e clica buscar
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
    'Espera o Internet Explorer carregar (Com negocinho que roda pra ficar bonitinho)
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
    'Extrai informação da ultima atualização do site.
    lr = lr + 1
    Cres.Cells(lr, 1).Value = rastreio
    Cres.Cells(lr, 2).Value = Crast.Cells(m, 2).Value
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
                            l = 5
                            Cres.Cells(lr, 3).Value = CDate(IE.document.getElementsByTagName("tr")(i).getElementsByTagName("td")(c - 2).innertext)
                            Cres.Cells(lr, 4).Value = stat
                            Cres.Cells(lr, 5).Value = IE.document.getElementsByTagName("tr")(i).getElementsByTagName("td")(c - 1).innertext
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
    'Mostra o Internet Explorer e fecha
    IE.Visible = True
    IE.Quit
    'Limpeza
    Set IE = Nothing
    Application.StatusBar = ""

End Sub
