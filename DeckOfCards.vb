Public deck As New Collection
Public face As New Collection
Public symbols As New Collection

Sub Cards()

face.Add "A"
For i = 2 To 10
    face.Add i
Next i
face.Add "J"
face.Add "Q"
face.Add "K"
symbols.Add ChrW(9826)
symbols.Add ChrW(9828)
symbols.Add ChrW(9825)
symbols.Add ChrW(9831)

End Sub

Sub Restart()

Set deck = Nothing
Set face = Nothing
Set symbols = Nothing
Dim i As Long
For i = 1 To 52
    deck.Add i
Next i
Cards
        
End Sub

Sub draw(draw As Long)
        
Dim randc As Long
Dim upperbound As Long
Dim cardnumber As Long
Dim actualcard As String
Dim i As Long
Dim lastrow As Long
lastrow = Cells(Rows.count, 14).End(xlUp).row
For i = 1 To draw
    upperbound = deck.count
    If upperbound = 0 Then
        MsgBox "Out of cards"
    Exit Sub
    End If
    randc = Int((upperbound - 1 + 1) * Rnd() + 1)
    If randc > upperbound Then
       randc = upperbound
    End If
    cardnumber = deck.Item(randc)
    deck.Remove (randc)
    actualcard = face.Item(cardnumber - (WorksheetFunction.RoundUp(cardnumber / 13, 0) - 1) * 13) & symbols.Item(WorksheetFunction.RoundUp(cardnumber / 13, 0))
    lastrow = lastrow + 1
    Cells(lastrow, 14).Value = actualcard
Next i

End Sub
