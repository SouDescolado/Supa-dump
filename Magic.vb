Option Explicit
Public vez As Long
Public grupo As New Collection

Public Sub magic(ByVal coluna As Long)
    Dim row As Long
    Dim col As Long
    Dim i As Long
    Dim j As Long
    Dim order As New Collection
    Dim g1 As New Collection
    Dim g2 As New Collection
    Dim g3 As New Collection
    Dim rand As New Collection
    Dim ub As Long
    Dim aleat As Long
    Dim check As String
    order.Add g1
    order.Add g2
    order.Add g3
    order.Add g1
    order.Add g2
    If vez < 2 Then
        If grupo.Count < 1 Then
            For row = 1 To 7
                g1.Add Cells(row, 1).Value
                g2.Add Cells(row, 2).Value
                g3.Add Cells(row, 3).Value
            Next row
            Do While order.Item(2).Item(1) <> Cells(1, coluna)
                order.Remove (1)
            Loop
            Do While order.Count > 3
                order.Remove (4)
            Loop
        Else
            check = grupo.Item(coluna)
            Do While grupo.Count > 0
                g1.Add grupo.Item(1)
                g2.Add grupo.Item(2)
                g3.Add grupo.Item(3)
                For i = 1 To 3
                    grupo.Remove (1)
                Next i
            Loop
            Do While order.Item(2).Item(1) <> check
                order.Remove (1)
            Loop
            Do While order.Count > 3
                order.Remove (4)
            Loop
        End If
        For i = 1 To 3
            For j = 1 To 7
                grupo.Add order.Item(i).Item(j)
            Next j
        Next i
        i = 1

        For row = 1 To 7
            For col = 1 To 3
                Cells(row, col).Value = grupo.Item(i)
                i = i + 1
            Next col
        Next row
        For col = 1 To 3
            Set rand = Nothing
            For row = 1 To 7
                rand.Add Cells(row, col).Value
            Next row
            For row = 1 To 7
                ub = rand.Count
                aleat = Int((ub - 1 + 1) * Rnd) + 1
                If aleat > ub Then
                    aleat = ub
                End If
                Cells(row, col).Value = rand.Item(aleat)
                rand.Remove (aleat)
            Next row
        Next col
        vez = vez + 1
    ElseIf vez = 2 Then
        check = grupo.Item(coluna)
        Do While grupo.Count > 0
            g1.Add grupo.Item(1)
            g2.Add grupo.Item(2)
            g3.Add grupo.Item(3)
            For i = 1 To 3
                grupo.Remove (1)
            Next i
        Loop
        Do While order.Item(2).Item(1) <> check
            order.Remove (1)
        Loop
        Do While order.Count > 3
            order.Remove (4)
        Loop
        For i = 1 To 3
            For j = 1 To 7
                grupo.Add order.Item(i).Item(j)
            Next j
        Next i
        Cells(5, 6) = grupo.Item(11)
    End If
End Sub

Private Sub CommandButton1_Click()
    magic (1)
End Sub
Private Sub CommandButton2_Click()
    magic (2)
End Sub
Private Sub CommandButton3_Click()
    magic (3)
End Sub


Private Sub CommandButton4_Click()
    Dim row As Long
    Dim col As Long
    Dim aleat As Long
    Dim col2 As Long
    Dim row2 As Long
    Dim upperbound As Long
    Dim lowerbound As Long
    upperbound = 24
    lowerbound = 1
    For row = 1 To 7
        For col = 1 To 3
            Cells(row, col) = vbNullString
        Next col
    Next row
    Set grupo = Nothing
    Cells(5, 6) = vbNullString
    vez = 0
    For row = 1 To 7
        For col = 1 To 3
Reset:
            aleat = Int((upperbound - lowerbound + 1) * Rnd()) + lowerbound
            For row2 = 1 To 7
                For col2 = 1 To 3
                    If Cells(row2, col2) = Chr(aleat + 65) Then
                        GoTo Reset
                    End If
                Next col2
            Next row2
            Cells(row, col) = Chr(aleat + 65)
        Next col
    Next row
End Sub