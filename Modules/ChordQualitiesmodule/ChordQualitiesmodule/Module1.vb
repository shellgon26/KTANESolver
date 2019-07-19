Module Module1
    Structure Toffsets
        Dim offset1, offset2, offset3, offset4, offset5 As Integer
    End Structure
    Structure tnotes
        Dim note1, note2, note3, note4 As String
    End Structure
    Structure TChord
        Dim offsets As Toffsets
        Dim quality As String
        Dim note As tnotes
    End Structure
    Sub Main()
        Dim notes() As String = {"A", "A#", "B", "C", "C#", "D", "D#", "E", "F", "F#", "G", "G#"}
        Dim chords(11) As TChord
        initializeoffsets(chords)
        Dim chord1 As TChord
        inputstartchord(chord1)
        chordoffsets(chord1, notes)
        Dim startquality = findstartquality(chord1)
        displayoffsets(chord1)
        Dim startandquality = findquality(chord1, chords)
        Dim startnote As String = Mid(startandquality, 1, 1)
        Dim quality As String = Mid(startandquality, 2)
        Console.WriteLine(startnote & ", " & quality)
        Console.ReadLine()
    End Sub
    Sub displayoffsets(ByVal chord As TChord)
        Console.WriteLine(chord.offsets.offset1 & ", " & chord.offsets.offset2 & ", " & chord.offsets.offset3 & ", " & chord.offsets.offset4)
    End Sub
    Sub chordoffsets(ByRef chord As TChord, ByVal notes() As String)
        Dim foundnotes As Integer = 0
        Dim offset As Integer = 0
        Dim startpoint = findstartchord(chord.note.note1, notes)
        chord.offsets.offset1 = 0
        If CStr(startpoint) = "Error" Then
            Console.WriteLine("Input Error")
            inputstartchord(chord)
            findstartchord(chord.note.note1, notes)
        End If
        Do
            For j = startpoint To startpoint + 11
                Dim checker = j Mod 12
                Select Case foundnotes
                    Case 0
                        If chord.note.note2 = notes(checker) Then
                            chord.offsets.offset2 = offset
                            foundnotes = foundnotes + 1
                        End If
                    Case 1
                        If chord.note.note3 = notes(checker) Then
                            chord.offsets.offset3 = offset
                            foundnotes = foundnotes + 1
                        End If
                    Case 2
                        If chord.note.note4 = notes(checker) Then
                            chord.offsets.offset4 = offset
                            chord.offsets.offset5 = 12 - offset
                            foundnotes = foundnotes + 1
                        End If
                End Select
                offset = offset + 1
            Next
        Loop Until foundnotes = 3

    End Sub
    Function findstartchord(ByVal note As String, ByVal notes() As String)
        For i = 0 To 11
            If note = notes(i) Then
                Return i
            End If
        Next
        Return "Did not enter a valid note"
    End Function
    Function findstartquality(ByVal chord As TChord)
        Select Case chord.note.note1
            Case "A"
                Return 0
            Case "A#"
                Return 1
            Case "B"
                Return 2
            Case "C"
                Return 3
            Case "C#"
                Return 4
            Case "D"
                Return 5
            Case "D#"
                Return 6
            Case "E"
                Return 7
            Case "F"
                Return 8
            Case "F#"
                Return 9
            Case "G"
                Return 10
            Case "G#"
                Return 11
            Case Else
                Return "Error"
        End Select
    End Function
    Sub inputstartchord(ByRef chord As TChord)
        Dim temp As String
        For i = 0 To 3
            Console.Write("Enter note " & i + 1 & ": ")
            temp = UCase(Console.ReadLine)
            Select Case i
                Case 0
                    chord.note.note1 = temp
                Case 1
                    chord.note.note2 = temp
                Case 2
                    chord.note.note3 = temp
                Case 3
                    chord.note.note4 = temp
            End Select
        Next
    End Sub
    Function findquality(ByRef chord As TChord, ByVal chords() As TChord)
        For i = 0 To chords.Length - 1
            For j = 0 To 3
                Select Case j
                    Case 0
                        If chord.offsets.offset1 = chords(i).offsets.offset1 Then
                            If chord.offsets.offset2 = chords(i).offsets.offset2 Then
                                If chord.offsets.offset3 = chords(i).offsets.offset3 Then
                                    If chord.offsets.offset4 = chords(i).offsets.offset4 Then
                                        Return j & chords(i).quality
                                    End If
                                End If
                            End If
                        End If
                    Case 1
                        If chord.offsets.offset1 = chords(i).offsets.offset2 Then
                            If chord.offsets.offset2 = chords(i).offsets.offset3 Then
                                If chord.offsets.offset3 = chords(i).offsets.offset4 Then
                                    If chord.offsets.offset4 = chords(i).offsets.offset1 Then
                                        Return j & chords(i).quality
                                    End If
                                End If
                            End If
                        End If
                    Case 2
                        If chord.offsets.offset1 = chords(i).offsets.offset3 Then
                            If chord.offsets.offset2 = chords(i).offsets.offset4 Then
                                If chord.offsets.offset3 = chords(i).offsets.offset1 Then
                                    If chord.offsets.offset4 = chords(i).offsets.offset2 Then
                                        Return j & chords(i).quality
                                    End If
                                End If
                            End If
                        End If
                    Case 3
                        If chord.offsets.offset1 = chords(i).offsets.offset4 Then
                            If chord.offsets.offset2 = chords(i).offsets.offset1 Then
                                If chord.offsets.offset3 = chords(i).offsets.offset2 Then
                                    If chord.offsets.offset4 = chords(i).offsets.offset3 Then
                                        Return j & chords(i).quality
                                    End If
                                End If
                            End If
                        End If
                End Select
            Next
        Next
        Return "Error"
    End Function
    Sub initializeoffsets(ByVal chords() As TChord)
        For i = 0 To 11
            chords(i).offsets.offset1 = 0
        Next
        chords(0).quality = "7"
        chords(0).offsets.offset2 = 4
        chords(0).offsets.offset3 = 7
        chords(0).offsets.offset4 = 10
        chords(1).quality = "-7"
        chords(1).offsets.offset2 = 3
        chords(1).offsets.offset3 = 7
        chords(1).offsets.offset4 = 10
        chords(2).quality = "Δ7"
        chords(2).offsets.offset2 = 4
        chords(2).offsets.offset3 = 7
        chords(2).offsets.offset4 = 11
        chords(3).quality = "-Δ7"
        chords(3).offsets.offset2 = 3
        chords(3).offsets.offset3 = 7
        chords(3).offsets.offset4 = 11
        chords(4).quality = "7#9"
        chords(4).offsets.offset2 = 3
        chords(4).offsets.offset3 = 4
        chords(4).offsets.offset4 = 10
        chords(5).quality = "ø"
        chords(5).offsets.offset2 = 3
        chords(5).offsets.offset3 = 6
        chords(5).offsets.offset4 = 10
        chords(6).quality = "add9"
        chords(6).offsets.offset2 = 2
        chords(6).offsets.offset3 = 4
        chords(6).offsets.offset4 = 7
        chords(7).quality = "-add9"
        chords(7).offsets.offset2 = 2
        chords(7).offsets.offset3 = 3
        chords(7).offsets.offset4 = 7
        chords(8).quality = "7#5"
        chords(8).offsets.offset2 = 4
        chords(8).offsets.offset3 = 8
        chords(8).offsets.offset4 = 10
        chords(9).quality = "Δ7#5"
        chords(9).offsets.offset2 = 4
        chords(9).offsets.offset3 = 8
        chords(9).offsets.offset4 = 11
        chords(9).quality = "7sus"
        chords(9).offsets.offset2 = 5
        chords(9).offsets.offset3 = 7
        chords(9).offsets.offset4 = 10
        chords(9).quality = "-Δ7#5"
        chords(9).offsets.offset2 = 3
        chords(9).offsets.offset3 = 8
        chords(9).offsets.offset4 = 11
    End Sub

End Module
