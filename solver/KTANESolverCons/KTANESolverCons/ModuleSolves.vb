Module ModuleSolves
    Sub solvecompwires()
        'solves complicated wires
    End Sub
    Sub solvewires()
        Dim wirecount As Integer
        Dim temp1 As String
        Dim valid As Boolean = False
        Do
            voice.Speak("Number of Wires: ")
            temp1 = Console.ReadLine()
            valid = checkint(temp1)
        Loop Until valid = True
        wirecount = temp1
        Dim wires(wirecount - 1) As String
        inputwires(wires, wirecount)
        Select Case wirecount
            Case 3
                If searchforwire("red", wirecount, wires) = 0 Then
                    voice.Speak("Cut wire 2")
                ElseIf wires(wirecount - 1) = "white" Then
                    voice.Speak("Cut wire " & wirecount)
                ElseIf searchforwire("blue", wirecount, wires) > 1 Then
                    voice.Speak("Cut wire " & 1 + findlastofcolour("blue", wirecount, wires))
                Else
                    voice.Speak("Cut wire 3")
                End If
            Case 4
                If (searchforwire("red", wirecount, wires) > 1) And checkeven(CInt(Mid(serial, 6))) = False Then
                    voice.Speak("Cut wire " & 1 + findlastofcolour("red", wirecount, wires))
                ElseIf wires(wirecount - 1) = "yellow" Or wires(wirecount - 1) = "y" And searchforwire("red", wirecount, wires) = 0 Then
                ElseIf searchforwire("blue", wirecount, wires) = 1 Then
                    voice.Speak("Cut wire 1")
                ElseIf searchforwire("yellow", wirecount, wires) > 1 Then
                    voice.Speak("Cut wire " & wirecount)
                Else
                    voice.Speak("Cut wire 2")
                End If
            Case 5
                If wires(wirecount - 1) = "black" Or wires(wirecount - 1) = "b" And checkeven(CInt(Mid(serial, 6))) = False Then
                    voice.Speak("Cut wire 4")
                ElseIf searchforwire("red", wirecount, wires) = 1 And searchforwire("yellow", wirecount, wires) > 1 Then
                    voice.Speak("Cut wire 1")
                ElseIf searchforwire("black", wirecount, wires) = 0 Then
                    voice.Speak("Cut wire 2")
                Else voice.Speak("Cut wire 1")
                End If
            Case 6
                If searchforwire("yellow", wirecount, wires) = 0 And checkeven(CInt(Mid(serial, 6))) = False Then
                    voice.Speak("Cut wire 3")
                ElseIf searchforwire("yellow", wirecount, wires) = 1 And searchforwire("white", wirecount, wires) > 1 Then
                    voice.Speak("Cut wire 4")
                ElseIf searchforwire("red", wirecount, wires) = 0 Then
                    voice.Speak("Cut wire " & wirecount)
                Else voice.Speak("Cut wire 4")
                End If
        End Select
    End Sub
    Sub solvewireseq()
        Dim red, blue, black As Integer
        Dim currentwire
        For i = 1 To 12
            currentwire = inputwire(i)
            Select Case currentwire
                Case "red"
                    red = red + 1
                    Select Case red
                        Case 1
                            voice.Speak("Cut if connected to C")
                        Case 2
                            voice.Speak("Cut if connected to B")
                        Case 3
                            voice.Speak("Cut if connected to A")
                        Case 4
                            voice.Speak("Cut if connected to A or C")
                        Case 5
                            voice.Speak("Cut if connected to B")
                        Case 6
                            voice.Speak("Cut if connected to A or C")
                        Case 7
                            voice.Speak("Cut")
                        Case 8
                            voice.Speak("Cut if connected to A or B")
                        Case 9
                            voice.Speak("Cut if connected to B")
                    End Select
                Case "blue"
                    blue = blue + 1
                    Select Case blue
                        Case 1
                            voice.Speak("Cut if connected to B")
                        Case 2
                            voice.Speak("Cut if connected to A or C")
                        Case 3
                            voice.Speak("Cut if connected to B")
                        Case 4
                            voice.Speak("Cut if connected to A")
                        Case 5
                            voice.Speak("Cut if connected to B")
                        Case 6
                            voice.Speak("Cut if connected to B or C")
                        Case 7
                            voice.Speak("Cut if connected to C")
                        Case 8
                            voice.Speak("Cut if connected to A or C")
                        Case 9
                            voice.Speak("Cut if connected to A")
                    End Select
                Case "black"
                    black = black + 1
                    Select Case black
                        Case 1
                            voice.Speak("Cut")
                        Case 2
                            voice.Speak("Cut if connected to A or C")
                        Case 3
                            voice.Speak("Cut if connected to B")
                        Case 4
                            voice.Speak("Cut if connected to A or C")
                        Case 5
                            voice.Speak("Cut if connected to B")
                        Case 6
                            voice.Speak("Cut if connected to B or C")
                        Case 7
                            voice.Speak("Cut if connected to A or B")
                        Case 8
                            voice.Speak("Cut if connected to C")
                        Case 9
                            voice.Speak("Cut if connected to C")
                    End Select
                Case ""
            End Select
        Next
    End Sub
    Sub SolveButton()
        Dim step2 As Boolean
        Dim solved As Boolean
        Dim btext, bcolour, scolour As String
        bcolour = ""
        btext = ""
        scolour = ""
        'input button colour
        voice.Speak("Button Colour: ")
        bcolour = Console.ReadLine
        'input button text
        voice.Speak("Button Text: ")
        btext = Console.ReadLine()
        Do
            If step2 = False Then
                If bcolour = "" Or btext = "" Then
                ElseIf LCase(bcolour) = "red" And LCase(btext) = "hold" Then
                    voice.Speak("Tap")
                    solved = True
                Else
                    step2 = True
                    'need to look at strip colour

                End If
            Else
                voice.Speak("Hold")
                voice.Speak("Input Strip Colour: ")
                scolour = Console.ReadLine()
                'check strip colour
                If scolour <> "" Then
                    Dim result As Integer
                    Select Case LCase(scolour)
                        Case "blue"
                            result = 4
                        Case "yellow"
                            result = 5
                        Case Else
                            result = 1
                    End Select
                    voice.Speak("Release on " & result)
                    solved = True
                End If
            End If
        Loop Until solved = True
    End Sub
    Sub Solvepassword()
        Dim extendedpassword As Integer = 0
        Dim answer As String
        Dim words() As String = {"about", "after", "again", "below", "could", "every", "first", "found", "great", "house", "large", "learn", "never", "other", "place", "plant", "point", "right", "small", "sound", "spell", "still", "study", "their", "there", "these", "thing", "think", "three", "water", "where", "which", "world", "would", "write"}
        Dim chars(5, 6) As Char
        inputletters(chars, extendedpassword)
        answer = ShortenList(words, chars, 0, extendedpassword)
        voice.Speak("The Answer is: " & answer)
    End Sub
    Function ShortenList(ByRef currentwords() As String, ByRef chars(,) As Char, ByRef letterpos As Integer, ByVal extendedpassword As Integer)
        Dim solution As String
        If currentwords.Length = 1 Then
            Return currentwords(0)
        End If
        Dim newwords() As String
        Dim loopend = currentwords.Length - 1
        For i = 0 To loopend
            Dim lettercheck = currentwords(i)(letterpos)
            If lettercheck = chars(letterpos, 0) Or lettercheck = chars(letterpos, 1) Or lettercheck = chars(letterpos, 2) Or lettercheck = chars(letterpos, 3) Or lettercheck = chars(letterpos, 4) Or lettercheck = chars(letterpos, 5) Then
                newwords.Append(currentwords(i))
            End If
        Next
        If newwords.Length = 0 Then
            Return "error"
        End If
        letterpos = letterpos + 1
        If letterpos <> 5 + extendedpassword Then
            solution = ShortenList(newwords, chars, letterpos, extendedpassword)
        End If
        Return solution
    End Function
    Sub solvecombinationlock()
        Dim num1, num2, num3 As Integer
        num1 = ((CInt(Mid(serial, 6)) + solvedmodules) + batteries) Mod 20
        num2 = modulecount + solvedmodules Mod 20
        num3 = (num1 + num2) Mod 20
        voice.Speak("Right to " & num1)
        voice.Speak("Then Left to " & num2)
        voice.Speak("Then Right to " & num3)
    End Sub
    Sub solvefastmath()
        Dim extratoadd As Integer
        Dim pair As String
        If inds(7).present = True And inds(7).lit = True Then
            extratoadd = extratoadd + 20
        End If
        If ports(0).count <> 0 Then
            extratoadd = extratoadd + 14
        End If
        If Checkforletterinserial("f") Or Checkforletterinserial("a") Or Checkforletterinserial("s") Or Checkforletterinserial("t") Then
            extratoadd = extratoadd - 5
        End If
        If ports(3).count <> 0 Then
            extratoadd = extratoadd + 27
        End If
        If batteries > 3 Then
            extratoadd = extratoadd - 15
        End If
        If extratoadd >= 100 Then
            extratoadd = extratoadd Mod 100
        ElseIf extratoadd < 0 Then
            extratoadd = extratoadd + 50
        End If
        voice.Speak("Enter first pair: ")
        pair = Console.ReadLine()
        'A=65
    End Sub
    Sub solvemaze()
            Dim mazeno As Integer = 0
            Dim mazes(8, 10, 10) As Char
            Dim startx, starty, endx, endy As Integer
            Dim movestring As String = ""
            initializemazes(mazes)
            displaymaze(mazes, mazeno, 1)
            inputstartandend(mazes, mazeno, startx, starty, endx, endy)
            movestring = findroute(mazes, mazeno, startx, starty, endx, endy, movestring)
            displaymovestring(movestring)
        End Sub
        Sub displaymovestring(ByVal movestring As String)
            For i = 0 To movestring.Length - 1
                Select Case movestring(i)
                    Case "U"
                    voice.Speak("Up")
                Case "R"
                    voice.Speak("Right")
                Case "D"
                    voice.Speak("Down")
                Case "L"
                    voice.Speak("Left")
            End Select
        Next
    End Sub
    Sub initializemazes(ByRef mazes(,,) As Char)
        'maze1
        'column1
        For i = 0 To 10
            mazes(0, 0, i) = "S"
        Next
        mazes(0, 0, 8) = "O"
        'column2
        For i = 0 To 10
            mazes(0, 1, i) = "*"
        Next
        mazes(0, 1, 0) = "S"
        mazes(0, 1, 2) = "S"
        mazes(0, 1, 10) = "S"
        'column3
        For i = 0 To 10
            mazes(0, 2, i) = "S"
        Next
        mazes(0, 2, 1) = "*"
        mazes(0, 2, 3) = "*"
        mazes(0, 2, 5) = "*"
        mazes(0, 2, 9) = "*"
        'column4
        mazes(0, 3, 0) = "*"
        mazes(0, 3, 1) = "*"
        mazes(0, 3, 2) = "S"
        mazes(0, 3, 3) = "*"
        mazes(0, 3, 4) = "S"
        mazes(0, 3, 5) = "*"
        mazes(0, 3, 6) = "S"
        mazes(0, 3, 7) = "*"
        mazes(0, 3, 8) = "S"
        mazes(0, 3, 9) = "*"
        mazes(0, 3, 10) = "S"
        'column5
        For i = 0 To 10
            mazes(0, 4, i) = "S"
        Next
        mazes(0, 4, 3) = "*"
        mazes(0, 4, 7) = "*"
        'column6
        For i = 0 To 10
            mazes(0, 5, i) = "*"
        Next
        mazes(0, 5, 0) = "S"
        mazes(0, 5, 4) = "S"
        'column7
        For i = 0 To 10
            mazes(0, 6, i) = "S"
        Next
        mazes(0, 6, 3) = "*"
        mazes(0, 6, 7) = "*"
        'column8
        For i = 0 To 10
            mazes(0, 7, i) = "*"
        Next
        mazes(0, 7, 2) = "S"
        mazes(0, 7, 6) = "S"
        mazes(0, 7, 8) = "S"
        mazes(0, 7, 10) = "S"
        'column9
        For i = 0 To 8 Step 2
            mazes(0, 8, i) = "S"
            mazes(0, 8, i + 1) = "*"
        Next
        mazes(0, 8, 10) = "S"
        'column10
        mazes(0, 9, 0) = "S"
        mazes(0, 9, 1) = "*"
        mazes(0, 9, 2) = "*"
        mazes(0, 9, 3) = "*"
        mazes(0, 9, 4) = "S"
        mazes(0, 9, 5) = "*"
        mazes(0, 9, 6) = "S"
        mazes(0, 9, 7) = "*"
        mazes(0, 9, 8) = "S"
        mazes(0, 9, 9) = "*"
        mazes(0, 9, 10) = "S"
        'column11
        For i = 0 To 10
            mazes(0, 10, i) = "S"
        Next
        mazes(0, 10, 6) = "O"
        mazes(0, 10, 9) = "*"
        'maze2
        'column1
        For i = 0 To 10
            mazes(1, 0, i) = "S"
        Next
        mazes(1, 0, 9) = "*"
        'column2
        For i = 0 To 8 Step 2
            mazes(1, 1, i) = "S"
            mazes(1, 1, i + 1) = "*"
        Next
        mazes(1, 1, 10) = "S"
        mazes(1, 1, 0) = "*"
        mazes(1, 1, 2) = "*"
        mazes(1, 1, 6) = "*"
        'column3
        For i = 0 To 10
            mazes(1, 2, i) = "S"
        Next
        mazes(1, 2, 3) = "*"
        mazes(1, 2, 4) = "O"
        mazes(1, 2, 7) = "*"
        'column4
        For i = 0 To 8 Step 2
            mazes(1, 3, i) = "S"
            mazes(1, 3, i + 1) = "*"
        Next
        mazes(1, 3, 10) = "S"
        mazes(1, 3, 2) = "*"
        mazes(1, 3, 4) = "*"
        mazes(1, 3, 8) = "*"
        'column5
        For i = 0 To 10
            mazes(1, 4, i) = "S"
        Next
        mazes(1, 4, 5) = "*"
        mazes(1, 4, 9) = "*"
        'column6
        For i = 0 To 10
            mazes(1, 5, i) = "*"
        Next
        mazes(1, 5, 4) = "S"
        mazes(1, 5, 8) = "S"
        'column7
        For i = 0 To 10
            mazes(1, 6, i) = "S"
        Next
        mazes(1, 6, 3) = "*"
        mazes(1, 6, 7) = "*"
        'column8
        For i = 0 To 8 Step 2
            mazes(1, 7, i) = "S"
            mazes(1, 7, i + 1) = "*"
        Next
        mazes(1, 7, 10) = "S"
        mazes(1, 7, 4) = "*"
        mazes(1, 7, 8) = "*"
        'column9
        For i = 0 To 10
            mazes(1, 8, i) = "S"
        Next
        mazes(1, 8, 1) = "*"
        mazes(1, 8, 5) = "*"
        mazes(1, 8, 7) = "*"
        mazes(1, 8, 8) = "O"
        'column10
        For i = 0 To 10
            mazes(1, 9, i) = "*"
        Next
        mazes(1, 9, 0) = "S"
        mazes(1, 9, 6) = "S"
        mazes(1, 9, 8) = "S"
        mazes(1, 9, 10) = "S"
        'column11
        For i = 0 To 10
            mazes(1, 10, i) = "S"
        Next
        mazes(1, 10, 9) = "*"
        'maze3
        'column1
        For i = 0 To 10
            mazes(2, 0, i) = "S"
        Next
        mazes(2, 0, 7) = "*"
        'column2
        For i = 0 To 10
            mazes(2, 1, i) = "*"
        Next
        mazes(2, 1, 0) = "S"
        mazes(2, 1, 6) = "S"
        mazes(2, 1, 10) = "S"
        'column3
        For i = 0 To 10
            mazes(2, 2, i) = "S"
        Next
        mazes(2, 2, 1) = "*"
        mazes(2, 2, 9) = "*"
        'column4
        For i = 0 To 10
            mazes(2, 3, i) = "*"
        Next
        mazes(2, 3, 0) = "S"
        mazes(2, 3, 6) = "S"
        mazes(2, 3, 10) = "S"
        'carry on from here
    End Sub
    Sub displaymaze(ByVal mazes(,,) As Char, ByVal mazeno As Integer, ByVal displayhidden As Integer)
        For i = 0 To 10 Step displayhidden + 1
            For j = 0 To 10 Step displayhidden + 1
                Console.Write(mazes(mazeno, j, 10 - i) & "  ")
            Next
            Console.WriteLine()
        Next
    End Sub
    Function checkcoord(ByVal coords As String)
        Dim test As Integer
        If Mid(coords, 2, 1) <> "," Then
            Return False
        End If
        For i = 1 To 3 Step 2
            Try
                test = CInt(Mid(coords, i, 1))
            Catch ex As Exception
                Return False
            End Try
            If test > 6 Or test < 1 Then
                Return False
            End If
        Next
        Return True
    End Function
    Sub inputstartandend(ByRef mazes(,,) As Char, ByVal mazeno As Integer, ByRef startx As Integer, ByRef starty As Integer, ByRef endx As Integer, ByRef endy As Integer)
        Dim startcoords As String
        Dim valid = False
        Do
            voice.Speak("Enter Starting Co-Ordinates: ")
            startcoords = Console.ReadLine()
            If checkcoord(startcoords) Then
                valid = True
            End If
        Loop Until valid = True
        starty = 2 * (CInt(Mid(startcoords, 3, 1)) - 1)
        startx = 2 * (CInt(Mid(startcoords, 1, 1)) - 1)
        Dim endcoords As String
        valid = False
        Do
            voice.Speak("Enter Ending Co-Ordinates: ")
            endcoords = Console.ReadLine()
            If checkcoord(endcoords) Then
                valid = True
            End If
        Loop Until valid = True
        endy = 2 * (CInt(Mid(endcoords, 3, 1)) - 1)
        endx = 2 * (CInt(Mid(endcoords, 1, 1)) - 1)
    End Sub
    Function findroute(ByVal mazes(,,) As Char, ByVal mazeno As Integer, ByRef currentposx As Integer, ByRef currentposy As Integer, ByVal endx As Integer, ByVal endy As Integer, ByRef movestring As String)
        Dim endfound = checkend(currentposx, currentposy, endx, endy)
        If endfound = True Then
            Return movestring
        End If
        Dim validmove As Boolean = False
        'check movement options
        If currentposy <> 10 Then
            Dim moveup1, moveup2 As Boolean
            moveup1 = ((mazes(mazeno, currentposx, currentposy + 1) = "S") Or (mazes(mazeno, currentposx, currentposy + 1) = "O")) Or (mazes(mazeno, currentposx, currentposy + 1) = "E")
            moveup2 = mazes(mazeno, currentposx, currentposy + 2) = "S" Or mazes(mazeno, currentposx, currentposy + 2) = "O" Or mazes(mazeno, currentposx, currentposy + 2) = "E"
            If moveup1 And moveup2 Then
                mazes(mazeno, currentposx, currentposy + 1) = "R"
                mazes(mazeno, currentposx, currentposy + 2) = "R"
                movestring = movestring & "U"
                currentposy = currentposy + 2
                validmove = True
                movestring = findroute(mazes, mazeno, currentposx, currentposy, endx, endy, movestring)
            End If
        End If
        endfound = checkend(currentposx, currentposy, endx, endy)
        If endfound = True Then
            Return movestring
        End If
        If currentposx <> 10 Then
            Dim moveright1, moveright2 As Boolean
            moveright1 = ((mazes(mazeno, currentposx + 1, currentposy) = "S") Or (mazes(mazeno, currentposx + 1, currentposy) = "O")) Or (mazes(mazeno, currentposx + 1, currentposy) = "E")
            moveright2 = ((mazes(mazeno, currentposx + 2, currentposy) = "S") Or (mazes(mazeno, currentposx + 2, currentposy) = "O")) Or (mazes(mazeno, currentposx + 2, currentposy) = "E")
            If moveright1 And moveright2 Then
                mazes(mazeno, currentposx + 1, currentposy) = "R"
                mazes(mazeno, currentposx + 2, currentposy) = "R"
                movestring = movestring & "R"
                currentposx = currentposx + 2
                validmove = True
                movestring = findroute(mazes, mazeno, currentposx, currentposy, endx, endy, movestring)
            End If
        End If
        endfound = checkend(currentposx, currentposy, endx, endy)
        If endfound = True Then
            Return movestring
        End If
        If currentposy <> 0 Then
            Dim movedown1, movedown2 As Boolean
            movedown1 = ((mazes(mazeno, currentposx, currentposy - 1) = "S") Or (mazes(mazeno, currentposx, currentposy - 1) = "O")) Or (mazes(mazeno, currentposx, currentposy - 1) = "E")
            movedown2 = mazes(mazeno, currentposx, currentposy - 2) = "S" Or mazes(mazeno, currentposx, currentposy - 2) = "O" Or mazes(mazeno, currentposx, currentposy - 2) = "E"
            If movedown1 And movedown2 Then
                mazes(mazeno, currentposx, currentposy - 1) = "R"
                mazes(mazeno, currentposx, currentposy - 2) = "R"
                movestring = movestring & "D"
                currentposy = currentposy - 2
                validmove = True
                movestring = findroute(mazes, mazeno, currentposx, currentposy, endx, endy, movestring)
            End If
        End If
        endfound = checkend(currentposx, currentposy, endx, endy)
        If endfound = True Then
            Return movestring
        End If
        If currentposx <> 0 Then
            Dim moveleft1, moveleft2 As Boolean
            moveleft1 = ((mazes(mazeno, currentposx - 1, currentposy) = "S") Or (mazes(mazeno, currentposx - 1, currentposy) = "O")) Or (mazes(mazeno, currentposx - 1, currentposy) = "E")
            moveleft2 = ((mazes(mazeno, currentposx - 2, currentposy) = "S") Or (mazes(mazeno, currentposx - 2, currentposy) = "O")) Or (mazes(mazeno, currentposx - 2, currentposy) = "E")
            If (mazes(mazeno, currentposx - 1, currentposy) = "S") And (mazes(mazeno, currentposx - 2, currentposy) = "S") Then
                mazes(mazeno, currentposx - 1, currentposy) = "R"
                mazes(mazeno, currentposx - 2, currentposy) = "R"
                movestring = movestring & "L"
                currentposx = currentposx - 2
                validmove = True
                movestring = findroute(mazes, mazeno, currentposx, currentposy, endx, endy, movestring)
            End If
        End If
        endfound = checkend(currentposx, currentposy, endx, endy)
        If endfound = True Then
            Return movestring
        End If
        If validmove = False Then
            mazes(mazeno, currentposx, currentposy) = "X"
            Select Case movestring(movestring.Length - 1)
                Case "U"
                    mazes(mazeno, currentposx, currentposy - 1) = "X"
                    currentposy = currentposy - 2
                Case "R"
                    mazes(mazeno, currentposx - 1, currentposy) = "X"
                    currentposx = currentposx - 2
                Case "D"
                    mazes(mazeno, currentposx, currentposy + 1) = "X"
                    currentposy = currentposy + 2
                Case "L"
                    mazes(mazeno, currentposx + 1, currentposy) = "X"
                    currentposx = currentposx + 2
            End Select
            movestring = Mid(movestring, 1, movestring.Length - 1)
        End If
        movestring = findroute(mazes, mazeno, currentposx, currentposy, endx, endy, movestring)
        Return movestring
    End Function
    Function checkend(ByVal currentposx As Integer, ByVal currentposy As Integer, ByVal endx As Integer, ByVal endy As Integer)
        If currentposx = endx And currentposy = endy Then
            Return True
        End If
        Return False
    End Function
    Sub solvekeypad()

        Dim symbols(3) As String
        Dim convertedsymbols(3) As Integer
        Dim reference(5, 6) As Integer
        initializereference(reference)
        'showtable(reference)
        inputsymbols(symbols)
        For i = 0 To 3
            convertedsymbols(i) = converttonums(symbols, i)
        Next
        Dim column As Integer = findcolumn(convertedsymbols, reference)
        findsymbolorder(convertedsymbols, reference, column)
        Console.ReadLine()
    End Sub
    Sub findsymbolorder(symbols() As Integer, reference(,) As Integer, ByVal column As Integer)
        Dim foundcount As Integer = 0
        For i = 0 To 6
            For k = 0 To 3
                If reference(column, i) = symbols(k) Then
                    voice.Speak(referencetotext(symbols(k)))
                    If foundcount < 3 Then
                        voice.Speak(", ")
                        foundcount = foundcount + 1
                    End If
                End If
            Next
        Next
    End Sub
    Function referencetotext(number As Integer)
        Select Case number
            Case 1
                Return "lollipop"
            Case 2
                Return "equation triangle"
            Case 3
                Return "lambda"
            Case 4
                Return "lightning"
            Case 5
                Return "rocket"
            Case 6
                Return "fancy H"
            Case 7
                Return "backwards C"
            Case 8
                Return "backwards E with umlaut"
            Case 9
                Return "loop"
            Case 10
                Return "unfilled star"
            Case 11
                Return "upside down question mark"
            Case 12
                Return "copyright"
            Case 13
                Return "butt"
            Case 14
                Return "XI"
            Case 15
                Return "broken 3"
            Case 16
                Return "flat 6"
            Case 17
                Return "paragraph"
            Case 18
                Return "b bar"
            Case 19
                Return "smiley"
            Case 20
                Return "trident"
            Case 21
                Return "forward C"
            Case 22
                Return "snake"
            Case 23
                Return "filled star"
            Case 24
                Return "jigsaw"
            Case 25
                Return "ae"
            Case 26
                Return "backwards N"
            Case 27
                Return "omega"
        End Select
        Return "error"
    End Function

    Sub initializereference(ByVal reference(,) As Integer)
        reference(0, 0) = 1
        reference(0, 1) = 2
        reference(0, 2) = 3
        reference(0, 3) = 4
        reference(0, 4) = 5
        reference(0, 5) = 6
        reference(0, 6) = 7
        reference(1, 0) = 8
        reference(1, 1) = 1
        reference(1, 2) = 7
        reference(1, 3) = 9
        reference(1, 4) = 10
        reference(1, 5) = 6
        reference(1, 6) = 11
        reference(2, 0) = 12
        reference(2, 1) = 13
        reference(2, 2) = 9
        reference(2, 3) = 14
        reference(2, 4) = 15
        reference(2, 5) = 3
        reference(2, 6) = 10
        reference(3, 0) = 16
        reference(3, 1) = 17
        reference(3, 2) = 18
        reference(3, 3) = 5
        reference(3, 4) = 14
        reference(3, 5) = 11
        reference(3, 6) = 19
        reference(3, 5) = 11
        reference(4, 0) = 20
        reference(4, 1) = 19
        reference(4, 2) = 18
        reference(4, 3) = 21
        reference(4, 4) = 17
        reference(4, 5) = 22
        reference(4, 6) = 23
        reference(5, 0) = 16
        reference(5, 1) = 8
        reference(5, 2) = 24
        reference(5, 3) = 25
        reference(5, 4) = 20
        reference(5, 5) = 26
        reference(5, 6) = 27
    End Sub
    Sub showtable(ByVal reference(,) As Integer)
        For i = 0 To 6
            For j = 0 To 5
                Console.Write(reference(j, i) & " ")
                If reference(j, i) < 10 Then
                    Console.Write(" ")
                End If
            Next
            Console.WriteLine()
        Next
    End Sub
    Function findcolumn(ByVal symbols() As Integer, ByVal table(,) As Integer)
        Dim foundcount As Integer = 0
        For i = 0 To 5
            For j = 0 To 6
                For k = 0 To 3
                    If symbols(k) = table(i, j) Then
                        foundcount = foundcount + 1
                    End If
                Next
            Next
            If foundcount = 4 Then
                Return i
            Else
                foundcount = 0
            End If
        Next
        Return "Error"
    End Function
    Sub inputsymbols(ByVal symbols() As String)
        For i = 0 To 3
            voice.Speak("Input Symbol " & i + 1 & ": ")
            symbols(i) = LCase(Console.ReadLine())
        Next
    End Sub
    Function converttonums(ByVal symbols() As String, ByVal symbolnumber As Integer)
        Select Case LCase(symbols(symbolnumber))
            Case "lollipop"
                Return 1
            Case "equation triangle"
                Return 2
            Case "lambda"
                Return 3
            Case "lightning"
                Return 4
            Case "rocket"
                Return 5
            Case "fancy h"
                Return 6
            Case "backwards c"
                Return 7
            Case "backwards e with umlaut"
                Return 8
            Case "Loop"
                Return 9
            Case "unfilled star"
                Return 10
            Case "upside down question mark "
                Return 11
            Case "copyright"
                Return 12
            Case "butt"
                Return 13
            Case "xi"
                Return 14
            Case "broken 3"
                Return 15
            Case "flat 6"
                Return 16
            Case "paragraph"
                Return 17
            Case "b bar"
                Return 18
            Case "smiley"
                Return 19
            Case "trident"
                Return 20
            Case "forward c"
                Return 21
            Case "snake"
                Return 22
            Case "filled star"
                Return 23
            Case "jigsaw"
                Return 24
            Case "ae"
                Return 25
            Case "backwards n"
                Return 26
            Case "omega"
                Return 27
        End Select
        Return "error"
    End Function
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
    Sub solvechordqs()
        Dim notes() As String = {"A", "A#", "B", "C", "C#", "D", "D#", "E", "F", "F#", "G", "G#"}
        Dim chords(11) As TChord
        initializeoffsets(chords)
        Dim chord1 As TChord
        inputstartchord(chord1)
        chordoffsets(chord1, notes)
        Dim startquality = findbaseind(chord1.note.note1)
        Dim startandquality = findquality(chord1, chords)
        If startandquality <> "CNF" Then
            Dim startnote As String = Mid(startandquality, 1, 1)
            Dim quality As String = Mid(startandquality, 2)
            findanswerchord(startnote, quality, chord1, chords)
        End If
        Console.ReadLine()
    End Sub
    Sub findanswerchord(start As Integer, quality As String, chord As TChord, chords() As TChord)
        Dim startnote As String = setstartnote(startnote, start, chord)
        Dim ansquality As String = roottoquality(startnote)
        Dim ansroot As String = qualitytoroot(quality)
        Dim qind As Integer = findqualityindex(ansquality, chords)
        Dim endstartind As Integer = findbaseind(ansroot)
        findnotesolution(chords, endstartind, qind)
    End Sub
    Sub findnotesolution(ByVal chords() As TChord, ByVal endstartind As Integer, ByVal qind As Integer)
        Dim ansnotes As tnotes
        Dim notes() As String = {"A", "A#", "B", "C", "C#", "D", "D#", "E", "F", "F#", "G", "G#"}
        ansnotes.note1 = notes(endstartind)
        ansnotes.note2 = notes((endstartind + chords(qind).offsets.offset2) Mod 12)
        ansnotes.note3 = notes((endstartind + chords(qind).offsets.offset3) Mod 12)
        ansnotes.note4 = notes((endstartind + chords(qind).offsets.offset4) Mod 12)
        voice.Speak("Enter: " & ansnotes.note1 & ", " & ansnotes.note2 & ", " & ansnotes.note3 & " and " & ansnotes.note4)
    End Sub
    Function setstartnote(ByRef startnote As String, ByVal start As Integer, chord As TChord)
        Select Case start
            Case 0
                Return chord.note.note1
            Case 1
                Return chord.note.note2
            Case 2
                Return chord.note.note3
            Case 3
                Return chord.note.note4
        End Select
    End Function
    Function findqualityindex(ByVal quality As String, chords() As TChord)
        For i = 0 To chords.Length - 1
            If chords(i).quality = quality Then
                Return i
            End If
        Next
        Return "Error"
    End Function
    Sub displayoffsets(ByVal chord As TChord)
        voice.Speak(chord.offsets.offset1 & ", " & chord.offsets.offset2 & ", " & chord.offsets.offset3 & ", " & chord.offsets.offset4)
    End Sub
    Sub chordoffsets(ByRef chord As TChord, ByVal notes() As String)
        Dim foundnotes As Integer = 0
        Dim offset As Integer = 0
        Dim startpoint = findstartchord(chord.note.note1, notes)
        chord.offsets.offset1 = 0
        If CStr(startpoint) = "error" Then

            Do Until CStr(startpoint) <> "Error"
                voice.Speak("Input Error")
                inputstartchord(chord)
                findstartchord(chord.note.note1, notes)
            Loop
        End If
        Do
            For j = startpoint To startpoint + 11
                Dim checker = j Mod 12
                Select Case foundnotes
                    Case 0
                        If chord.note.note2 = notes(checker) Then
                            chord.offsets.offset2 = offset
                            offset = 0
                            foundnotes = foundnotes + 1
                        End If
                    Case 1
                        If chord.note.note3 = notes(checker) Then
                            chord.offsets.offset3 = offset
                            offset = 0
                            foundnotes = foundnotes + 1
                        End If
                    Case 2
                        If chord.note.note4 = notes(checker) Then
                            chord.offsets.offset4 = offset
                            chord.offsets.offset5 = 12 - (chord.offsets.offset2 + chord.offsets.offset3 + chord.offsets.offset4)
                            offset = 0
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
        Return "Error"
    End Function
    Function findbaseind(ByVal startnote As String)
        Select Case startnote
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
            voice.Speak("Enter note " & i + 1 & ": ")
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
                If chord.offsets.offset1 = chords(i).offsets.offset1 Then
                    If chord.offsets.offset2 + chord.offsets.offset1 = chords(i).offsets.offset2 Then
                        If chord.offsets.offset3 + chord.offsets.offset2 + chord.offsets.offset1 = chords(i).offsets.offset3 Then
                            If chord.offsets.offset4 + chord.offsets.offset3 + chord.offsets.offset2 + chord.offsets.offset1 = chords(i).offsets.offset4 Then
                                Return j & chords(i).quality
                            End If
                        End If
                    End If
                End If
                rotateoffsets(chord)
            Next
        Next
        voice.Speak("chord not found")
        Return "CNF"
    End Function
    Sub rotateoffsets(ByRef chord As TChord)
        Dim temp As Integer
        temp = chord.offsets.offset2
        chord.offsets.offset2 = chord.offsets.offset3
        chord.offsets.offset3 = chord.offsets.offset4
        chord.offsets.offset4 = chord.offsets.offset5
        chord.offsets.offset5 = temp
    End Sub
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
        chords(2).quality = "delta7"
        chords(2).offsets.offset2 = 4
        chords(2).offsets.offset3 = 7
        chords(2).offsets.offset4 = 11
        chords(3).quality = "-delta7"
        chords(3).offsets.offset2 = 3
        chords(3).offsets.offset3 = 7
        chords(3).offsets.offset4 = 11
        chords(4).quality = "7#9"
        chords(4).offsets.offset2 = 3
        chords(4).offsets.offset3 = 4
        chords(4).offsets.offset4 = 10
        chords(5).quality = "ostroke"
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
        chords(9).quality = "delta7#5"
        chords(9).offsets.offset2 = 4
        chords(9).offsets.offset3 = 8
        chords(9).offsets.offset4 = 11
        chords(9).quality = "7sus"
        chords(9).offsets.offset2 = 5
        chords(9).offsets.offset3 = 7
        chords(9).offsets.offset4 = 10
        chords(9).quality = "-delta7#5"
        chords(9).offsets.offset2 = 3
        chords(9).offsets.offset3 = 8
        chords(9).offsets.offset4 = 11
    End Sub
    Function roottoquality(ByVal startnote As String)
        Select Case startnote
            Case "A"
                Return "-delta7#5"
            Case "A#"
                Return "delta7#5"
            Case "B"
                Return "-7"
            Case "C"
                Return "ostroke"
            Case "C#"
                Return "-add9"
            Case "D"
                Return "delta7"
            Case "D#"
                Return "7#9"
            Case "E"
                Return "7sus"
            Case "F"
                Return "add9"
            Case "F#"
                Return "7"
            Case "G"
                Return "-delta7"
            Case "G#"
                Return "7#5"
        End Select
        Return "Error"
    End Function
    Function qualitytoroot(ByRef quality As String)
        Select Case quality
            Case "7"
                Return "G"
            Case "-7"
                Return "G#"
            Case "delta7"
                Return "A#"
            Case "-delta7"
                Return "F"
            Case "7#9"
                Return "A"
            Case "ostroke"
                Return "C#"
            Case "add9"
                Return "D#"
            Case "-add9"
                Return "E"
            Case "7#5"
                Return "F#"
            Case "delta7#5"
                Return "C"
            Case "7sus"
                Return "D"
            Case "-delta7#5"
                Return "B"
        End Select
        Return "Error"
    End Function
End Module
