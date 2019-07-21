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
                End If
            End If
        Loop Until solved = True
    End Sub
    Sub solveextendedpassword()
        Dim ewords() As String = {"adjust", "anchor", "bowtie", "button", "cipher", "corner", "dampen", "demote", "enlist", "evolve", "forget", "finish", "geyser", "global", "hammer", "helium", "ignite", "indigo", "jigsaw", "juliet", "karate", "lambda", "listen", "matter", "memory", "nebula", "nickel", "overdo", "oxygen", "peanut", "photon", "quartz", "quebec", "resist", "riddle", "sierra", "strike", "teapot", "twenty", "untold", "ultima", "victor", "violet", "wither", "wrench", "xenons", "xylose", "yellow", "yogurt", "zenith", "zodiac"}
        beginpasswordsolve(ewords, 1)
    End Sub
    Sub solvepassword()
        Dim words() As String = {"about", "after", "again", "below", "could", "every", "first", "found", "great", "house", "large", "learn", "never", "other", "place", "plant", "point", "right", "small", "sound", "spell", "still", "study", "their", "there", "these", "thing", "think", "three", "water", "where", "which", "world", "would", "write"}
        beginpasswordsolve(words, 0)
    End Sub

    Sub beginpasswordsolve(words(), extendedpassword)
        Dim chars(4 + extendedpassword, 5) As Char
        inputchars(chars, extendedpassword)
        Dim answer As String
        answer = shortenlist(words, chars, 0, extendedpassword)
        voice.Speak(answer)
        Console.ReadLine()
    End Sub

    Sub inputchars(ByRef chars(,) As Char, ByVal extendedpassword As Integer)

        For i = 0 To 4 + extendedpassword
            For j = 0 To 5
                Console.Write("Enter character " & j + 1 & " in position " & i + 1 & ": ")
                chars(i, j) = Console.ReadLine
            Next
        Next
    End Sub
    Function shortenlist(ByRef words() As String, ByVal chars(,) As Char, ByRef letterindex As Integer, ByVal extended As Integer)
        If words.Length = 1 Then
            Return (words(0))
        End If
        If letterindex > 4 + extended Then
            Return "Error"
        End If
        Dim currentword As String
        Dim wordsisvalid As Boolean
        Dim newwords(-1) As String
        For i = 0 To words.Length - 1
            currentword = words(i)
            For j = 0 To 5
                If currentword(letterindex) = chars(letterindex, j) Then
                    wordsisvalid = True
                End If
            Next
            If wordsisvalid = True Then
                ReDim Preserve newwords(newwords.Length)
                newwords(newwords.Length - 1) = currentword
            End If
            wordsisvalid = False
        Next
        Dim answer = shortenlist(newwords, chars, letterindex + 1, extended)
        Return answer
    End Function
    Sub solvebitwise()
        Dim bits(6, 0) As Boolean
        Dim op As String
        checkbitconditions(bits)
        op = inputoperator()
        Console.Write(calcoutput(op, bits))
    End Sub
    Function calcoutput(ByVal op As String, ByVal bits(,) As Boolean)
        Dim solution As String = ""
        Select Case op
            Case "AND"
                For i = 0 To 7
                    If bits(i, 0) And bits(i, 1) Then
                        solution = solution & 1
                    Else
                        solution = solution & 0
                    End If
                Next
            Case "OR"
                For i = 0 To 7
                    If bits(i, 0) Or bits(i, 1) Then
                        solution = solution & 1
                    Else
                        solution = solution & 0
                    End If
                Next
            Case "XOR"
                For i = 0 To 7
                    If bits(i, 0) Xor bits(i, 1) Then
                        solution = solution & 1
                    Else
                        solution = solution & 0
                    End If
                Next
            Case "NOT"
                For i = 0 To 7
                    If bits(i, 0) Then
                        solution = solution & 0
                    Else
                        solution = solution & 1
                    End If
                Next
        End Select
    End Function
    Function inputoperator()
        Dim valid As Boolean
        Do
            Console.Write("Input Operator: ")
            Dim temp1 As String = UCase(Console.ReadLine)
            Select Case temp1
                Case "AND"
                    Return temp1
                Case "OR"
                    Return temp1
                Case "XOR"
                    Return temp1
                Case "NOT"
                    Return temp1
                Case Else
                    Console.WriteLine("Invalid Operator")
            End Select
        Loop
    End Function
    Sub checkbitconditions(bits(,) As Boolean)
        If aabats = 0 Then
            bits(0, 0) = True
        Else
            bits(0, 0) = False
        End If
        If Dbats > 0 Then
            bits(0, 1) = True
        Else
            bits(0, 1) = False
        End If
        If ports(1).count > 0 Then
            bits(1, 0) = True
        Else
            bits(1, 0) = False
        End If
        If countports(ports) >= 3 Then
            bits(1, 1) = True
        Else
            bits(1, 1) = False
        End If
        If inds(8).present And inds(8).lit Then
            bits(2, 0) = True
        Else
            bits(2, 0) = False
        End If
        If holders >= 2 Then
            bits(2, 1) = True
        Else
            bits(2, 1) = False
        End If
        If modulecount > time Then
            bits(3, 0) = True
        Else
            bits(3, 0) = False
        End If
        If inds(10).present Then
            bits(3, 1) = True
        Else
            bits(3, 1) = False
        End If
        If litunlit(0) > 1 Then
            bits(4, 0) = True
        Else
            bits(4, 0) = False
        End If
        If litunlit(1) > 1 Then
            bits(4, 1) = True
        Else
            bits(4, 1) = False
        End If

        If modulecount Mod 3 = 0 Then
            bits(5, 0) = True
        Else
            bits(5, 0) = False
        End If
        If Mid(serial, 4, 0) Then
            bits(5, 1) = True
        Else
            bits(5, 1) = False
        End If
        If Dbats < 2 Then
            bits(6, 0) = True
        Else
            bits(6, 0) = False
        End If
        If checkeven(modulecount) Then
            bits(6, 1) = True
        Else
            bits(6, 1) = False
        End If
        If countports() < 4 Then
            bits(7, 0) = True
        Else
            bits(7, 0) = False
        End If
        If batteries >= 2 Then
            bits(7, 1) = True
        Else
            bits(7, 1) = False
        End If
    End Sub
    Function countports()
        Dim total As Integer
        For i = 0 To ports.Length - 1
            total = total + ports(i).count
        Next
        Return total
    End Function
    Sub solvemaze()
            Dim mazeno As Integer = 0
            Dim mazes(8, 10, 10) As Char
            Dim startx, starty, endx, endy As Integer
            Dim movestring As String = ""
            initializemazes(mazes)
        displaymaze(mazes, mazeno, 0)
        inputstartandend(mazes, mazeno, startx, starty, endx, endy)
            movestring = findroute(mazes, mazeno, startx, starty, endx, endy, movestring)
            displaymovestring(movestring)
        End Sub
        Sub displaymovestring(ByVal movestring As String)
            For i = 0 To movestring.Length - 1
                Select Case movestring(i)
                    Case "U"
                        Console.WriteLine("Up")
                    Case "R"
                        Console.WriteLine("Right")
                    Case "D"
                        Console.WriteLine("Down")
                    Case "L"
                        Console.WriteLine("Left")
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
                Console.Write("Enter Starting Co-Ordinates: ")
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
                Console.Write("Enter Ending Co-Ordinates: ")
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

    End Module
