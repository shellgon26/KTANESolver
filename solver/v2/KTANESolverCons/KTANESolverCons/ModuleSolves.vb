Module ModuleSolves
    Class TreeNode
        Public Value As Object
        Public leftNode As TreeNode
        Public RightNode As TreeNode
    End Class
    Class Tree
        Public _root As TreeNode
        Sub New(ByVal rootValue As Object)
            _root = GetNode(rootValue)
        End Sub

        Private Function GetNode(ByVal value As Object) As TreeNode
            Dim node As New TreeNode()
            node.Value = value
            Return node
        End Function
        Public Sub AddtoTree(ByVal value As Double)
            Dim currentNode As TreeNode = _root
            Dim nextNode As TreeNode = _root

            While currentNode.Value <> value And Not nextNode Is Nothing
                currentNode = nextNode
                If nextNode.Value < value Then
                    nextNode = nextNode.RightNode
                Else
                    nextNode = nextNode.leftNode
                End If
            End While
            If currentNode.Value = value Then
                'red colour MSD - 0000/1000
                'blue colour 2nd - 0000/0100
                'led 3rd - 0000/0010
                'star LSD - 0000/0001
                Select Case value
                    Case 0, 1, 9
                        cuttypec()
                    Case 2, 5, 15
                        cuttyped()
                    Case 3, 10, 11
                        cuttypeb()
                    Case 4, 8, 12, 14
                        cuttypes()
                    Case 6, 7, 13
                        cuttypep()
                End Select
            ElseIf currentNode.Value < value Then
                currentNode.RightNode = GetNode(value)
                voice.Speak("Added: " & value)
            Else
                currentNode.leftNode = GetNode(value)
                voice.Speak("Added: " & value)
            End If
        End Sub
        Sub inOrderTraverse(ByVal node As TreeNode)
            If Not node Is Nothing Then
                inOrderTraverse(node.leftNode)
                voice.Speak(node.Value)
                inOrderTraverse(node.RightNode)
            End If
        End Sub
    End Class
    Sub cuttypec()
        voice.Speak("Cut")
    End Sub
    Sub cuttyped()
        voice.Speak("Don't Cut")
    End Sub
    Sub cuttypep()
        Dim parallel As Boolean
        If parallel Then
            cuttypec()
        Else
            cuttyped()
        End If
    End Sub
    Sub cuttypes()
        Dim lastserialnum As Integer
        If lastserialnum Mod 2 = 0 Then
            cuttypec()
        Else
            cuttyped()
        End If
    End Sub
    Sub cuttypeb()
        Dim batteries As Integer
        If batteries >= 2 Then
            cuttypec()
        Else
            cuttyped()
        End If
    End Sub
    Sub solvecompwires()
        Dim btree As New Tree(7.5)
        initialisebinarytree(btree)
        Dim desciptionstring As String = "red, led, star"
        Dim keywords = {"red", "blue", "led", "star"}
        Dim wirescore As Integer = 0
        For j = 0 To 3
            For i = 0 To desciptionstring.Length - keywords(j).Length
                Dim checkword = Mid(desciptionstring, i + 1, keywords(j).Length)
                If checkword = keywords(j) Then
                    wirescore += 2 ^ (3 - j)
                End If
            Next
        Next
        btree.AddtoTree(wirescore)
    End Sub
    Sub initialisebinarytree(ByRef btree As Tree)
        Dim values() As Integer = {1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15}
        For i = 0 To values.Length - 1
            btree.AddtoTree(values(i))
        Next
    End Sub
    Sub Solvewires()
        Dim wirecount As Integer
        Dim temp1 As String
        takespeech("Number of Wires: ", temp1, True)
            wirecount = temp1
        Dim wires(wirecount - 1) As String
        Inputwires(wires, wirecount)
        Select Case wirecount
            Case 3
                If Searchforwire("red", wirecount, wires) = 0 Then
                    voice.Speak("Cut wire 2")
                ElseIf wires(wirecount - 1) = "white" Then
                    voice.Speak("Cut wire " & wirecount)
                ElseIf Searchforwire("blue", wirecount, wires) > 1 Then
                    voice.Speak("Cut wire " & 1 + Findlastofcolour("blue", wirecount, wires))
                Else
                    voice.Speak("Cut wire 3")
                End If
            Case 4
                If (Searchforwire("red", wirecount, wires) > 1) And Checkeven(CInt(Mid(serial, 6))) = False Then
                    voice.Speak("Cut wire " & 1 + Findlastofcolour("red", wirecount, wires))
                ElseIf wires(wirecount - 1) = "yellow" Or wires(wirecount - 1) = "y" And Searchforwire("red", wirecount, wires) = 0 Then
                ElseIf Searchforwire("blue", wirecount, wires) = 1 Then
                    voice.Speak("Cut wire 1")
                ElseIf Searchforwire("yellow", wirecount, wires) > 1 Then
                    voice.Speak("Cut wire " & wirecount)
                Else
                    voice.Speak("Cut wire 2")
                End If
            Case 5
                If wires(wirecount - 1) = "black" Or wires(wirecount - 1) = "b" And Checkeven(CInt(Mid(serial, 6))) = False Then
                    voice.Speak("Cut wire 4")
                ElseIf Searchforwire("red", wirecount, wires) = 1 And Searchforwire("yellow", wirecount, wires) > 1 Then
                    voice.Speak("Cut wire 1")
                ElseIf Searchforwire("black", wirecount, wires) = 0 Then
                    voice.Speak("Cut wire 2")
                Else voice.Speak("Cut wire 1")
                End If
            Case 6
                If Searchforwire("yellow", wirecount, wires) = 0 And Checkeven(CInt(Mid(serial, 6))) = False Then
                    voice.Speak("Cut wire 3")
                ElseIf Searchforwire("yellow", wirecount, wires) = 1 And Searchforwire("white", wirecount, wires) > 1 Then
                    voice.Speak("Cut wire 4")
                ElseIf Searchforwire("red", wirecount, wires) = 0 Then
                    voice.Speak("Cut wire " & wirecount)
                Else voice.Speak("Cut wire 4")
                End If

            Case Else
                voice.Speak("ERROR")
        End Select
    End Sub
    Sub Solvewireseq()
        Dim red, blue, black As Integer
        Dim currentwire
        For i = 1 To 12
            currentwire = Inputwire(i)
            Select Case currentwire
                Case "red"
                    red += 1
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
                        Case Else
                            voice.Speak("ERROR")
                    End Select
                Case "blue"
                    blue += 1
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
                        Case Else
                            voice.Speak("ERROR")
                    End Select
                Case "black"
                    black += 1
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
                        Case Else
                            voice.Speak("ERROR")
                    End Select
                Case ""
                Case Else
                    voice.Speak("ERROR")
            End Select
        Next
    End Sub
    Sub SolveButton()
        Dim step2 As Boolean
        Dim solved As Boolean
        Dim btext, bcolour, scolour As String
        'input button colour
        takespeech("Button Colour: ", bcolour)
        'input button text
        takespeech("Button Text: ", btext)
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
                takespeech("Input Strip Colour: ", scolour)
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
        Inputletters(chars, extendedpassword)
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
#Disable Warning BC42104 ' Variable is used before it has been assigned a value
                newwords.Append(currentwords(i))
#Enable Warning BC42104 ' Variable is used before it has been assigned a value
            End If
        Next
        If newwords.Length = 0 Then
            Return "error"
        End If
        letterpos += 1
        If letterpos <> 5 + extendedpassword Then
            solution = ShortenList(newwords, chars, letterpos, extendedpassword)
        End If
#Disable Warning BC42104 ' Variable is used before it has been assigned a value
        Return solution
#Enable Warning BC42104 ' Variable is used before it has been assigned a value
    End Function
    Sub Solvecombinationlock()
        Dim num1, num2, num3 As Integer
        num1 = ((CInt(Mid(serial, 6)) + solvedmodules) + batteries) Mod 20
        num2 = modulecount + solvedmodules Mod 20
        num3 = (num1 + num2) Mod 20
        voice.Speak("Right to " & num1)
        voice.Speak("Then Left to " & num2)
        voice.Speak("Then Right to " & num3)
    End Sub
    Sub Solvefastmath()
        Dim extratoadd As Integer
        Dim pair As String
        If inds(7).present = True And inds(7).lit = True Then
            extratoadd += 20
        End If
        If ports(0).count <> 0 Then
            extratoadd += 14
        End If
        If Checkforletterinserial("f") Or Checkforletterinserial("a") Or Checkforletterinserial("s") Or Checkforletterinserial("t") Then
            extratoadd -= 5
        End If
        If ports(3).count <> 0 Then
            extratoadd += 27
        End If
        If batteries > 3 Then
            extratoadd -= 15
        End If
        If extratoadd >= 100 Then
            extratoadd = extratoadd Mod 100
        ElseIf extratoadd < 0 Then
            extratoadd += 50
        End If
        takespeech("Enter first pair: ", pair)
        'A=65
    End Sub
    Sub Solvemaze()
        Dim mazeno As Integer = 3
        Dim mazes(8, 10, 10) As Char
        Dim startx, starty, endx, endy As Integer
        Dim movestring As String = ""
        Initializemazes(mazes)
        Displaymaze(mazes, mazeno, 0)
        Inputstartandend(startx, starty, endx, endy)
        movestring = Findroute(mazes, mazeno, startx, starty, endx, endy, movestring)
        Displaymovestring(movestring)
    End Sub
    Sub Displaymovestring(ByVal movestring As String)
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
    Sub Initializemazes(ByRef mazes(,,) As Char)
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
        mazes(2, 3, 2) = "S"
        mazes(2, 3, 10) = "S"
        For i = 0 To 10
            mazes(2, 4, i) = "S"
        Next
        mazes(2, 4, 1) = "*"
        mazes(2, 5, 0) = "S"
        For i = 1 To 10
            mazes(2, 5, i) = "*"
        Next
        For i = 0 To 10
            mazes(2, 6, i) = "S"
        Next
        mazes(2, 6, 7) = "*"
        For i = 0 To 10
            mazes(2, 7, i) = "*"
        Next
        mazes(2, 7, 6) = "S"
        mazes(2, 7, 8) = "S"
        For i = 0 To 10
            mazes(2, 8, i) = "S"
        Next
        mazes(2, 8, 7) = "S"
        For i = 1 To 9
            mazes(2, 9, i) = "*"
        Next
        mazes(2, 9, 0) = "S"
        mazes(2, 9, 10) = "S"
        For i = 0 To 10
            mazes(2, 10, i) = "S"
        Next
        mazes(2, 6, 4) = "O"
        mazes(2, 10, 4) = "O"
        'maze 4
        For i = 0 To 10
            mazes(3, 0, i) = "S"
        Next
        For i = 0 To 10
            mazes(3, 1, i) = "*"
        Next
        mazes(3, 1, 0) = "S"
        mazes(3, 1, 2) = "S"
        mazes(3, 1, 10) = "S"
        For i = 0 To 10 Step 2
            For j = 2 To 9
                mazes(3, j, i) = "S"
                mazes(3, 2, i + 1) = "*"
            Next
        Next
        mazes(3, 2, 7) = "S"
        mazes(3, 2, 9) = "S"
        mazes(3, 3, 8) = "*"
        mazes(3, 3, 10) = "*"
        mazes(3, 4, 7) = "S"
        mazes(3, 5, 0) = "*"
        mazes(3, 5, 6) = "*"
        mazes(3, 6, 5) = "S"
        mazes(3, 8, 1) = "S"
        mazes(3, 9, 0) = "*"
        mazes(3, 9, 2) = "*"
        mazes(3, 9, 6) = "*"
        For i = 0 To 10
            mazes(3, 10, i) = "S"
        Next
    End Sub
    Sub Displaymaze(ByVal mazes(,,) As Char, ByVal mazeno As Integer, ByVal displayhidden As Integer)
        For i = 0 To 10 Step displayhidden + 1
            For j = 0 To 10 Step displayhidden + 1
                voice.Speak(mazes(mazeno, j, 10 - i) & "  ")
            Next
            voice.Speak("")
        Next
    End Sub
    Function Checkcoord(ByVal coords As String)
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
    Sub Inputstartandend(ByRef startx As Integer, ByRef starty As Integer, ByRef endx As Integer, ByRef endy As Integer)
        Dim startcoords As String
        Dim valid = False
        Do
            takespeech("Enter Starting Co-Ordinates: ", startcoords)
            If Checkcoord(startcoords) Then
                valid = True
            End If
        Loop Until valid = True
        starty = 2 * (CInt(Mid(startcoords, 3, 1)) - 1)
        startx = 2 * (CInt(Mid(startcoords, 1, 1)) - 1)
        Dim endcoords As String
        valid = False
        Do
            takespeech("Enter Ending Co-Ordinates: ", endcoords)
            If Checkcoord(endcoords) Then
                valid = True
            End If
        Loop Until valid = True
        endy = 2 * (CInt(Mid(endcoords, 3, 1)) - 1)
        endx = 2 * (CInt(Mid(endcoords, 1, 1)) - 1)
    End Sub
    Function Findroute(ByVal mazes(,,) As Char, ByVal mazeno As Integer, ByRef currentposx As Integer, ByRef currentposy As Integer, ByVal endx As Integer, ByVal endy As Integer, ByRef movestring As String)
        Dim endfound = Checkend(currentposx, currentposy, endx, endy)
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
                movestring &= "U"
                currentposy += 2
                validmove = True
                movestring = Findroute(mazes, mazeno, currentposx, currentposy, endx, endy, movestring)
            End If
        End If
        endfound = Checkend(currentposx, currentposy, endx, endy)
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
                movestring &= "R"
                currentposx += 2
                validmove = True
                movestring = Findroute(mazes, mazeno, currentposx, currentposy, endx, endy, movestring)
            End If
        End If
        endfound = Checkend(currentposx, currentposy, endx, endy)
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
                movestring &= "D"
                currentposy -= 2
                validmove = True
                movestring = Findroute(mazes, mazeno, currentposx, currentposy, endx, endy, movestring)
            End If
        End If
        endfound = Checkend(currentposx, currentposy, endx, endy)
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
                movestring &= "L"
                currentposx -= 2
                validmove = True
                movestring = Findroute(mazes, mazeno, currentposx, currentposy, endx, endy, movestring)
            End If
        End If
        endfound = Checkend(currentposx, currentposy, endx, endy)
        If endfound = True Then
            Return movestring
        End If
        If validmove = False Then
            mazes(mazeno, currentposx, currentposy) = "X"
            Select Case movestring(movestring.Length - 1)
                Case "U"
                    mazes(mazeno, currentposx, currentposy - 1) = "X"
                    currentposy -= 2
                Case "R"
                    mazes(mazeno, currentposx - 1, currentposy) = "X"
                    currentposx -= 2
                Case "D"
                    mazes(mazeno, currentposx, currentposy + 1) = "X"
                    currentposy = currentposy + 2
                Case "L"
                    mazes(mazeno, currentposx + 1, currentposy) = "X"
                    currentposx = currentposx + 2
            End Select
            movestring = Mid(movestring, 1, movestring.Length - 1)
        End If
        movestring = Findroute(mazes, mazeno, currentposx, currentposy, endx, endy, movestring)
        Return movestring
    End Function
    Function Checkend(ByVal currentposx As Integer, ByVal currentposy As Integer, ByVal endx As Integer, ByVal endy As Integer)
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
            convertedsymbols(i) = Converttonums(symbols, i)
        Next
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
                voice.Speak(reference(j, i) & " ")
                If reference(j, i) < 10 Then
                    voice.Speak(" ")
                End If
            Next
            voice.Speak("")
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
            takespeech("Input Symbol " & i + 1 & ": ", symbols(i))
        Next
    End Sub
    Function Converttonums(ByVal symbols() As String, ByVal symbolnumber As Integer)
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
#Disable Warning BC42108 ' Variable is passed by reference before it has been assigned a value
        inputstartchord(chord1)
#Enable Warning BC42108 ' Variable is passed by reference before it has been assigned a value
        Chordoffsets(chord1, notes)
        Dim startquality = findbaseind(chord1.note.note1)
        Dim startandquality = findquality(chord1, chords)
        If startandquality <> "CNF" Then
            Dim startnote As String = Mid(startandquality, 1, 1)
            Dim quality As String = Mid(startandquality, 2)
            findanswerchord(startnote, quality, chord1, chords)
        End If
    End Sub
    Sub findanswerchord(start As Integer, quality As String, chord As TChord, chords() As TChord)
#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
        Dim startnote As String = setstartnote(startnote, start, chord)
#Enable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
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
        Return "ERROR"
    End Function
    Function findqualityindex(ByVal quality As String, chords() As TChord)
        For i = 0 To chords.Length - 1
            If chords(i).quality = quality Then
                Return i
            End If
        Next
        Return "Error"
    End Function
    Sub Displayoffsets(ByVal chord As TChord)
        voice.Speak(chord.offsets.offset1 & ", " & chord.offsets.offset2 & ", " & chord.offsets.offset3 & ", " & chord.offsets.offset4)
    End Sub
    Sub Chordoffsets(ByRef chord As TChord, ByVal notes() As String)
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
            UCase(takespeech("Enter note " & i + 1 & ": ", temp))
            Select Case i
                Case 0
                    chord.note.note1 = UCase(temp)
                Case 1
                    chord.note.note2 = UCase(temp)
                Case 2
                    chord.note.note3 = UCase(temp)
                Case 3
                    chord.note.note4 = UCase(temp)
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
    Sub solveskewed()
        Dim primes() As Integer = {2, 3, 5, 7, 11, 13, 17, 23, 29}
        Dim fibonacci() As Integer = {1, 1, 2, 3, 5, 8, 13, 21, 34, 55, 89, 144}
        Dim slot(2) As Integer
        Dim originalslots(2) As Integer

        For i = 0 To 2
            takespeech("Enter digit in slot " & i, slot(i))
            originalslots(i) = slot(i)
        Next
        'All Slots
        For i = 0 To 2
            If slot(i) = 2 Then
                slot(i) = 5
            ElseIf slot(i) = 7 Then
                slot(i) = 0
            End If '011
            slot(i) = slot(i) + Calclitinds() - Calcunlitinds()
            If slot(i) Mod 3 = 0 Then
                slot(i) += 4
            ElseIf slot(i) > 7 Then
                slot(i) *= 2
            ElseIf (slot(i)) < 3 And (slot(i) Mod 2 = 0) Then
                slot(i) /= 2
            ElseIf ports(4).count >= 1 Or ports(2).count >= 1 Then
            Else slot(i) += batteries
            End If
        Next
        '1St Slot 
        If slot(0) Mod 2 = 0 And slot(0) > 5 Then
            slot(0) /= 2
        ElseIf isprime(primes, slot(0)) Then
            slot(0) += CInt(Mid(serial, 5, 1))
        ElseIf ports(1).count >= 1 Then
            slot(0) *= -1
        ElseIf originalslots(0) Mod 2 = 1 Then
        Else slot(0) -= 2 '211
        End If

        '2nd Slot 
        If Not inds(10).lit Then
            If slot(1) = 0 Then
                slot(1) += originalslots(0)
            ElseIf isfibonacci(slot(1), fibonacci) Then
                slot(1) += nextfib(slot(1), fibonacci)
            ElseIf slot(1) >= 7 Then
                slot(1) += 4
            Else slot(1) *= 3
            End If
        End If

        '3rd Slot 
        If serial Then
            slot(2) += largestserial(serial)
        ElseIf sameasorig(originalslots, slot(2)) Then
        ElseIf slot(2) >= 5 Then
            slot(2) += Addbinary(originalslots(2))
        Else slot(2) += 1
        End If
        For i = 0 To 2
            slot(i) = slot(i) Mod 10
        Next
        For i = 0 To 2
            voice.Speak(slot(i))
        Next
    End Sub
    Function isfibonacci(ByVal number As Integer, ByVal fibonacci() As Integer)
        For i = 0 To fibonacci.Length
            If number = fibonacci(i) Then
                Return True
            End If
        Next
        Return False
    End Function
    Function nextfib(ByVal number As Integer, ByVal fibonacci() As Integer)
        For i = 0 To fibonacci.Length
            If number = fibonacci(i) Then
                Return fibonacci(i + 1)
            End If
        Next
        Return "ERROR"
    End Function
    Function largestserial(serial)
        Dim serialtemp, largest As Integer
        Dim num As Boolean
        For i = 0 To 5
            Try
                serialtemp = CInt(Mid(serial, i, 1))
                If serialtemp > largest Then
                    largest = serialtemp
                End If
            Catch ex As Exception
            End Try
        Next
        Return largest
    End Function
    Function sameasorig(ByVal slot() As Integer, ByVal current As Integer)
        For i = 0 To 2
            If slot(i) = current Then
                Return True
            End If
        Next
        Return False
    End Function
    Function Addbinary(orig)
        Dim binary() As Integer
        Dim loopcount As Integer = 0
        Do
            ReDim Preserve binary(loopcount)
            binary(loopcount) = orig Mod 2
            orig -= binary(loopcount) * 2 ^ loopcount
            orig /= 2
            loopcount += 1
        Loop Until orig / 2 < 1
        ReDim Preserve binary(loopcount)
        binary(loopcount) = orig
        Dim total As Integer = 0
        For i = 0 To loopcount
            total += binary(i)
        Next
        Return total
    End Function
    Function isprime(primes() As Integer, number As Integer)
        For i = 0 To primes.Length - 1
            If primes(i) = number Then
                Return True
            End If
        Next
        Return False
    End Function
    Sub convertledstopropperenglish(ByRef leds(,) As String)
        For i = 0 To 1
            For j = 0 To 3
                If leds(i, j) = "gray" Then
                    leds(i, j) = "grey"
                End If
            Next
        Next
    End Sub
    Sub SolveColourMath()
        Dim leftnum, rightnum As Integer
        Dim numsolution As Integer
        Dim leds(1, 3) As String
        Dim centre As Toperator
        Dim coloursolution(3) As String
        InputLed("left", leds)
        convertledstopropperenglish(leds)
        leftnum = converttonumber(leds, "left")
        enteroperatordata(centre)
        If centre.colour = "green" Then
            InputLed("right", leds)
            rightnum = converttonumber(leds, "right")
        Else
            Select Case batteries
                Case 0, 1
                    rightnum = serialnums(0) & Calcunlitinds() & 9 & ports(3).count
                Case 2, 3
                    rightnum = 0 & ports(2).count & serialletters.Length & serialnums(serialnums.Length - 1)
                Case 4, 5
                    rightnum = serialvowels & holders & ports(0).count & 4
                Case Else
                    rightnum = ports(5).count & 5 & (serialletters.Length - serialvowels) & Calclitinds()
            End Select

        End If
        numsolution = findnumsolution(leftnum, rightnum, centre)
        convertbacktocolour(numsolution, coloursolution)
        voice.Speak(coloursolution(0))
        For i = 1 To 3
            voice.Speak(", " & coloursolution(i))
        Next
    End Sub
    Sub convertbacktocolour(ByRef numsolution As Integer, ByRef coloursolution() As String)
        Dim digit1, digit2, digit3, digit4 As Integer
        digit1 = numsolution \ 1000
        digit2 = (numsolution Mod 1000) \ 100
        digit3 = (numsolution Mod 100) \ 10
        digit4 = numsolution Mod 10
        Select Case digit1
            Case 0
                coloursolution(0) = "Grey"
            Case 1
                coloursolution(0) = "Green"
            Case 2
                coloursolution(0) = "Orange"
            Case 3
                coloursolution(0) = "White"
            Case 4
                coloursolution(0) = "Purple"
            Case 5
                coloursolution(0) = "Blue"
            Case 6
                coloursolution(0) = "Magenta"
            Case 7
                coloursolution(0) = "Black"
            Case 8
                coloursolution(0) = "Yellow"
            Case 9
                coloursolution(0) = "Red"
        End Select
        Select Case digit2
            Case 0
                coloursolution(1) = "Blue"
            Case 1
                coloursolution(1) = "Green"
            Case 2
                coloursolution(1) = "Black"
            Case 3
                coloursolution(1) = "Purple"
            Case 4
                coloursolution(1) = "Magenta"
            Case 5
                coloursolution(1) = "Red"
            Case 6
                coloursolution(1) = "Grey"
            Case 7
                coloursolution(1) = "Yellow"
            Case 8
                coloursolution(1) = "Orange"
            Case 9
                coloursolution(1) = "White"
        End Select
        Select Case digit3
            Case 0
                coloursolution(2) = "Magenta"
            Case 1
                coloursolution(2) = "Yellow"
            Case 2
                coloursolution(2) = "Blue"
            Case 3
                coloursolution(2) = "Grey"
            Case 4
                coloursolution(2) = "Red"
            Case 5
                coloursolution(2) = "Black"
            Case 6
                coloursolution(2) = "Green"
            Case 7
                coloursolution(2) = "Purple"
            Case 8
                coloursolution(2) = "Orange"
            Case 9
                coloursolution(2) = "White"
        End Select
        Select Case digit4
            Case 0
                coloursolution(3) = "Grey"
            Case 1
                coloursolution(3) = "Blue"
            Case 2
                coloursolution(3) = "Purple"
            Case 3
                coloursolution(3) = "Red"
            Case 4
                coloursolution(3) = "Yellow"
            Case 5
                coloursolution(3) = "Magenta"
            Case 6
                coloursolution(3) = "Black"
            Case 7
                coloursolution(3) = "Orange"
            Case 8
                coloursolution(3) = "Green"
            Case 9
                coloursolution(3) = "White"
        End Select
    End Sub
    Function findnumsolution(ByVal num1 As Integer, ByVal num2 As Integer, ByVal op As Toperator)
        Select Case UCase(op.action)
            Case "A"
                Return (num1 + num2) Mod 10000
            Case "S"
                Return Math.Abs(num1 - num2)
            Case "M"
                Return (num1 * num2) Mod 10000
            Case "D"
                Return num1 \ num2
        End Select
        'if no sign has been entered displays the message to contact the developer
        voice.Speak("This Message Should not appear. If it does, Contact Chris ASAP")
        Return 0
    End Function
    Function converttonumber(ByVal leds(,) As String, ByVal side As String)
        Dim sidenum As Integer = side.Length - 4
        Dim num As Integer = 0
        Select Case sidenum
            Case 0
                For i = 0 To 3
                    num *= 10
                    Select Case i
                        Case 0
                            Select Case LCase(leds(sidenum, i))
                                Case "blue"
                                    num += 6
                                Case "green"
                                    num += 1
                                Case "purple"
                                    num += 2
                                Case "yellow"
                                    num += 4
                                Case "white"
                                    num += 9
                                Case "magenta"
                                    num += 0
                                Case "red"
                                    num += 8
                                Case "orange"
                                    num += 5
                                Case "grey"
                                    num += 3
                                Case "black"
                                    num += 7
                            End Select
                        Case 1
                            Select Case LCase(leds(sidenum, i))
                                Case "blue"
                                    num += 8
                                Case "green"
                                    num += 1
                                Case "purple"
                                    num += 9
                                Case "yellow"
                                    num += 4
                                Case "white"
                                    num += 3
                                Case "magenta"
                                    num += 6
                                Case "red"
                                    num += 0
                                Case "orange"
                                    num += 5
                                Case "grey"
                                    num += 7
                                Case "black"
                                    num += 2
                            End Select
                        Case 2
                            Select Case LCase(leds(sidenum, i))
                                Case "blue"
                                    num += 4
                                Case "green"
                                    num += 1
                                Case "purple"
                                    num += 9
                                Case "yellow"
                                    num += 7
                                Case "white"
                                    num += 0
                                Case "magenta"
                                    num += 2
                                Case "red"
                                    num += 5
                                Case "orange"
                                    num += 3
                                Case "grey"
                                    num += 8
                                Case "black"
                                    num += 6
                            End Select
                        Case 3
                            Select Case LCase(leds(sidenum, i))
                                Case "blue"
                                    num += 6
                                Case "green"
                                    num += 8
                                Case "purple"
                                    num += 7
                                Case "yellow"
                                    num += 5
                                Case "white"
                                    num += 4
                                Case "magenta"
                                    num += 9
                                Case "red"
                                    num += 1
                                Case "orange"
                                    num += 3
                                Case "grey"
                                    num += 0
                                Case "black"
                                    num += 2
                            End Select
                    End Select
                Next
            Case 1
                For i = 0 To 3
                    num *= 10
                    Select Case i
                        Case 0
                            Select Case LCase(leds(sidenum, i))
                                Case "blue"
                                    num += 0
                                Case "green"
                                    num += 6
                                Case "purple"
                                    num += 5
                                Case "yellow"
                                    num += 4
                                Case "white"
                                    num += 3
                                Case "magenta"
                                    num += 7
                                Case "red"
                                    num += 9
                                Case "orange"
                                    num += 8
                                Case "grey"
                                    num += 1
                                Case "black"
                                    num += 2
                            End Select
                        Case 1
                            Select Case LCase(leds(sidenum, i))
                                Case "blue"
                                    num += 2
                                Case "green"
                                    num += 9
                                Case "purple"
                                    num += 8
                                Case "yellow"
                                    num += 0
                                Case "white"
                                    num += 5
                                Case "magenta"
                                    num += 3
                                Case "red"
                                    num += 4
                                Case "orange"
                                    num += 7
                                Case "grey"
                                    num += 1
                                Case "black"
                                    num += 6
                            End Select
                        Case 2
                            Select Case LCase(leds(sidenum, i))
                                Case "blue"
                                    num += 5
                                Case "green"
                                    num += 0
                                Case "purple"
                                    num += 6
                                Case "yellow"
                                    num += 4
                                Case "white"
                                    num += 2
                                Case "magenta"
                                    num += 7
                                Case "red"
                                    num += 9
                                Case "orange"
                                    num += 3
                                Case "grey"
                                    num += 8
                                Case "black"
                                    num += 1
                            End Select
                        Case 3
                            Select Case LCase(leds(sidenum, i))
                                Case "blue"
                                    num += 5
                                Case "green"
                                    num += 4
                                Case "purple"
                                    num += 2
                                Case "yellow"
                                    num += 9
                                Case "white"
                                    num += 8
                                Case "magenta"
                                    num += 6
                                Case "red"
                                    num += 7
                                Case "orange"
                                    num += 1
                                Case "grey"
                                    num += 3
                                Case "black"
                                    num += 0
                            End Select
                    End Select
                Next
        End Select
        Return num
    End Function
    Sub InputLed(ByVal side As String, ByRef leds(,) As String)
        'collects colours of leds on 1 side
        Dim sidenum As Integer = side.Length - 4
        'side 0 = left
        'side 1 = right
        For i = 0 To 3
            Dim validinput = False
            Do

                Dim temp As String
                LCase(takespeech("Enter LED Colour for LED " & i + 1 & " on the " & side & ": ", temp))
                If checkcolour(temp) Then
                    leds(sidenum, i) = temp
                    validinput = True
                Else
                    voice.Speak("Invalid Colour, please re-enter.")
                End If
            Loop Until validinput = True
        Next
    End Sub
    Function checkcolour(ByVal colour As String)
        'ensures colours are valid
        Select Case colour
            Case "blue"
                Return True
            Case "green"
                Return True
            Case "purple"
                Return True
            Case "yellow"
                Return True
            Case "white"
                Return True
            Case "magenta"
                Return True
            Case "red"
                Return True
            Case "orange"
                Return True
            Case "gray"
                Return True
            Case "grey"
                Return True
            Case "black"
                Return True
            Case Else
                Return False
        End Select
        Return "Error"
    End Function
    Sub enteroperatordata(ByRef op As Toperator)
        enteroperatorcolour(op.colour)
        enteroperatortype(op.action)
    End Sub
    Sub enteroperatorcolour(ByRef colourholder As String)
        Dim validcolour As Boolean = False
        Dim temp As String
        Do
            voice.Speak("Enter colour of the operator in the centre: ")
            LCase(takespeech("Enter colour of the operator in the centre: ", temp))
            Select Case temp
                Case "red"
                    colourholder = "red"
                    validcolour = True
                Case "green"
                    colourholder = "green"
                    validcolour = True
                Case Else
                    voice.Speak("Invalid Colour please re-enter.")
                    validcolour = False
            End Select
        Loop Until validcolour = True
    End Sub
    Sub enteroperatortype(ByRef actionholder As String)
        Dim validoperation As Boolean = False
        Dim temp As String
        Do
            UCase(takespeech("Enter letter in the centre: ", temp))
            Select Case temp
                Case "A", "S", "M", "D"
                    actionholder = temp
                    validoperation = True
                Case Else
                    voice.Speak("Invalid Letter please re-enter.")
                    validoperation = False
            End Select
        Loop Until validoperation = True
    End Sub
    Sub solvesimonsays()
        Dim sscomplete As Boolean = False
        Dim colourstring As String = ""
        Dim completecheck As Boolean = True
        Do
            Dim flash As String = inputnewflashcolour()
            Select Case LCase(flash)
                Case "finish"
                    sscomplete = True
                Case Else
                    colourstring &= converttooutputcolour(flash) & " "
                    voice.Speak(colourstring)
            End Select
        Loop Until sscomplete = True

    End Sub
    Function converttooutputcolour(newflash)
        Dim strikes As Integer = 0
        Dim vowelpresent As Boolean = True
        Select Case vowelpresent
            Case True
                Select Case strikes
                    Case 0
                        Select Case newflash
                            Case "red"
                                Return "blue"
                            Case "blue"
                                Return "red"
                            Case "yellow"
                                Return "green"
                            Case "green"
                                Return "yellow"
                        End Select
                    Case 1
                        Select Case newflash
                            Case "red"
                                Return "yellow"
                            Case "blue"
                                Return "green"
                            Case "yellow"
                                Return "red"
                            Case "green"
                                Return "blue"
                        End Select
                    Case Else
                        Select Case newflash
                            Case "red"
                                Return "green"
                            Case "blue"
                                Return "red"
                            Case "yellow"
                                Return "blue"
                            Case "green"
                                Return "yellow"
                        End Select
                End Select
            Case False
                Select Case strikes
                    Case 0
                        Select Case newflash
                            Case "red"
                                Return "blue"
                            Case "blue"
                                Return "yellow"
                            Case "yellow"
                                Return "red"
                            Case "green"
                                Return "green"
                        End Select
                    Case 1
                        Select Case newflash
                            Case "red"
                                Return "red"
                            Case "blue"
                                Return "blue"
                            Case "yellow"
                                Return "green"
                            Case "green"
                                Return "yellow"
                        End Select
                    Case Else
                        Select Case newflash
                            Case "red"
                                Return "yellow"
                            Case "blue"
                                Return "green"
                            Case "yellow"
                                Return "red"
                            Case "green"
                                Return "blue"
                        End Select
                End Select
        End Select
        Return "this message should not appearm if it does, contact chris ASAP"
    End Function
    Function inputnewflashcolour()
        Dim temp As String
        Dim valid As Boolean = False
        Do
            takespeech("Enter Last Flash in Sequence, Say Finish to End: ", temp)
            Select Case LCase(temp)
                Case "red", "blue", "yellow", "green", "finish"
                    valid = True
                Case Else
                    voice.Speak("Invalid colour, please re-enter.")
            End Select
        Loop Until valid = True
        Return temp
    End Function
    Sub Inputletters(ByRef chars(,) As Char, ByRef extendedpassword As Integer)
        'enter letters for passwords and extended password
        For i = 0 To (4 + extendedpassword)
            For j = 0 To 5
                takespeech("Enter Char number " & j & " In position " & i & ": ", chars(i, j))
            Next
        Next
    End Sub
    Function Inputwire(ByVal pos As Integer)
        'stores wire for complicated wires
        Dim temp As String
        Do
            voice.Speak("Enter wire connected to " & pos & " (say pass if no wire): ")
            LCase(takespeech("Enter wire connected to " & pos & " (press enter if no wire): ", temp))
        Loop Until temp = "red" Or temp = "blue" Or temp = "black" Or temp = "pass"
        Return temp
    End Function

    Function Findlastofcolour(ByVal colour, ByVal wirecount, ByVal wires())
        'finds the last wire of a colour
        For i = wirecount - 1 To 0 Step -1
            If wires(i) = colour Or wires(i) = colour(0) Then
                Return i
            End If
        Next
        Return "ERROR"
    End Function
    Function Searchforwire(ByVal colour, ByVal wirecount, ByVal wires())
        'finds count of a particular colour of wire
        Dim count As Integer = 0
        For i = 0 To wirecount - 1
            If wires(i) = colour Or wires(i) = colour(0) Then
                count += 1
            End If
        Next
        Return count
    End Function
    Sub Inputwires(ByRef wires() As String, ByRef wirecount As Integer)
        'stores the colour of each wire
        For i = 0 To wirecount - 1
            LCase(takespeech("Colour of wire number " & i + 1 & ": ", wires(i)))
        Next
    End Sub
End Module