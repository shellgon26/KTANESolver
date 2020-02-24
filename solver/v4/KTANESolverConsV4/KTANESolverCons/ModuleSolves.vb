Module ModuleSolves

    'Creates a tree for navigation in complicated wires
    Class TreeNode
        'defines each node in the tree as having a value, and a left and right node branching off of it
        Public Value As Object
        Public leftNode As TreeNode
        Public RightNode As TreeNode
    End Class
    Class Tree
        'creates a root node for the tree
        Public _root As TreeNode
        Sub New(ByVal rootValue As Object)
            _root = GetNode(rootValue)
        End Sub

        Private Function GetNode(ByVal value As Object) As TreeNode
            'creates a new node with value passed to the function
            Dim node As New TreeNode()
            node.Value = value
            Return node
        End Function
        Public Sub AddtoTree(ByVal value As Double)
            'adds the passed value onto the tree
            'if a value matches, a solution for complicated wires has been found and the correct wire cut procedure is used
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
                        Cuttypec()
                    Case 2, 5, 15
                        Cuttyped()
                    Case 3, 10, 11
                        Cuttypeb()
                    Case 4, 8, 12, 14
                        Cuttypes()
                    Case 6, 7, 13
                        Cuttypep()
                End Select
            ElseIf currentNode.Value < value Then
                currentNode.RightNode = GetNode(value)
            Else
                currentNode.leftNode = GetNode(value)
            End If
        End Sub
        Sub InOrderTraverse(ByVal node As TreeNode)
            'outputs each node in order of ascending value --For Testing purposes only--
            If Not node Is Nothing Then
                InOrderTraverse(node.leftNode)
                saymessage(node.Value)
                InOrderTraverse(node.RightNode)
            End If
        End Sub
    End Class
    'functions used for wire cutting procedures
    Sub Cuttypec()
        'cut wire regardless
        saymessage("Cut")
    End Sub
    Sub Cuttyped()
        'Do not cut wire regardless
        saymessage("Don't Cut")
    End Sub
    Sub Cuttypep()
        'only cut if a parallel port is present
        Dim parallel As Boolean
        If parallel Then
            Cuttypec()
        Else
            Cuttyped()
        End If
    End Sub
    Sub Cuttypes()
        'only cut if the last serial digit is even
        Dim lastserialnum As Integer
        If Checkeven(lastserialnum) Then
            Cuttypec()
        Else
            Cuttyped()
        End If
    End Sub
    Sub Cuttypeb()
        'only cut if there is more than one battery
        Dim batteries As Integer
        If batteries >= 2 Then
            Cuttypec()
        Else
            Cuttyped()
        End If
    End Sub
    Sub Solvecompwires()
        'solves complicated wires module
        Dim btree As New Tree(7.5)
        Initialisebinarytree(btree)
        Dim desciptionstring As String = "red, led, star"
        Dim keywords = {"red", "blue", "led", "star", "nothing"}
        Dim wirescore As Integer = 0
        Dim validdesc, breakdesc As Boolean
        Do
            takespeech("Enter Wire Details, say done when there are no wires left: ", desciptionstring)
            wirescore = 0
            desciptionstring = LCase(desciptionstring)
            validdesc = False
            breakdesc = False
            For j = 0 To keywords.Length - 1
                For i = 0 To desciptionstring.Length - keywords(j).Length
                    If breakdesc = False Then
                        Dim checkword = Mid(desciptionstring, i + 1, keywords(j).Length)
                        If checkword = keywords(j) Then
                            If keywords(j) = "nothing" Then
                                Cuttypec()
                            Else
                                wirescore += 2 ^ (3 - j)
                            End If
                            validdesc = True
                        ElseIf checkword = "complete" Or checkword = "done" Then
                            validdesc = True
                        ElseIf checkword = "no " Then
                            If TestMode = 1 Then
                                saymessage("do not mention what is not present, only what is")
                                Console.WriteLine()
                            Else
                                saymessage("do not mention what is not present, only what is")
                            End If
                            breakdesc = True
                        End If
                    End If
                Next
            Next
            If validdesc = False Then
                saymessage("Invalid Description")
            Else
                btree.AddtoTree(wirescore)
            End If
        Loop Until LCase(desciptionstring) = "complete" Or LCase(desciptionstring) = "done"

    End Sub
    Sub Initialisebinarytree(ByRef btree As Tree)
        'sets the values for the binary tree
        Dim values() As Integer = {1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15}
        For i = 0 To values.Length - 1
            btree.AddtoTree(values(i))
        Next
    End Sub
    Sub Solvewires()
        'solves the basic wires module
        Dim wirecount As Integer
        Dim temp1 As String
        takespeech("Number of Wires: ", temp1, True)
        wirecount = temp1
        Dim wires(wirecount - 1) As String
        Inputwires(wires, wirecount)
        Select Case wirecount
            Case 3
                If Searchforwire("red", wirecount, wires) = 0 Then
                    saymessage("Cut wire 2")
                ElseIf wires(wirecount - 1) = "white" Then
                    saymessage("Cut wire " & wirecount)
                ElseIf Searchforwire("blue", wirecount, wires) > 1 Then
                    saymessage("Cut wire " & 1 + Findlastofcolour("blue", wirecount, wires))
                Else
                    saymessage("Cut wire 3")
                End If
            Case 4
                If (Searchforwire("red", wirecount, wires) > 1) And Checkeven(CInt(Mid(serial, 6))) = False Then
                    saymessage("Cut wire " & 1 + Findlastofcolour("red", wirecount, wires))
                ElseIf wires(wirecount - 1) = "yellow" Or wires(wirecount - 1) = "y" And Searchforwire("red", wirecount, wires) = 0 Then
                ElseIf Searchforwire("blue", wirecount, wires) = 1 Then
                    saymessage("Cut wire 1")
                ElseIf Searchforwire("yellow", wirecount, wires) > 1 Then
                    saymessage("Cut wire " & wirecount)
                Else
                    saymessage("Cut wire 2")
                End If
            Case 5
                If wires(wirecount - 1) = "black" Or wires(wirecount - 1) = "b" And Checkeven(CInt(Mid(serial, 6))) = False Then
                    saymessage("Cut wire 4")
                ElseIf Searchforwire("red", wirecount, wires) = 1 And Searchforwire("yellow", wirecount, wires) > 1 Then
                    saymessage("Cut wire 1")
                ElseIf Searchforwire("black", wirecount, wires) = 0 Then
                    saymessage("Cut wire 2")
                Else saymessage("Cut wire 1")
                End If
            Case 6
                If Searchforwire("yellow", wirecount, wires) = 0 And Checkeven(CInt(Mid(serial, 6))) = False Then
                    saymessage("Cut wire 3")
                ElseIf Searchforwire("yellow", wirecount, wires) = 1 And Searchforwire("white", wirecount, wires) > 1 Then
                    saymessage("Cut wire 4")
                ElseIf Searchforwire("red", wirecount, wires) = 0 Then
                    saymessage("Cut wire " & wirecount)
                Else saymessage("Cut wire 4")
                End If

            Case Else
                saymessage("ERROR")
        End Select
    End Sub
    Sub Solvewireseq()
        'solves the wire sequences module
        Dim red, blue, black As Integer
        Dim currentwire
        For i = 1 To 12
            currentwire = Inputwire(i)
            Select Case currentwire
                Case "red"
                    red += 1
                    Select Case red
                        Case 1
                            saymessage("Cut if connected to C")
                        Case 2
                            saymessage("Cut if connected to B")
                        Case 3
                            saymessage("Cut if connected to A")
                        Case 4
                            saymessage("Cut if connected to A or C")
                        Case 5
                            saymessage("Cut if connected to B")
                        Case 6
                            saymessage("Cut if connected to A or C")
                        Case 7
                            saymessage("Cut")
                        Case 8
                            saymessage("Cut if connected to A or B")
                        Case 9
                            saymessage("Cut if connected to B")
                        Case Else
                            saymessage("ERROR")
                    End Select
                Case "blue"
                    blue += 1
                    Select Case blue
                        Case 1
                            saymessage("Cut if connected to B")
                        Case 2
                            saymessage("Cut if connected to A or C")
                        Case 3
                            saymessage("Cut if connected to B")
                        Case 4
                            saymessage("Cut if connected to A")
                        Case 5
                            saymessage("Cut if connected to B")
                        Case 6
                            saymessage("Cut if connected to B or C")
                        Case 7
                            saymessage("Cut if connected to C")
                        Case 8
                            saymessage("Cut if connected to A or C")
                        Case 9
                            saymessage("Cut if connected to A")
                        Case Else
                            saymessage("ERROR")
                    End Select
                Case "black"
                    black += 1
                    Select Case black
                        Case 1
                            saymessage("Cut")
                        Case 2
                            saymessage("Cut if connected to A or C")
                        Case 3
                            saymessage("Cut if connected to B")
                        Case 4
                            saymessage("Cut if connected to A or C")
                        Case 5
                            saymessage("Cut if connected to B")
                        Case 6
                            saymessage("Cut if connected to B or C")
                        Case 7
                            saymessage("Cut if connected to A or B")
                        Case 8
                            saymessage("Cut if connected to C")
                        Case 9
                            saymessage("Cut if connected to C")
                        Case Else
                            saymessage("ERROR")
                    End Select
                Case ""
                Case Else
                    saymessage("ERROR")
            End Select
        Next
    End Sub
    Sub SolveButton()
        'solves the basic button module
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
                    saymessage("Tap")
                    solved = True
                Else
                    step2 = True
                    'need to look at strip colour

                End If
            Else
                saymessage("Hold")
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
                    saymessage("Release on " & result)
                    solved = True
                End If
            End If
        Loop Until solved = True
    End Sub
    Sub Solvepassword()
        'solves the basic password module
        Dim extendedpassword As Integer = 0
        Dim answer As String
        Dim words() As String = {"about", "after", "again", "below", "could", "every", "first", "found", "great", "house", "large", "learn", "never", "other", "place", "plant", "point", "right", "small", "sound", "spell", "still", "study", "their", "there", "these", "thing", "think", "three", "water", "where", "which", "world", "would", "write"}
        Dim chars(4, 5) As Char
        Inputletters(chars, extendedpassword)
        answer = ShortenList(words, chars, 0, extendedpassword)
        saymessage("The Answer is: " & answer)
    End Sub
    Function ShortenList(ByRef currentwords() As String, ByRef chars(,) As Char, ByRef letterpos As Integer, ByVal extendedpassword As Integer)
        'recursively shortnes the list of valid words by checking which words are possible to make
        Dim solution As String
        If currentwords.Length = 1 Then
            Return currentwords(0)
        End If
        Dim newwords(0) As String
        Dim loopend = currentwords.Length - 1
        For i = 0 To loopend
            Dim currentword = currentwords(i)
            Dim lettercheck = currentword(letterpos)
            If lettercheck = chars(letterpos, 0) Or lettercheck = chars(letterpos, 1) Or lettercheck = chars(letterpos, 2) Or lettercheck = chars(letterpos, 3) Or lettercheck = chars(letterpos, 4) Or lettercheck = chars(letterpos, 5) Then
                newwords.Append(currentwords(i))
            End If
        Next
        If newwords.Length = 0 Then
            Return "error"
        End If
        letterpos += 1
        If letterpos <> 5 + extendedpassword Then
            solution = ShortenList(newwords, chars, letterpos, extendedpassword)
        End If
        Return solution
    End Function
    Sub Solvecombinationlock()
        'solves combination lock module
        Dim num1, num2, num3 As Integer
        num1 = ((CInt(Mid(serial, 6)) + solvedmodules) + batteries) Mod 20
        num2 = modulecount + solvedmodules Mod 20
        num3 = (num1 + num2) Mod 20
        saymessage("Right to " & num1)
        saymessage("Then Left to " & num2)
        saymessage("Then Right to " & num3)
    End Sub
    Sub Solvefastmath()
        'solves fash math module --UNLIKLY TO BE FULLY FUNCTIONAL--
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
        'solves basic maze module --NOT FULLY FUNCTIONAL--
        Dim mazeno As Integer = 3
        Dim mazes(8, 10, 10) As Char
        Dim startx, starty, endx, endy As Integer
        Dim movestring As String = ""
        Initializemazes(mazes)
        Displaymaze(mazes, mazeno, 1)
        Inputstartandend(startx, starty, endx, endy)
        movestring = Findroute(mazes, mazeno, startx, starty, endx, endy, movestring)
        Displaymovestring(movestring)
    End Sub
    Sub Displaymovestring(ByVal movestring As String)
        'outputs the string of moves the user must take to complete the maze
        For i = 0 To movestring.Length - 1
            Select Case movestring(i)
                Case "U"
                    saymessage("Up")
                Case "R"
                    saymessage("Right")
                Case "D"
                    saymessage("Down")
                Case "L"
                    saymessage("Left")
            End Select
        Next
    End Sub
    Sub Initializemazes(ByRef mazes(,,) As Char)
        'initialise all the mazes --ONLY 4 OF 9 MAZES IMPLEMENTED--
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
        For i = 0 To 8 Step 2
            For j = 2 To 9
                mazes(3, j, i) = "S"
                mazes(3, 2, i + 1) = "*"
            Next
        Next
        For i = 2 To 9
            mazes(3, i, 10) = "S"
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
        'displays the maze --FOR TESTING ONLY--
        For i = 0 To 10 Step displayhidden + 1
            For j = 0 To 10 Step displayhidden + 1
                Console.Write(mazes(mazeno, j, 10 - i) & "  ")
                saymessage(mazes(mazeno, j, 10 - i) & "  ")
            Next
            Console.WriteLine("")
        Next
    End Sub
    Function Checkcoord(ByVal coords As String)
        'checks if the user input is a valid coordinate on the maze
        Dim test As Integer
        For i = 1 To 2
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
        'allows the user to input their starting coordinates and their endpoint coordinates 
        Dim startcoords As String
        Dim valid = False
        Do
            takespeech("Enter Starting Co-Ordinates: ", startcoords)
            If Checkcoord(startcoords) Then
                valid = True
            End If
        Loop Until valid = True
        starty = 2 * (CInt(Mid(startcoords, 2, 1)) - 1)
        startx = 2 * (CInt(Mid(startcoords, 1, 1)) - 1)
        Dim endcoords As String
        valid = False
        Do
            takespeech("Enter Ending Co-Ordinates: ", endcoords)
            If Checkcoord(endcoords) Then
                valid = True
            End If
        Loop Until valid = True
        endy = 2 * (CInt(Mid(endcoords, 2, 1)) - 1)
        endx = 2 * (CInt(Mid(endcoords, 1, 1)) - 1)
    End Sub
    Function Findroute(ByVal mazes(,,) As Char, ByVal mazeno As Integer, ByRef currentposx As Integer, ByRef currentposy As Integer, ByVal endx As Integer, ByVal endy As Integer, ByRef movestring As String)
        'recusively finds the solution to the maze
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
                    currentposy += 2
                Case "L"
                    mazes(mazeno, currentposx + 1, currentposy) = "X"
                    currentposx += 2
            End Select
            movestring = Mid(movestring, 1, movestring.Length - 1)
        End If
        movestring = Findroute(mazes, mazeno, currentposx, currentposy, endx, endy, movestring)
        Return movestring
    End Function
    Function Checkend(ByVal currentposx As Integer, ByVal currentposy As Integer, ByVal endx As Integer, ByVal endy As Integer)
        'checks if player position is the smae as the maze exit
        If currentposx = endx And currentposy = endy Then
            Return True
        End If
        Return False
    End Function
    Sub Solvekeypad()
        'solves keypad module
        Dim symbols(3) As String
        Dim convertedsymbols(3) As Integer
        Dim reference(5, 6) As Integer
        Initializereference(reference)
        'showtable(reference)
        Inputsymbols(symbols)
        For i = 0 To 3
            convertedsymbols(i) = Converttonums(symbols, i)
        Next
    End Sub
    Sub Findsymbolorder(symbols() As Integer, reference(,) As Integer, ByVal column As Integer)
        'findds the order to output the symbols in
        Dim foundcount As Integer = 0
        For i = 0 To 6
            For k = 0 To 3
                If reference(column, i) = symbols(k) Then
                    saymessage(Referencetotext(symbols(k)))
                    If foundcount < 3 Then
                        saymessage(", ")
                        foundcount += 1
                    End If
                End If
            Next
        Next
    End Sub
    Function Referencetotext(number As Integer)
        'converts the reference number passed to text
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

    Sub Initializereference(ByVal reference(,) As Integer)
        'initialises the reference columns
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
    Sub Showtable(ByVal reference(,) As Integer)
        'shows the reference columns --FOR TESTING PURPOSES--
        For i = 0 To 6
            For j = 0 To 5
                saymessage(reference(j, i) & " ")
                If reference(j, i) < 10 Then
                    saymessage(" ")
                End If
            Next
            saymessage("")
        Next
    End Sub
    Function Findcolumn(ByVal symbols() As Integer, ByVal table(,) As Integer)
        'finds the column which contains the solution
        Dim foundcount As Integer = 0
        For i = 0 To 5
            For j = 0 To 6
                For k = 0 To 3
                    If symbols(k) = table(i, j) Then
                        foundcount += 1
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
    Sub Inputsymbols(ByVal symbols() As String)
        'allows the user to input the symbols displayed to them
        For i = 0 To 3
            takespeech("Input Symbol " & i + 1 & ": ", symbols(i))
        Next
    End Sub
    Function Converttonums(ByVal symbols() As String, ByVal symbolnumber As Integer)
        'converts the users description to a reference number
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
    Structure Tnotes
        Dim note1, note2, note3, note4 As String
    End Structure
    Structure TChord
        Dim offsets As Toffsets
        Dim quality As String
        Dim note As Tnotes
    End Structure
    Sub Solvechordqs()
        ' solves the chord qualities module
        Dim notes() As String = {"A", "A#", "B", "C", "C#", "D", "D#", "E", "F", "F#", "G", "G#"}
        Dim chords(11) As TChord
        Initializeoffsets(chords)
        Dim chord1 As TChord
#Disable Warning BC42108 ' Variable is passed by reference before it has been assigned a value
        Inputstartchord(chord1)
#Enable Warning BC42108 ' Variable is passed by reference before it has been assigned a value
        Chordoffsets(chord1, notes)
        Dim startquality = Findbaseind(chord1.note.note1)
        Dim startandquality = Findquality(chord1, chords)
        If startandquality <> "CNF" Then
            Dim startnote As String = Mid(startandquality, 1, 1)
            Dim quality As String = Mid(startandquality, 2)
            Findanswerchord(startnote, quality, chord1, chords)
        End If
    End Sub
    Sub Findanswerchord(start As Integer, quality As String, chord As TChord, chords() As TChord)
        'finds the chord that the solution belongs to and outputs the notes in that chord
#Disable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
        Dim startnote As String = Setstartnote(start, chord)
#Enable Warning BC42030 ' Variable is passed by reference before it has been assigned a value
        Dim ansquality As String = Roottoquality(startnote)
        Dim ansroot As String = Qualitytoroot(quality)
        Dim qind As Integer = Findqualityindex(ansquality, chords)
        Dim endstartind As Integer = Findbaseind(ansroot)
        Findnotesolution(chords, endstartind, qind)
    End Sub
    Sub Findnotesolution(ByVal chords() As TChord, ByVal endstartind As Integer, ByVal qind As Integer)
        'outputs the notes the solution belongs to
        Dim ansnotes As Tnotes
        Dim notes() As String = {"A", "A#", "B", "C", "C#", "D", "D#", "E", "F", "F#", "G", "G#"}
        ansnotes.note1 = notes(endstartind)
        ansnotes.note2 = notes((endstartind + chords(qind).offsets.offset2) Mod 12)
        ansnotes.note3 = notes((endstartind + chords(qind).offsets.offset3) Mod 12)
        ansnotes.note4 = notes((endstartind + chords(qind).offsets.offset4) Mod 12)
        saymessage("Enter: " & ansnotes.note1 & ", " & ansnotes.note2 & ", " & ansnotes.note3 & " and " & ansnotes.note4)
    End Sub
    Function Setstartnote(ByVal start As Integer, chord As TChord)
        'sets the start note of the chord for checking
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
    Function Findqualityindex(ByVal quality As String, chords() As TChord)
        'finds the quality index of the chord
        For i = 0 To chords.Length - 1
            If chords(i).quality = quality Then
                Return i
            End If
        Next
        Return "Error"
    End Function
    Sub Displayoffsets(ByVal chord As TChord)
        'displays the offsets currently being used --FOR TESTING PURPOSES--
        saymessage(chord.offsets.offset1 & ", " & chord.offsets.offset2 & ", " & chord.offsets.offset3 & ", " & chord.offsets.offset4)
    End Sub
    Sub Chordoffsets(ByRef chord As TChord, ByVal notes() As String)
        'finds the offsets of each note
        Dim foundnotes As Integer = 0
        Dim offset As Integer = 0
        Dim startpoint = Findstartchord(chord.note.note1, notes)
        chord.offsets.offset1 = 0
        If CStr(startpoint) = "error" Then

            Do Until CStr(startpoint) <> "Error"
                saymessage("Input Error")
                Inputstartchord(chord)
                Findstartchord(chord.note.note1, notes)
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
                            foundnotes += 1
                        End If
                    Case 1
                        If chord.note.note3 = notes(checker) Then
                            chord.offsets.offset3 = offset
                            offset = 0
                            foundnotes += 1
                        End If
                    Case 2
                        If chord.note.note4 = notes(checker) Then
                            chord.offsets.offset4 = offset
                            chord.offsets.offset5 = 12 - (chord.offsets.offset2 + chord.offsets.offset3 + chord.offsets.offset4)
                            offset = 0
                            foundnotes += 1
                        End If
                End Select
                offset += 1
            Next
        Loop Until foundnotes = 3

    End Sub
    Function Findstartchord(ByVal note As String, ByVal notes() As String)
        'finds the start note of the chord
        For i = 0 To 11
            If note = notes(i) Then
                Return i
            End If
        Next
        Return "Error"
    End Function
    Function Findbaseind(ByVal startnote As String)
        'given a rot, returns its index
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
    Sub Inputstartchord(ByRef chord As TChord)
        'allows the user to enter the notes of the starting chord
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
    Function Findquality(ByRef chord As TChord, ByVal chords() As TChord)
        'finds the quality of the starting chord
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
                Rotateoffsets(chord)
            Next
        Next
        saymessage("chord not found")
        Return "CNF"
    End Function
    Sub Rotateoffsets(ByRef chord As TChord)
        'cycles the offsets
        Dim temp As Integer
        temp = chord.offsets.offset2
        chord.offsets.offset2 = chord.offsets.offset3
        chord.offsets.offset3 = chord.offsets.offset4
        chord.offsets.offset4 = chord.offsets.offset5
        chord.offsets.offset5 = temp
    End Sub
    Sub Initializeoffsets(ByVal chords() As TChord)
        'initialieses the offsets for each quality in the check table
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
    Function Roottoquality(ByVal startnote As String)
        'converts the passed root to the answer quality
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
    Function Qualitytoroot(ByRef quality As String)
        'converts the passed quality to the answer root
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
    Sub Solveskewed()
        'solves skewed slots module
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
        ElseIf Isprime(primes, slot(0)) Then
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
            ElseIf Isfibonacci(slot(1), fibonacci) Then
                slot(1) += Nextfib(slot(1), fibonacci)
            ElseIf slot(1) >= 7 Then
                slot(1) += 4
            Else slot(1) *= 3
            End If
        End If

        '3rd Slot 
        If serial Then
            slot(2) += Largestserial(serial)
        ElseIf Sameasorig(originalslots, slot(2)) Then
        ElseIf slot(2) >= 5 Then
            slot(2) += Addbinary(originalslots(2))
        Else slot(2) += 1
        End If
        For i = 0 To 2
            slot(i) = slot(i) Mod 10
        Next
        For i = 0 To 2
            saymessage(slot(i))
        Next
    End Sub
    Function Isfibonacci(ByVal number As Integer, ByVal fibonacci() As Integer)
        'checks if passed number is a memeber of the fibonacci sequence
        For i = 0 To fibonacci.Length
            If number = fibonacci(i) Then
                Return True
            End If
        Next
        Return False
    End Function
    Function Nextfib(ByVal number As Integer, ByVal fibonacci() As Integer)
        'returns the enxt number in the fibonacci sequence after the number passed
        For i = 0 To fibonacci.Length
            If number = fibonacci(i) Then
                Return fibonacci(i + 1)
            End If
        Next
        Return "ERROR"
    End Function
    Function Largestserial(serial)
        'finds the largest integer in the serial number
        Dim serialtemp, largest As Integer
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
    Function Sameasorig(ByVal slot() As Integer, ByVal current As Integer)
        'checks if a slot has remained unchanged
        For i = 0 To 2
            If slot(i) = current Then
                Return True
            End If
        Next
        Return False
    End Function
    Function Addbinary(orig)
        'converts the passed number to binary and sums the sigits of this binary number
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
    Function Isprime(primes() As Integer, number As Integer)
        'checks if the passed number is prime
        For i = 0 To primes.Length - 1
            If primes(i) = number Then
                Return True
            End If
        Next
        Return False
    End Function
    Sub Convertledstopropperenglish(ByRef leds(,) As String)
        'converts americanisms to proper english
        For i = 0 To 1
            For j = 0 To 3
                If leds(i, j) = "gray" Then
                    leds(i, j) = "grey"
                End If
            Next
        Next
    End Sub
    Sub SolveColourMath()
        'solves colour math module
        Dim leftnum, rightnum As Integer
        Dim numsolution As Integer
        Dim leds(1, 3) As String
        Dim centre As Toperator
        Dim coloursolution(3) As String
        InputLed("left", leds)
        Convertledstopropperenglish(leds)
        leftnum = Converttonumber(leds, "left")
        Enteroperatordata(centre)
        If centre.colour = "green" Then
            InputLed("right", leds)
            rightnum = Converttonumber(leds, "right")
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
        numsolution = Findnumsolution(leftnum, rightnum, centre)
        Convertbacktocolour(numsolution, coloursolution)
        saymessage(coloursolution(0))
        For i = 1 To 3
            saymessage(", " & coloursolution(i))
        Next
    End Sub
    Sub Convertbacktocolour(ByRef numsolution As Integer, ByRef coloursolution() As String)
        'comvers the numbers of the answer back into the colour solution
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
    Function Findnumsolution(ByVal num1 As Integer, ByVal num2 As Integer, ByVal op As Toperator)
        'finds the number solution to the colour math module
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
        saymessage("This Message Should not appear. If it does, Contact Chris ASAP")
        Return 0
    End Function
    Function Converttonumber(ByVal leds(,) As String, ByVal side As String)
        'converts the entered colours into integers for calculation
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
        Dim validcolours = {"blue", "green", "purple", "yellow", "white", "magenta", "red", "orange", "gray", "grey", "black"}
        'collects colours of leds on 1 side
        Dim sidenum As Integer = side.Length - 4
        'side 0 = left
        'side 1 = right
        For i = 0 To 3
            Dim validinput = False
            Do

                Dim temp As String
                LCase(takespeech("Enter LED Colour for LED " & i + 1 & " on the " & side & ": ", temp))
                If Checkcolour(temp, validcolours) Then
                    leds(sidenum, i) = temp
                    validinput = True
                Else
                    saymessage("Invalid Colour, please re-enter.")
                End If
            Loop Until validinput = True
        Next
    End Sub
    Function Checkcolour(ByVal colour As String, ByVal valid() As String)
        'ensures colours are valid
        For i = 0 To valid.Length - 1
            If colour = valid(i) Then
                Return True
            End If
        Next
        Return False
        Return "Error"
    End Function
    Sub Enteroperatordata(ByRef op As Toperator)
        'colelcts the needed details on the mmiddle operator
        Enteroperatorcolour(op.colour)
        Enteroperatortype(op.action)
    End Sub
    Sub Enteroperatorcolour(ByRef colourholder As String)
        'allows the user to enter the operator colour
        Dim validcolour As Boolean
        Dim temp As String
        Do
            saymessage("Enter colour of the operator in the centre: ")
            LCase(takespeech("Enter colour of the operator in the centre: ", temp))
            Select Case temp
                Case "red"
                    colourholder = "red"
                    validcolour = True
                Case "green"
                    colourholder = "green"
                    validcolour = True
                Case Else
                    saymessage("Invalid Colour please re-enter.")
                    validcolour = False
            End Select
        Loop Until validcolour = True
    End Sub
    Sub Enteroperatortype(ByRef actionholder As String)
        'allows the user to the enter the character displayed on the operator
        Dim validoperation As Boolean
        Dim temp As String
        Do
            UCase(takespeech("Enter letter in the centre: ", temp))
            Select Case temp
                Case "A", "S", "M", "D"
                    actionholder = temp
                    validoperation = True
                Case Else
                    saymessage("Invalid Letter please re-enter.")
                    validoperation = False
            End Select
        Loop Until validoperation = True
    End Sub
    Sub Solvesimonsays()
        'solves simon says module
        Dim sscomplete As Boolean = False
        Dim colourstring As String = ""
        Do
            Dim flash As String = Inputnewflashcolour()
            Select Case LCase(flash)
                Case "finish"
                    sscomplete = True
                Case Else
                    colourstring &= Converttooutputcolour(flash) & " "
                    saymessage(colourstring)
            End Select
        Loop Until sscomplete = True

    End Sub
    Function Converttooutputcolour(newflash)
        'converts the input to the correct output colour
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
        Return "this message should not appear if it does, contact chris ASAP"
    End Function
    Function Inputnewflashcolour()
        'allows the user to enter the most recent flash colour
        Dim temp As String
        Dim valid As Boolean = False
        Do
            takespeech("Enter Last Flash in Sequence, Say Finish to End: ", temp)
            Select Case LCase(temp)
                Case "red", "blue", "yellow", "green", "finish"
                    valid = True
                Case Else
                    saymessage("Invalid colour, please re-enter.")
            End Select
        Loop Until valid = True
        Return temp
    End Function
    Sub Inputletters(ByRef chars(,) As Char, ByRef extendedpassword As Integer)
        'enter letters for passwords and extended password
        For i = 0 To (4 + extendedpassword)
            saymessage("Position" & i)
            For j = 0 To 5
                takespeech("", chars(i, j))
                chars(i, j) = LCase(chars(i, j))
                saymessage("Next")
            Next
        Next
    End Sub
    Function Inputwire(ByVal pos As Integer)
        'stores wire for wire sequence
        Dim temp As String
        Do
            LCase(takespeech("Enter wire connected to " & pos & " (say pass if no wire): ", temp))
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
        Dim validcolours = {"red", "blue", "black", "yellow", "white"}
        'stores the colour of each wire
        For i = 0 To wirecount - 1
            takespeech("Colour of wire number " & i + 1 & ": ", wires(i))
            wires(i) = LCase(wires(i))
            If Not Checkcolour(wires(i), validcolours) Then
                saymessage("Invalid Input ")
                i -= 1
            End If
        Next
    End Sub
End Module