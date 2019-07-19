Module Module1

    Sub Main()
        solvemaze()
        Console.ReadLine()
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