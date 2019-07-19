Module Module1
    'Symbol Key
    'lollipop = 1
    'Equation Triangle = 2 *
    'lambda = 3
    'lightning = 4 *
    'rocket = 5
    'fancy H = 6
    'backwards C = 7
    'backwards E with umlaut = 8
    'loop = 9
    'unfilled star = 10
    'upside down question mark = 11
    'copyright = 12 *
    'Butt = 13 *
    'XI = 14
    'broken 3 = 15 *
    'flat 6 = 16
    'paragraph = 17
    'b bar = 18
    'smiley = 19
    'trident = 20
    'forward C = 21 *
    'snake = 22 *
    'filled star = 23 *
    'jigsaw = 24 *
    'ae = 25 *
    'Backwards N = 26  *
    'Omega = 27
    Sub Main()

        Dim symbols(3) As String
        Dim convertedsymbols(3) As Integer
        Dim reference(5, 6) As Integer
        initializereference(reference)
        showtable(reference)
        Console.ReadLine()
        inputsymbols(symbols)
        For i = 0 To converttonums(symbols, i)
        Next
        Dim column = findcolumn(symbols, reference)
    End Sub
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
            Next
            Console.WriteLine()
        Next
    End Sub
    Function findcolumn(ByVal symbols() As String, ByVal table(,) As Integer)
        Dim column = ""
        Return column
    End Function
    Sub inputsymbols(ByVal symbols() As String)
        For i = 0 To 3
            Console.Write("Input Symbol " & i + 1 & ": ")
            symbols(i) = LCase(Console.ReadLine())
        Next
    End Sub
    Function converttonums(ByVal symbols() As String, ByVal symbolnumber As Integer)
        Dim newsymbol As Integer
        Select Case symbols(symbolnumber)
            Case "lollipop"
                Return 1
        End Select
    End Function
End Module
