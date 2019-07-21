Module Module1

    Sub Main()
        Dim words() As String = {"about", "after", "again", "below", "could", "every", "first", "found", "great", "house", "large", "learn", "never", "other", "place", "plant", "point", "right", "small", "sound", "spell", "still", "study", "their", "there", "these", "thing", "think", "three", "water", "where", "which", "world", "would", "write"}
        Dim ewords() As String = {"adjust", "anchor", "bowtie", "button", "cipher", "corner", "dampen", "demote", "enlist", "evolve", "forget", "finish", "geyser", "global", "hammer", "helium", "ignite", "indigo", "jigsaw", "juliet", "karate", "lambda", "listen", "matter", "memory", "nebula", "nickel", "overdo", "oxygen", "peanut", "photon", "quartz", "quebec", "resist", "riddle", "sierra", "strike", "teapot", "twenty", "untold", "ultima", "victor", "violet", "wither", "wrench", "xenons", "xylose", "yellow", "yogurt", "zenith", "zodiac"}
        Dim epassword As Integer
        Dim response As String
        Console.Write("Extended?: ")
        response = Console.ReadLine
        If LCase(response(0)) = "y" Then
            epassword = 1
        Else
            epassword = 0
        End If
        Dim chars(4 + epassword, 5) As Char
        For i = 0 To 4 + epassword
            For j = 0 To 5
                Console.Write("Enter character " & j + 1 & " in position " & i + 1 & ": ")
                chars(i, j) = LCase(Console.ReadLine)
            Next
        Next
        For i = 0 To 4 + epassword
            For j = 0 To 5
                Console.Write(chars(i, j) & ", ")
            Next
            Console.WriteLine()
        Next
        Dim answer As String
        If epassword = 1 Then
            answer = shortenlist(ewords, chars, 0, epassword)
        Else
            answer = shortenlist(words, chars, 0, epassword)
        End If
        Console.WriteLine(answer)
        Console.ReadLine()

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
End Module
