Imports System.Speech.Synthesis
Module Module1
    Public speechdetect As Boolean = False
    Public mostrecentspeech As String
    Structure Tindicators
        Dim text As String
        Dim lit As Boolean
        Dim present As Boolean
    End Structure
    Structure TPort
        Dim name As String
        Dim count As Integer
    End Structure
    Public voice As New SpeechSynthesizer
    Public serial As String
    Public inds(11) As Tindicators
    Public modulecount, solvedmodules, batteries, widgets, plates, holders, aabats, Dbats As Integer
    Public ports(6) As TPort
    Sub Main()
        InitialiseSpeech()
        initializeports()
        initializeinds()
        initializemodulecount()
        Edgework()
        Console.Clear()
        SolveModule()
        Console.ReadLine()
    End Sub
    Sub SolveModule()
        Dim complete As Boolean
        Dim modulechoice As String
        Do
        voice.Speak("Enter Module to Solve: ")
            modulechoice = LCase(Console.ReadLine())
            solvedmodules = +1
            Select Case LCase(modulechoice)
                Case "button"
                    SolveButton()
                Case "the button"
                    SolveButton()
                Case "wires"
                    solvewires()
                Case "wire sequence"
                    solvewireseq()
                Case "complicated wires"
                    solvecompwires()
                Case "password"
                    Solvepassword()
                Case "combination Lock"
                    solvecombinationlock()
                Case "maze"
                    solvemaze()
                Case "keypad"
                    solvekeypad()
                Case "chord qualities"
                    solvechordqs()
                Case Else
                    voice.Speak("Unrecognised")
                    solvedmodules -= 1
            End Select
            If modulecount = solvedmodules Then
                complete = True
                voice.Speak("Bomb Complete!")
            End If
        Loop Until complete = True
    End Sub
    Sub inputletters(ByRef chars(,) As Char, ByRef extendedpassword As Integer)
        For i = 0 To (4 + extendedpassword)
            For j = 0 To 5
                Console.Write("Enter Char number " & j & " In position " & i & ": ")
                chars(i, j) = Console.ReadLine
            Next
        Next
    End Sub
    Function inputwire(ByVal pos As Integer)
        Dim temp As String
        Do
            voice.Speak("Enter wire connected to " & pos & " (press enter if no wire): ")
            temp = LCase(Console.ReadLine())
        Loop Until temp = "red" Or temp = "blue" Or temp = "black" Or temp = ""
        Return temp
    End Function
    Function checkeven(ByVal number As Integer)
        If number Mod 2 = 0 Then
            Return True
        Else : Return False
        End If
    End Function
    Function findlastofcolour(ByVal colour, ByVal wirecount, ByVal wires())
        For i = (wirecount - 1) To 0 Step -1
            If wires(i) = colour Or wires(i) = colour(0) Then
                Return i
            End If
        Next
    End Function
    Function searchforwire(ByVal colour, ByVal wirecount, ByVal wires())
        Dim count As Integer = 0
        For i = 0 To wirecount - 1
            If wires(i) = colour Or wires(i) = colour(0) Then
                count = count + 1
            End If
        Next
        Return count
    End Function
    Sub inputwires(ByRef wires() As String, ByRef wirecount As Integer)
        For i = 0 To wirecount - 1
            voice.Speak("Colour of wire number " & i + 1 & ": ")
            wires(i) = LCase(Console.ReadLine)
        Next
    End Sub
    Function Checkforletterinserial(ByVal letter)
        For i = 0 To 5
            If Mid(serial, i, 1) = letter Then
                Return True
            End If
        Next
        Return False
    End Function
    Sub resetspeech()
        speechdetect = False
    End Sub
End Module
