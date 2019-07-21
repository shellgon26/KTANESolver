Imports System.Speech.Recognition
Imports System.Speech.Synthesis
Imports System.Threading
Imports System.Globalization
Module Module1

    Structure Tindicators
        Dim text As String
        Dim lit As Boolean
        Dim present As Boolean
    End Structure
    Structure TPort
        Dim name As String
        Dim count As Integer
    End Structure
    Dim recogniser As New SpeechRecognizer
    Dim grammerbuilderStartStop As New GrammarBuilder

    Public voice As New SpeechSynthesizer
    Public serial As String
    Public inds(11) As Tindicators
    Public solvedmodules, modulecount, batteries, widgets, plates, holders, aabats, Dbats As Integer
    Public ports(6) As TPort
    Public time As Integer
    Public litunlit(1) As Integer
    Sub Main()
        initializeports()
        initializeinds()
        inputstartingtime(time)
        Edgework()
        Console.Clear()
        SolveModule()
        Console.ReadLine()
    End Sub
    Function speechregogniser(words() As String)
        For i = 0 To words.Length
            grammerbuilderStartStop.Append(words(i))
        Next
        Dim grammerstartstop As New Grammar(grammerbuilderStartStop)
        grammerstartstop.Enabled = True
        recogniser.LoadGrammar(grammerstartstop)
    End Function
    Sub SolveModule()
        Dim complete As Boolean
        Dim modulechoice As String
        Do
            voice.Speak("Enter Module to Solve: ")
            modulechoice = LCase(Console.ReadLine())
            Select Case modulechoice
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
                    solvepassword()
                Case "extended password"
                    solveextendedpassword()
                Case "extpassword"
                    solveextendedpassword()
                Case "bitwise"
                    solvebitwise()
                Case "bitwise operators"
                    solvebitwise()
                Case "maze"
                    solvemaze()
                Case Else
                    voice.Speak("Unrecognised")
            End Select
            If solvedmodules = modulecount Then
                complete = True
            End If
        Loop Until complete = True
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
        Else Return False
        End If
    End Function
    Function findlastofcolour(colour, wirecount, wires())
        For i = (wirecount - 1) To 0 Step -1
            If wires(i) = colour Or wires(i) = colour(0) Then
                Return i
            End If
        Next
    End Function
    Function searchforwire(colour, wirecount, wires())
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

End Module
