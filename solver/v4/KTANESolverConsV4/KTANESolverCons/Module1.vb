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
    Structure Toperator
        Dim colour As String
        Dim action As String
    End Structure
    Public voice As New SpeechSynthesizer
    Public serial As String
    Public inds(11) As Tindicators
    Public modulecount, solvedmodules, batteries, widgets, plates, holders, aabats, Dbats, serialvowels As Integer
    Public ports(6) As TPort
    Public serialletters(0) As Char
    Public serialnums(0) As Integer
    Sub RunProgram()
        ' runs the initialisation functions and then enters core of the program
        'voice.Rate = 10
        initializeports()
        initializeinds()
        initializemodulecount()
        Edgework()
        Console.Clear()
        SolveModule()
    End Sub
    Sub SolveModule()
        'finds the module to complete
        Dim complete As Boolean
        Dim modulechoice As String
        Do
            takespeech("Enter Module to Solve: ", modulechoice)
            solvedmodules += 1
            Select Case LCase(modulechoice)
                Case "button", "the button"
                    SolveButton()
                Case "wires"
                    Solvewires()
                Case "wire sequence"
                    Solvewireseq()
                Case "complicated wires"
                    solvecompwires()
                Case "password"
                    Solvepassword()
                Case "combination lock"
                    Solvecombinationlock()
                Case "maze"
                    Solvemaze()
                Case "keypad"
                    Solvekeypad()
                Case "chord qualities"
                    Solvechordqs()
                Case "skewed slots", "skewed"
                    Solveskewed()
                Case "simon says"
                    Solvesimonsays()
                Case "colour math", "color math"
                    SolveColourMath()
                Case Else
                    saymessage("Unrecognised")
                    solvedmodules -= 1
            End Select
            If modulecount = solvedmodules Then
                complete = True
                saymessage("Bomb Complete!")
            End If
        Loop Until complete = True
    End Sub

    Function Checkforletterinserial(ByVal letter)
        'checks if the passed letter is in the serian numbe entered
        For i = 0 To 5
            If Mid(serial, i, 1) = letter Then
                Return True
            End If
        Next
        Return False
    End Function
    Public Function Calclitinds()
        ' counts and returns the number of lit indicatrs
        Dim litinds As Integer = 0
        For i = 0 To inds.Length
            If inds(i).present = True Then
                If inds(i).lit = True Then
                    litinds += 1
                End If
            End If
        Next
        Return litinds
    End Function
    Public Function Calcunlitinds()
        'counts and returns the number of unlit indicators
        Dim unlitinds As Integer = 0
        For i = 0 To inds.Length
            If inds(i).present = True Then
                If inds(i).lit = False Then
                    unlitinds += 1
                End If
            End If
        Next
        Return unlitinds
    End Function
    Function Checkeven(ByVal number As Integer)
        'checks if number passed is even
        If number Mod 2 = 0 Then
            Return True
        Else : Return False
        End If
    End Function
End Module
