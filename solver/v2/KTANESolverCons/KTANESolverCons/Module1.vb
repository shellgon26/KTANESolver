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
            solvedmodules = +1
            Select Case LCase(modulechoice)
                Case "button"
                    SolveButton()
                Case "the button"
                    SolveButton()
                Case "wires"
                    Solvewires()
                Case "wire sequence"
                    Solvewireseq()
                Case "complicated wires"
                    solvecompwires()
                Case "password"
                    Solvepassword()
                Case "combination Lock"
                    Solvecombinationlock()
                Case "maze"
                    Solvemaze()
                Case "keypad"
                    solvekeypad()
                Case "chord qualities"
                    solvechordqs()
                Case "skewed slots"
                    solveskewed()
                Case "simon says"
                    solvesimonsays()
                Case "colour math", "color math"
                    SolveColourMath()
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

    Function Checkforletterinserial(ByVal letter)
        For i = 0 To 5
            If Mid(serial, i, 1) = letter Then
                Return True
            End If
        Next
        Return False
    End Function
    Public Function Calclitinds()
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
        'checks if number is even
        If number Mod 2 = 0 Then
            Return True
        Else : Return False
        End If
    End Function
End Module
