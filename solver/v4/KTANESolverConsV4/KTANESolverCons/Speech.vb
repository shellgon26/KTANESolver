Imports System.Speech.Recognition
Module Speech
    Public Const TestMode = 1
    Public Sub Main()
        If TestMode = 1 Then
            RunProgram()
        Else
            'initialises the speech recogniser and runs the program within it
            ' Create an in-process speech recognizer for the en-US locale.  
            Using recognizer As New SpeechRecognitionEngine(
      New Globalization.CultureInfo("en-GB"))

                ' Create and load a dictation grammar.  
                recognizer.LoadGrammar(New DictationGrammar())

                ' Add a handler for the speech recognized event.
                AddHandler recognizer.SpeechRecognized,
                    AddressOf recognizer_SpeechRecognized

                ' Configure input to the speech recognizer.
                Try
                    recognizer.SetInputToDefaultAudioDevice()
                Catch
                    saymessage("No Microphone Detected, closing as a result.")
                    Environment.Exit(0)
                End Try
                ' Start asynchronous, continuous speech recognition.  
                recognizer.RecognizeAsync(RecognizeMode.Multiple)

                'run the program
                RunProgram()
            End Using
        End If
    End Sub

    ' Handle the SpeechRecognized event.  
    Sub recognizer_SpeechRecognized(sender As Object, e As SpeechRecognizedEventArgs)
        'assigns the detected speech for use
        speechdetect = True
        mostrecentspeech = e.Result.Text
    End Sub
    Function takespeech(ByVal message As String, ByRef asignee As Object, Optional ByVal numintended As Boolean = False)
        'takes a message passed to the function, outputs the message, and awaits speech input, 
        'once Speech Is detected tests If the detecting speech Is a valid input And If it Is, asigns it to the passed variable

        Dim temp As String
        saymessage(message)
        Do
            If TestMode = 1 Then
                temp = Console.ReadLine
            Else
                resetspeech()
                Do
                Loop Until speechdetect = True
                Threading.Thread.Sleep(50)
                temp = mostrecentspeech
                checkdigitstring()
                temp = mostrecentspeech
            End If
            Console.WriteLine(temp)
            If temp <> "" Then
                If checkint(temp) = True Or numintended = False Then
                    asignee = temp
                    Return True
                Else
                    saymessage("Invalid Input")
                    saymessage(message)
                End If
            End If
        Loop
    End Function
    Sub resetspeech()
        'resets the speech holder for another detection
        speechdetect = False
        mostrecentspeech = ""
    End Sub
    Sub checkdigitstring()
        'sets all sections of the string that match the word form of a single digit integer to said single digit integer
        Dim testspeech As String = " " & mostrecentspeech & " "
        If testspeech.Length < 5 Then
            For i = 3 To testspeech.Length
                For j = 0 To testspeech.Length - (i)
                    Select Case LCase(Mid(testspeech, j + 1, i))
                        Case "one"
                            testspeech = Left(testspeech, j - 1) & 1 & Right(testspeech, testspeech.Length - (i + j))
                            isallnum()
                        Case "two"
                            testspeech = Left(testspeech, j - 1) & 2 & Right(testspeech, testspeech.Length - (i + j))
                            isallnum()
                        Case "three"
                            testspeech = Left(testspeech, j - 1) & 3 & Right(testspeech, testspeech.Length - (i + j))
                            isallnum()
                        Case "four"
                            testspeech = Left(testspeech, j - 1) & 4 & Right(testspeech, testspeech.Length - (i + j))
                            isallnum()
                        Case "five"
                            testspeech = Left(testspeech, j - 1) & 5 & Right(testspeech, testspeech.Length - (i + j))
                            isallnum()
                        Case "six"
                            testspeech = Left(testspeech, j - 1) & 6 & Right(testspeech, testspeech.Length - (i + j))
                            isallnum()
                        Case "seven"
                            testspeech = Left(testspeech, j - 1) & 7 & Right(testspeech, testspeech.Length - (i + j))
                            isallnum()
                        Case "eight"
                            testspeech = Left(testspeech, j - 1) & 8 & Right(testspeech, testspeech.Length - (i + j))
                            isallnum()
                        Case "nine"
                            testspeech = Left(testspeech, j) & 9 & Right(testspeech, testspeech.Length - (i + j))
                            isallnum()
                        Case "zero"
                            testspeech = Left(testspeech, j - 1) & 0 & Right(testspeech, testspeech.Length - (i + j))
                            isallnum()
                    End Select
                Next
            Next
        Else
            For i = 3 To 5
                For j = 0 To testspeech.Length - (i)
                    Select Case LCase(Mid(testspeech, j + 1, i))
                        Case "one"
                            mostrecentspeech = Left(testspeech, j - 1) & 1 & Right(testspeech, testspeech.Length - (i + j))
                            isallnum()
                        Case "two"
                            mostrecentspeech = Left(testspeech, j - 1) & 2 & Right(testspeech, testspeech.Length - (i + j))
                            isallnum()
                        Case "three"
                            mostrecentspeech = Left(testspeech, j - 1) & 3 & Right(testspeech, testspeech.Length - (i + j))
                            isallnum()
                        Case "four"
                            mostrecentspeech = Left(testspeech, j - 1) & 4 & Right(testspeech, testspeech.Length - (i + j))
                            isallnum()
                        Case "five"
                            mostrecentspeech = Left(testspeech, j - 1) & 5 & Right(testspeech, testspeech.Length - (i + j))
                            isallnum()
                        Case "six"
                            mostrecentspeech = Left(testspeech, j - 1) & 6 & Right(testspeech, testspeech.Length - (i + j))
                            isallnum()
                        Case "seven"
                            mostrecentspeech = Left(testspeech, j - 1) & 7 & Right(testspeech, testspeech.Length - (i + j))
                            isallnum()
                        Case "eight"
                            mostrecentspeech = Left(testspeech, j - 1) & 8 & Right(testspeech, testspeech.Length - (i + j))
                            isallnum()
                        Case "nine"
                            mostrecentspeech = Left(testspeech, j - 1) & 9 & Right(testspeech, testspeech.Length - (i + j))
                            isallnum()
                        Case "zero"
                            mostrecentspeech = Left(testspeech, j - 1) & 0 & Right(testspeech, testspeech.Length - (i + j))
                            isallnum()
                    End Select
                Next
            Next
        End If
    End Sub
    Sub isallnum()
        'tests if the most recent set of speech only consists of numbers
        Dim temp As String = mostrecentspeech
        Dim allnum As Boolean = True
        Dim asciicharcheck As Integer
        For i = 0 To temp.Length - 1
            If i < temp.Length Then
                asciicharcheck = Asc(LCase(temp(i))) - 97
                Select Case asciicharcheck
                    Case 0 To 25
                        allnum = False
                    Case -65
                        If i - 1 > 0 And i + 1 < temp.Length Then
                            temp = Mid(temp, 1, i - 1) & Mid(temp, i + 1, temp.Length - i)
                        ElseIf i = 0 Then
                            temp = Mid(temp, i + 1, temp.Length - i)
                        ElseIf i + 1 >= temp.Length Then
                            temp = Mid(temp, 1, i)
                        Else

                        End If
                End Select
            End If
        Next
        If allnum = True Then
            mostrecentspeech = CInt(temp)
        End If
    End Sub
    Function removespaces(ByRef phrase As String)
        'removes all spaces from the string passed to it
        For i = 0 To phrase.Length - 1
            If i < phrase.Length Then
                If Asc(phrase(i)) - 97 = -65 Then
                    If i - 1 > 0 And i + 1 < phrase.Length Then
                        phrase = Mid(phrase, 1, i - 1) & Mid(phrase, i + 1, phrase.Length - i)
                    ElseIf i = 0 Then
                        phrase = Mid(phrase, i + 1, phrase.Length - i)
                    ElseIf i + 1 >= phrase.Length Then
                        phrase = Mid(phrase, 1, i)
                    Else

                    End If
                End If
            End If
        Next
        Return phrase
    End Function
    Sub saymessage(ByVal message As String)
        If TestMode = 1 Then
            If message(message.Length - 2) <> ":" Then
                Console.WriteLine(message)
            Else
                Console.Write(message)
            End If
        Else
            voice.Speak(message)
        End If
    End Sub
End Module
