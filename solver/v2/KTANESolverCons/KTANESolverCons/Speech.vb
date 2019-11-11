Imports System.Speech.Recognition
Module Speech

    Public Sub Main()
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
                voice.Speak("No Microphone Detected, closing as a result.")
                Environment.Exit(0)
            End Try
            ' Start asynchronous, continuous speech recognition.  
            recognizer.RecognizeAsync(RecognizeMode.Multiple)

            'run the program
            RunProgram()
        End Using

    End Sub

    ' Handle the SpeechRecognized event.  
    Sub recognizer_SpeechRecognized(sender As Object, e As SpeechRecognizedEventArgs)
        speechdetect = True
        mostrecentspeech = e.Result.Text
    End Sub
    Function takespeech(ByVal message As String, ByRef asignee As Object, Optional ByVal numintended As Boolean = False)
        Dim temp As String
        voice.Speak(message)
        Do
            resetspeech()
            Do
            Loop Until speechdetect = True
            checkdigitstring()
            temp = mostrecentspeech
            Console.WriteLine(temp)
            If temp <> "" Then
                If checkint(temp) = True Or numintended = False Then
                    voice.Speak("Valid")
                    asignee = mostrecentspeech
                    Return True
                Else
                    voice.Speak("Invalid Input")
                    voice.Speak(message)
                End If
            End If
        Loop
    End Function
    Sub resetspeech()
        speechdetect = False
        mostrecentspeech = ""
    End Sub
    Sub checkdigitstring()
        Dim testspeech As String = " " & mostrecentspeech & " "
        If mostrecentspeech.Length < 5 Then
            For i = 3 To mostrecentspeech.Length
                For j = 0 To mostrecentspeech.Length - (i)
                    Select Case LCase(Mid(mostrecentspeech, j + 1, i))
                        Case "one"
                            mostrecentspeech = Left(mostrecentspeech, j - 1) & 1 & Right(mostrecentspeech, mostrecentspeech.Length - (i + j))
                            isallnum()
                        Case "two"
                            mostrecentspeech = Left(mostrecentspeech, j - 1) & 2 & Right(mostrecentspeech, mostrecentspeech.Length - (i + j))
                            isallnum()
                        Case "three"
                            mostrecentspeech = Left(mostrecentspeech, j - 1) & 3 & Right(mostrecentspeech, mostrecentspeech.Length - (i + j))
                            isallnum()
                        Case "four"
                            mostrecentspeech = Left(mostrecentspeech, j - 1) & 4 & Right(mostrecentspeech, mostrecentspeech.Length - (i + j))
                            isallnum()
                        Case "five"
                            mostrecentspeech = Left(mostrecentspeech, j - 1) & 5 & Right(mostrecentspeech, mostrecentspeech.Length - (i + j))
                            isallnum()
                        Case "six"
                            mostrecentspeech = Left(mostrecentspeech, j - 1) & 6 & Right(mostrecentspeech, mostrecentspeech.Length - (i + j))
                            isallnum()
                        Case "seven"
                            mostrecentspeech = Left(mostrecentspeech, j - 1) & 7 & Right(mostrecentspeech, mostrecentspeech.Length - (i + j))
                            isallnum()
                        Case "eight"
                            mostrecentspeech = Left(mostrecentspeech, j - 1) & 8 & Right(mostrecentspeech, mostrecentspeech.Length - (i + j))
                            isallnum()
                        Case "nine"
                            mostrecentspeech = Left(mostrecentspeech, j) & 9 & Right(mostrecentspeech, mostrecentspeech.Length - (i + j))
                            isallnum()
                        Case "zero"
                            mostrecentspeech = Left(mostrecentspeech, j - 1) & 0 & Right(mostrecentspeech, mostrecentspeech.Length - (i + j))
                            isallnum()
                    End Select
                Next
            Next
        Else
            For i = 3 To 5
                For j = 0 To mostrecentspeech.Length - (i)
                    Select Case LCase(Mid(mostrecentspeech, j + 1, i))
                        Case "one"
                            mostrecentspeech = Left(mostrecentspeech, j - 1) & 1 & Right(mostrecentspeech, mostrecentspeech.Length - (i + j))
                            isallnum()
                        Case "two"
                            mostrecentspeech = Left(mostrecentspeech, j - 1) & 2 & Right(mostrecentspeech, mostrecentspeech.Length - (i + j))
                            isallnum()
                        Case "three"
                            mostrecentspeech = Left(mostrecentspeech, j - 1) & 3 & Right(mostrecentspeech, mostrecentspeech.Length - (i + j))
                            isallnum()
                        Case "four"
                            mostrecentspeech = Left(mostrecentspeech, j - 1) & 4 & Right(mostrecentspeech, mostrecentspeech.Length - (i + j))
                            isallnum()
                        Case "five"
                            mostrecentspeech = Left(mostrecentspeech, j - 1) & 5 & Right(mostrecentspeech, mostrecentspeech.Length - (i + j))
                            isallnum()
                        Case "six"
                            mostrecentspeech = Left(mostrecentspeech, j - 1) & 6 & Right(mostrecentspeech, mostrecentspeech.Length - (i + j))
                            isallnum()
                        Case "seven"
                            mostrecentspeech = Left(mostrecentspeech, j - 1) & 7 & Right(mostrecentspeech, mostrecentspeech.Length - (i + j))
                            isallnum()
                        Case "eight"
                            mostrecentspeech = Left(mostrecentspeech, j - 1) & 8 & Right(mostrecentspeech, mostrecentspeech.Length - (i + j))
                            isallnum()
                        Case "nine"
                            mostrecentspeech = Left(mostrecentspeech, j - 1) & 9 & Right(mostrecentspeech, mostrecentspeech.Length - (i + j))
                            isallnum()
                        Case "zero"
                            mostrecentspeech = Left(mostrecentspeech, j - 1) & 0 & Right(mostrecentspeech, mostrecentspeech.Length - (i + j))
                            isallnum()
                    End Select
                Next
            Next
        End If
    End Sub
    Sub isallnum()
        Dim temp As String = mostrecentspeech
        Dim allnum As Boolean = True
        For i = 0 To temp.Length - 1
            Select Case Asc(temp(i)) - 97
                Case Asc(temp(i)) > -1 And Asc(temp(i)) < 26
                    allnum = False
                Case -65
                    temp = Mid(temp, 0, i - 1) & Mid(temp, 0, i + 1)
            End Select
        Next
        If allnum = True Then
            mostrecentspeech = CInt(temp)
        End If
    End Sub
End Module
