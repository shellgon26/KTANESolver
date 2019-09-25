Imports System.Speech.Recognition
Module Speech

    Public Sub InitialiseSpeech()
        ' Create an in-process speech recognizer for the en-US locale.  
        Using recognizer As New SpeechRecognitionEngine(
      New Globalization.CultureInfo("en-GB"))

            ' Create and load a dictation grammar.  
            recognizer.LoadGrammar(New DictationGrammar())

            ' Add a handler for the speech recognized event.
            AddHandler recognizer.SpeechRecognized,
                AddressOf recognizer_SpeechRecognized

            ' Configure input to the speech recognizer.  
            recognizer.SetInputToDefaultAudioDevice()

            ' Start asynchronous, continuous speech recognition.  
            recognizer.RecognizeAsync(RecognizeMode.Multiple)

            ' Keep the console window open.  
        End Using

    End Sub

    ' Handle the SpeechRecognized event.  
    Sub recognizer_SpeechRecognized(sender As Object, e As SpeechRecognizedEventArgs)
        speechdetect = True
        mostrecentspeech = e.Result.Text
    End Sub

End Module
