' NOTE: Must target .NET Framework, 3.0 or later (not .NET Core!)
Imports System.Console

' Make reference to System.Speech (System.Speech.dll)
' https://docs.microsoft.com/en-us/dotnet/api/system.speech.recognition.speechrecognitionengine
Imports System.Speech.Recognition

Module Module1
    Sub Main()

        ' Create an in-process speech recognizer for the en-US locale.  
        Using recognizer As New SpeechRecognitionEngine(
          New Globalization.CultureInfo("en-US"))

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
            Do
                ReadLine()
            Loop

        End Using

    End Sub

    ' Handle the SpeechRecognized event.  
    Sub recognizer_SpeechRecognized(sender As Object, e As SpeechRecognizedEventArgs)
        WriteLine("Recognized text: " + e.Result.Text)
    End Sub

End Module