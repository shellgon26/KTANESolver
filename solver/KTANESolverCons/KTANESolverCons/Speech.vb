Module Speech
    Public WithEvents recognizer As New System.Speech.Recognition.SpeechRecognitionEngine
    Dim gram As New System.Speech.Recognition.DictationGrammar()
    Dim cmd As String
    Public Sub Load()
        recognizer.SetInputToDefaultAudioDevice()
        recognizer.RecognizeAsync()
    End Sub
    Private Sub GotSpeech(ByVal phrase As System.Speech.Recognition.SpeechRecognizedEventArgs)
        cmd = phrase.Result.Text
        If cmd.IndexOf("Run") <> 0 Or cmd.IndexOf("run") <> 0 Then
            If cmd.Split(" ")(1) = "Notepad" Or cmd.Split(" ")(1) = "notepad" Then

                voice.Speak("Running Notepad.")

                Shell("notepad.exe", AppWinStyle.NormalFocus, False)

            End If
        End If
    End Sub
End Module
