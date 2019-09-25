
Module ModuleStartup
    Sub CheckEdgework()
        If CheckSerial() = True Then
            If checkbats() = True Then
                If (indicatorcount() + plates + holders <> widgets) Then

                End If
            End If
        End If
    End Sub
    Sub initializemodulecount()
        Dim temp
        voice.Speak("Number of modules: ")
        temp = Console.ReadLine()
        If checkint(temp) = True Then
            modulecount = temp
        End If
    End Sub
    Function CheckSerial()
        If serial.Length <> 6 Then
            Return False
        Else
            Return True
        End If

    End Function
    Sub inpWidgets()
        Dim temp As String = "no"
        While checkint(temp) = False
            voice.Speak("Widgets: ")
            temp = Console.ReadLine()
        End While
        widgets = temp
    End Sub
    Function checkbats()
        If batteries > holders * 2 Or holders > batteries Then
            Return False
        Else
            Return True
        End If
    End Function
    Sub Edgework()
        addbats()
        While checkbats() = False
            voice.Speak("Inputted batteries Impossible, please Re-enter")
            addbats()
        End While
        calculatetypes()
        addindicators()
        addports()
        inputserial()
    End Sub
    Sub inputserial()
        Do
            voice.Speak("Enter Serial Number: ")
            serial = LCase(Console.ReadLine)
        Loop Until CheckSerial() = True
    End Sub
    Function checkint(input)
        Try
            input = CInt(input)
        Catch ex As Exception
            Return False
        End Try
        Return True
    End Function
    Sub calculatetypes()
        aabats = (batteries - holders) * 2
        Dbats = batteries - aabats
        If (Dbats + aabats <> batteries) Or ((Dbats + aabats / 2) <> holders) Then
            voice.Speak("Error")
        End If
    End Sub
    Sub addindicators()
        Dim validind As Boolean
        Dim index As Integer
        Dim indname As String = ""
        Dim litcheck As String
        Do Until indname = "X"
            voice.Speak("Input Indicator Name (Input X to end indicator input): ")
            indname = UCase(Console.ReadLine())
            Select Case indname
                Case "IND"
                    index = 0
                    validind = True
                Case "FRK"
                    index = 1
                    validind = True
                Case "FRQ"
                    index = 2
                    validind = True
                Case "SND"
                    index = 3
                    validind = True
                Case "SLR"
                    index = 4
                    validind = True
                Case "CAR"
                    index = 5
                    validind = True
                Case "SIG"
                    index = 6
                    validind = True
                Case "MSA"
                    index = 7
                    validind = True
                Case "NSA"
                    index = 8
                    validind = True
                Case "TRN"
                    index = 9
                    validind = True
                Case "BOB"
                    index = 10
                    validind = True
                Case "X"
                Case Else
                    voice.Speak("Not a valid indicator")
            End Select
            If validind = True Then
                inds(index).present = True
                voice.Speak("Is the Indicator Lit?: ")
                litcheck = UCase(Console.ReadLine())
                If (litcheck = "Y") Or (litcheck = "YES") Or (litcheck = "TRUE") Then
                    inds(index).lit = True
                End If
            End If
        Loop
    End Sub
    Function indicatorcount()
        Dim count As Integer
        For i = 0 To 10
            If inds(i).present = True Then
                count = count + 1
            End If
        Next
        Return count
    End Function
    Sub initializeports()
        ports(0).name = "Serial"
        ports(1).name = "Parallel"
        ports(2).name = "PS2"
        ports(3).name = "RJ-45"
        ports(4).name = "Stereo RCA"
        ports(5).name = "DVI-D"
        ports(0).count = 0
        ports(1).count = 0
        ports(2).count = 0
        ports(3).count = 0
        ports(4).count = 0
        ports(5).count = 0
    End Sub
    Sub initializeinds()
        inds(0).text = "IND"
        inds(1).text = "FRK"
        inds(2).text = "FRQ"
        inds(3).text = "SND"
        inds(4).text = "SLR"
        inds(5).text = "CAR"
        inds(6).text = "SIG"
        inds(7).text = "MSA"
        inds(8).text = "NSA"
        inds(9).text = "TRN"
        inds(10).text = "BOB"
    End Sub
    Sub addports()
        Dim portname As String = ""
        Dim index, count As Integer
        Dim validport As Boolean = False
        While portname <> "x"
            voice.Speak("Enter Port Name, (Enter X if no more remain): ")
            portname = LCase(Console.ReadLine)
            Select Case portname
                Case "serial"
                    index = 0
                    validport = True
                Case "parallel"
                    index = 1
                    validport = True
                Case "ps2"
                    index = 2
                    validport = True
                Case "rj45"
                    index = 3
                    validport = True
                Case "rj-45"
                    index = 3
                    validport = True
                Case "rj"
                    index = 3
                    validport = True
                Case "stereo rca"
                    index = 4
                    validport = True
                Case "rca"
                    index = 4
                    validport = True
                Case "dvid"
                    index = 5
                    validport = True
                Case "dvi-d"
                    index = 5
                    validport = True
                Case "x"
                Case Else
                    voice.Speak("Invalid port")
            End Select
            If validport = True Then
                voice.Speak("Count: ")
                count = Console.ReadLine
                ports(index).count = ports(index).count + count
            End If
        End While
    End Sub
    Sub addbats()
        Dim valid1, valid2 As Boolean
        valid1 = False
        valid2 = False
        Dim temp1, temp2 As String
        Do
            voice.Speak("Batteries: ")
            temp1 = Console.ReadLine
            valid1 = checkint(temp1)
            voice.Speak("Holders: ")
            temp2 = Console.ReadLine
            valid2 = checkint(temp2)
        Loop Until valid1 = True And valid2 = True
        batteries = temp1
        holders = temp2
    End Sub
End Module
