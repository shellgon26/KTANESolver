
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
        If TestMode >= 1 Then
            modulecount = 6
        Else
            Do
            Loop Until takespeech("Number of Modules: ", modulecount, True) = True
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
            takespeech("Widgets: ", temp)
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
        If TestMode >= 1 Then
            batteries = 2
            holders = 2
            addindicators()
            addports()
            serial = "ABC123"
        Else
            addbats()
            While checkbats() = False
                saymessage("Inputted batteries Impossible, please Re-enter")
                addbats()
            End While
            calculatetypes()
            addindicators()
            addports()
            inputserial()
        End If
    End Sub
    Sub inputserial()
        Do
            takespeech("Enter Serial Number: ", serial)
            serial = removespaces(serial)
        Loop Until CheckSerial() = True
        letternuminserial(UCase(serial))
    End Sub
    Function checkint(ByRef input As String)
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
            saymessage("Error")
        End If
    End Sub
    Sub addindicators()
        Dim validind As Boolean
        Dim index As Integer
        Dim indname As String = ""
        Dim litcheck As String
        Do
            takespeech("Input Indicator Name (Input X to end indicator input): ", indname)
            validind = False
            Select Case UCase(indname)
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
                    saymessage("Not a valid indicator")
            End Select
            If validind = True Then
                inds(index).present = True
                takespeech("Is the Indicator Lit?: ", litcheck)
                If (UCase(litcheck) = "Y") Or (UCase(litcheck) = "YES") Or (UCase(litcheck) = "TRUE") Then
                    inds(index).lit = True
                End If
            End If
        Loop Until UCase(indname) = "X"
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
        While LCase(portname) <> "x"
            takespeech("Enter Port Name, (Enter X if no more remain): ", portname)
            validport = False
            Select Case LCase(portname)
                Case "serial"
                    index = 0
                    validport = True
                Case "parallel"
                    index = 1
                    validport = True
                Case "ps2"
                    index = 2
                    validport = True
                Case "rj-45", "rj", "rj45"
                    index = 3
                    validport = True
                Case "stereo rca", "rca"
                    index = 4
                    validport = True
                Case "dvid", "dvi-d"
                    index = 5
                    validport = True
                Case "x"
                Case Else
                    saymessage("Invalid port")
            End Select
            If validport = True Then
                takespeech("Count: ", count, True)
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
            valid1 = takespeech("Batteries: ", temp1, True)
            valid2 = takespeech("Holders: ", temp2, True)
        Loop Until valid1 = True And valid2 = True
        batteries = temp1
        holders = temp2
    End Sub
    Sub letternuminserial(serial As String)
        Dim isnum As Integer
        For i = 1 To serial.Length + 1
            Dim temp As String = Mid(serial, i, 1)
            Try
                temp = CInt(temp)
            Catch ex As Exception
                isnum = False
            End Try
            If isnum Then
                serialnums.Append(temp)
            Else
                serialletters.Append(temp)
                Select Case UCase(temp)
                    Case "A", "E", "I", "O", "U"
                        serialvowels += 1
                End Select
            End If
        Next
    End Sub
End Module
