Module Module1
    Dim litinds, unlitinds As Integer
    Dim rca, ps2, parallel, serial As Boolean
    Dim unlitbob As Boolean = False
    Dim batteries As Integer
    Dim serialnum As String
    Dim primes() As Integer = {2, 3, 5, 7, 11, 13, 17, 23, 29}
    Dim fibonacci() As Integer = {1, 1, 2, 3, 5, 8, 13, 21, 34, 55, 89, 144}
    Dim lastserial As Integer
    Sub Main()
        Dim slot(2) As Integer
        Dim originalslots(2) As Integer

        For i = 0 To 2
            slot(i) = Console.ReadLine
            originalslots(i) = slot(i)
        Next
        'All Slots - 011 -> 421
        For i = 0 To 2
            If slot(i) = 2 Then
                slot(i) = 5
            ElseIf slot(i) = 7 Then
                slot(i) = 0
            End If '011
            slot(i) = slot(i) + litinds - unlitinds
            If slot(i) Mod 3 = 0 Then
                slot(i) += 4 '411
            ElseIf slot(i) > 7 Then
                slot(i) *= 2
            ElseIf (slot(i)) < 3 And (slot(i) Mod 2 = 0) Then
                slot(i) /= 2
            ElseIf rca Or ps2 Then
            Else slot(i) += batteries '411
            End If
        Next
        '1St Slot 
        If slot(0) Mod 2 = 0 And slot(0) > 5 Then
            slot(0) /= 2
        ElseIf isprime(primes, slot(0)) Then
            slot(0) += lastserial
        ElseIf parallel Then
            slot(0) *= -1
        ElseIf originalslots(0) Mod 2 = 1 Then
        Else slot(0) -= 2 '211
        End If

        '2nd Slot 
        If Not unlitbob Then
            If slot(1) = 0 Then
                slot(1) += originalslots(0)
            ElseIf isfibonacci(slot(1)) Then
                slot(1) += nextfib(slot(1))
            ElseIf slot(1) >= 7 Then
                slot(1) += 4
            Else slot(1) *= 3
            End If
        End If

        '3rd Slot 
        If serial Then
            slot(2) += largestserial(serialnum)
        ElseIf sameasorig(originalslots, slot(2)) Then
        ElseIf slot(2) >= 5 Then
            slot(2) += addbinary(originalslots(2))
        Else slot(2) += 1
        End If
        For i = 0 To 2
            slot(i) = slot(i) Mod 10
        Next
        For i = 0 To 2
            Console.Write(slot(i))
        Next
        Console.ReadLine()
    End Sub
    Function prime(number)
        For i = 0 To primes.Length
            If number = primes(i) Then
                Return True
            End If
        Next
        Return False
    End Function
    Function isfibonacci(number)
        For i = 0 To primes.Length
            If number = fibonacci(i) Then
                Return True
            End If
        Next
        Return False
    End Function
    Function nextfib(number)
        For i = 0 To primes.Length
            If number = fibonacci(i) Then
                Return fibonacci(i + 1)
            End If
        Next
        Return "ERROR"
    End Function
    Function largestserial(serial)
        Dim serialtemp, largest As Integer
        For i = 0 To 5
            Try
                serialtemp = CInt(Mid(serial, i, 1))
            Catch ex As Exception

            End Try
            If serialtemp > largest Then
                largest = serialtemp
            End If
        Next
        Return largest
    End Function
    Function sameasorig(ByVal slot() As Integer, ByVal current As Integer)
        For i = 0 To 2
            If slot(i) = current Then
                Return True
            End If
        Next
        Return False
    End Function
    Function addbinary(orig)
        Dim binary() As Integer
        Dim loopcount As Integer = 0
        Do
            ReDim Preserve binary(loopcount)
            binary(loopcount) = orig Mod 2
            orig -= binary(loopcount) * 2 ^ loopcount
            orig /= 2
            loopcount += 1
        Loop Until orig / 2 < 1
        ReDim Preserve binary(loopcount)
        binary(loopcount) = orig
        Dim total As Integer = 0
        For i = 0 To loopcount
            total += binary(i)
        Next
        Return total
    End Function
    Function isprime(primes() As Integer, number As Integer)
        For i = 0 To primes.Length - 1
            If primes(i) = number Then
                Return True
            End If
        Next
        Return False
    End Function
End Module
