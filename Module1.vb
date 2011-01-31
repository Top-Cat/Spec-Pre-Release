Module Module1

    Sub Main()
        Dim Choice As Integer
        Dim Finished As Boolean
        Finished = False ' Does nothing
        While Not Finished
            ShowMenu()
            GetResponse(Choice)
            Select Case Choice
                Case 1 : ConvertNumber(True)
                Case 2 : ConvertNumber(False)
                Case 3 : ConvertFile()
                Case 4 : DisplayFile()
                Case 5 : Finished = True
            End Select
            Console.ReadLine()
        End While
    End Sub ' of main

    Sub ShowMenu()
        Console.WriteLine("Please choose an option")
        Console.Writeline()
        Console.Writeline("1 - Convert a Hex numnber to Binary")
        Console.Writeline("2 - Convert a Binary numnber to Hex")
        Console.Writeline("3 - Convert a text file of Hex numbers")
        Console.WriteLine("4 - Display text file")
        Console.WriteLine("5 - Exit program")
        Console.Writeline()
    End Sub ' of ShowMenu

    Sub GetResponse(ByRef Response As Integer)
        Response = 0
        While Response > 4 Or Response < 1
            Console.Write("Enter option number: ")
            Response = Val(Console.ReadLine())
        End While
    End Sub ' of GetResponse

    Function Hex(ByVal Binary As String) As String
        Dim Result As String
        While Math.IEEERemainder(Len(Binary), 4) <> 0
            Binary = "0" & Binary
        End While
        For i = 1 To (Len(Binary) / 4)
            Select Case Mid(Binary, ((i - 1) * 4) + 1, 4)
                Case "0000" : Result &= "0" 'corrected bin equiv of hex digits
                Case "0001" : Result &= "1"
                Case "0010" : Result &= "2"
                Case "0011" : Result &= "3"
                Case "0100" : Result &= "4"
                Case "0101" : Result &= "5"
                Case "0110" : Result &= "6"
                Case "0111" : Result &= "7"
                Case "1000" : Result &= "8"
                Case "1001" : Result &= "9"
                Case "1010" : Result &= "A"
                Case "1011" : Result &= "B"
                Case "1100" : Result &= "C"
                Case "1101" : Result &= "D"
                Case "1110" : Result &= "E"
                Case "1111" : Result &= "F"
            End Select
        Next
        Return Result
    End Sub

    Function Binary(ByVal Hex As String) As String
        Dim Result, ThisHexDigit, BinaryEquivalent As String
        Dim HexDigit, NoOfHexDigits As Integer
        Dim ErrorFound As Boolean = False
        Result = ""
        BinaryEquivalent = "" 'added - but not neccessary
        NoOfHexDigits = Len(Hex)
        For HexDigit = 1 To NoOfHexDigits
            ThisHexDigit = Mid(Hex, HexDigit, 1)
            If InStr("0123456789ABCDEF", ThisHexDigit) <> 0 Then
                Select Case ThisHexDigit
                    Case "0" : BinaryEquivalent = "0000" 'corrected bin equiv of hex digits
                    Case "1" : BinaryEquivalent = "0001"
                    Case "2" : BinaryEquivalent = "0010"
                    Case "3" : BinaryEquivalent = "0011"
                    Case "4" : BinaryEquivalent = "0100"
                    Case "5" : BinaryEquivalent = "0101"
                    Case "6" : BinaryEquivalent = "0110"
                    Case "7" : BinaryEquivalent = "0111"
                    Case "8" : BinaryEquivalent = "1000"
                    Case "9" : BinaryEquivalent = "1001"
                    Case "A" : BinaryEquivalent = "1010"
                    Case "B" : BinaryEquivalent = "1011"
                    Case "C" : BinaryEquivalent = "1100"
                    Case "D" : BinaryEquivalent = "1101"
                    Case "E" : BinaryEquivalent = "1110"
                    Case "F" : BinaryEquivalent = "1111"
                End Select
            Else
                ErrorFound = True
                Exit For
            End If
            Result = Result + BinaryEquivalent 'Should use '&' but won't error. Due to hardcoded strings
        Next
        If ErrorFound Then
            Return "Invalid Input"
        Else
            Return Result
        End If
    End Function ' of Binary

    Sub ConvertNumber(ByVal HtoB as Boolean)
        Dim Hexadecimal , Converted, con As String
        con = "Binary"
        If HtoB Then
            con = "Hexadecimal"
        End If
        Console.Write("Enter a " + con + " number: ")
        Hexadecimal = Console.ReadLine
        If HtoB Then
            Converted = Binary(Hexadecimal)
        Else
            Converted = Hex(Hexadecimal)
        End If
        Console.WriteLine(Converted)
    End Sub ' of ConvertNumber

    Sub ConvertFile()
        Dim HexNumber, BinaryNumber As String
        FileOpen(1, "C:\Documents and Settings\VPCuser\My Documents\HexData.txt", OpenMode.Input)
        FileOpen(2, "C:\Documents and Settings\VPCuser\My Documents\BinaryData.txt", OpenMode.Output)
        Console.WriteLine()
        Do While Not EOF(1)
            HexNumber = LineInput(1)
            BinaryNumber = Binary(HexNumber)
            Console.WriteLine(BinaryNumber)
            FilePut(2, BinaryNumber)
        Loop
        FileClose(1)
        FileClose(2)
        Console.ReadLine()
    End Sub ' of ConvertFile

    Sub DisplayFile()
        Dim NextNumber As String
        FileOpen(1, "C:\BinaryData.txt", OpenMode.Input)
        Console.WriteLine()
        Do While Not EOF(1)
            NextNumber = LineInput(1)
            Console.WriteLine(NextNumber)
        Loop
        FileClose(1)
    End Sub ' of DisplayFile

End Module