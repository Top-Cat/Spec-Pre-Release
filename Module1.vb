Module Module1

    Sub Main() 'Maybe make a loop to stop program exiting before requested
        Dim Choice As Integer
        Dim Finished As Boolean
        Finished = False ' Does nothing
        While Not Finished
            ShowMenu()
            GetResponse(Choice)
            Select Case Choice
                Case 1 : ConvertNumber()
                Case 2 : ConvertFile()
                Case 3 : DisplayFile()
                Case 4 : Finished = True ' end program - doesn't (it ends anyway)
            End Select
            Console.ReadLine()
        End While
    End Sub ' of main

    Sub ShowMenu()
        Console.WriteLine("Please choose an option")
        Console.Writeline()
        Console.Writeline("1 - Convert a Hex numnber to Binary")
        Console.Writeline("2 - Convert a text file of Hex numbers")
        Console.WriteLine("3 - Display text file")
        Console.WriteLine("4 - Exit program")
        Console.Writeline()
    End Sub ' of ShowMenu

    Sub GetResponse(ByRef Response As Integer)
        Response = 0
        While Response > 4 Or Response < 1
            Console.Write("Enter option number: ")
            Response = Val(Console.ReadLine()) 'Will error if non-number is enter
        End While
    End Sub ' of GetResponse

    Function Binary(ByVal Hex As String) As String
        Dim Result As String
        Dim HexDigit As Integer
        Dim NoOfHexDigits As Integer
        Dim ThisHexDigit As String
        Dim BinaryEquivalent As String
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
            Else ' REALLY, just really?
            End If
            Result = Result + BinaryEquivalent 'Should use '&' but won't error. Due to hardcoded strings
        Next
        Binary = Result 'RAWR dislike
    End Function ' of Binary

    Sub ConvertNumber()
        Dim Hexadecimal As String
        Dim Converted As String
        Console.Write("Enter a Hexadecimal number: ")
        Hexadecimal = Console.ReadLine
        Converted = Binary(Hexadecimal)
        Console.WriteLine(Converted)
    End Sub ' of ConvertNumber

    Sub ConvertFile()
        Dim HexNumber As String
        Dim BinaryNumber As String
        FileOpen(1, "C:\Documents and Settings\VPCuser\My Documents\COMP1 practice\COMP1_specimen\HexData.txt", OpenMode.Input)
        Console.WriteLine()
        Do While Not EOF(1)
            HexNumber = LineInput(1)
            BinaryNumber = Binary(HexNumber)
            Console.WriteLine(BinaryNumber)
        Loop
        FileClose(1)
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