Module Module1

    Public Function DecimalToBinary(DecimalNum As Long) As _
       String
        Dim tmp As String
        Dim n As Long

        n = DecimalNum

        tmp = Trim(Str(n Mod 2))
        n = n \ 2

        Do While n <> 0
            tmp = Trim(Str(n Mod 2)) & tmp
            n = n \ 2
        Loop

        DecimalToBinary = tmp
    End Function

    Public Function BinaryToDecimal(Binary As String) As Long
        Dim n As Long
        Dim s As Integer

        For s = 1 To Len(Binary)
            n = n + (Mid(Binary, Len(Binary) - s + 1, 1) * (2 ^ _
                (s - 1)))
        Next s

        BinaryToDecimal = n
    End Function

End Module
