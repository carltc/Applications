Public Class Form2

    Dim img As Bitmap

    Private Sub Button3_Click(sender As System.Object, e As System.EventArgs) Handles Button3.Click


        Dim bits As BitArray
        Dim bytes(0 To 0) As BitArray
        Dim chars() As Char
        chars = TextBox1.Text.ToArray()

        For i = 0 To UBound(chars)
            Dim binary As String
            Dim binaries() As Char
            binary = DecimalToBinary(Asc(chars(i)))
            binaries = binary.ToArray
            If binaries.Length > 8 Then
                MsgBox("Binary too long")
                Exit Sub
            ElseIf binaries.Length <= 7 Then
                Dim l As Integer = binaries.Length
                Array.Resize(binaries, 8)
                For j = 0 To 6
                    binaries(6 - j + 1) = binaries(6 - j)
                Next
                For j = 0 To (8 - l - 1)
                    binaries(j) = "0"
                Next
            Else
                'MsgBox("Other error.")
                'Exit Sub
            End If
            bits = New BitArray(8)
            For j = 0 To 7
                If binaries(j) = "0" Then
                    bits(j) = False
                Else
                    bits(j) = True
                End If
            Next
            bytes(UBound(bytes)) = bits
            ReDim Preserve bytes(0 To UBound(bytes) + 1)
        Next

        Dim binarycode As String = ""
        For i = 0 To UBound(bytes) - 1
            For j = 0 To 7
                If bytes(i)(j) = True Then
                    binarycode = binarycode & "1"
                Else
                    binarycode = binarycode & "0"
                End If
            Next
        Next

        Dim bmp As Bitmap = New Bitmap(img)
        Dim ints() As Char
        ints = binarycode.ToArray
        Dim w, h As Integer
        w = 0
        h = 0

        For i = 0 To UBound(ints)
            If ints(i) = "0" Then
                Dim colour As Color
                colour = bmp.GetPixel(w, h)
                If colour.R <= 255 And colour.R > 0 Then
                    colour = Color.FromArgb(colour.R - 1, colour.G, colour.B)
                ElseIf colour.G <= 255 And colour.G > 0 Then
                    colour = Color.FromArgb(colour.R, colour.G - 1, colour.B)
                ElseIf colour.B <= 255 And colour.B > 0 Then
                    colour = Color.FromArgb(colour.R, colour.G, colour.B - 1)
                ElseIf colour.R < 255 And colour.R >= 0 Then
                    colour = Color.FromArgb(colour.R + 1, colour.G, colour.B)
                ElseIf colour.G < 255 And colour.G >= 0 Then
                    colour = Color.FromArgb(colour.R, colour.G + 1, colour.B)
                ElseIf colour.B < 255 And colour.B >= 0 Then
                    colour = Color.FromArgb(colour.R, colour.G, colour.B + 1)
                End If
                bmp.SetPixel(w, h, colour)
            End If
            w = w + 1
            If w >= bmp.Width Then
                w = 0
                h = h + 1
            ElseIf h >= bmp.Height Then
                Exit For
            End If
        Next

        PictureBox1.Image = bmp

    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click

        Dim loadfile As New OpenFileDialog
        Dim response As DialogResult
        response = loadfile.ShowDialog

        img = New Bitmap(loadfile.FileName)

    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click

        Dim savefile As New SaveFileDialog
        Dim response As DialogResult
        response = savefile.ShowDialog

        PictureBox1.Image.Save(savefile.FileName & ".jpg")

    End Sub
End Class