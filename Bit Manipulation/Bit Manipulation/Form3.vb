Public Class Form3

    Dim org, enc As Bitmap

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click

        Dim loadfile As New OpenFileDialog
        Dim response As DialogResult
        response = loadfile.ShowDialog

        enc = New Bitmap(loadfile.FileName)

        PictureBox1.SizeMode = PictureBoxSizeMode.Zoom
        PictureBox1.Image = New Bitmap(loadfile.FileName)

    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click

        Dim loadfile As New OpenFileDialog
        Dim response As DialogResult
        response = loadfile.ShowDialog

        org = New Bitmap(loadfile.FileName)

        PictureBox2.SizeMode = PictureBoxSizeMode.Zoom
        PictureBox2.Image = New Bitmap(loadfile.FileName)

    End Sub

    Private Sub Button3_Click(sender As System.Object, e As System.EventArgs) Handles Button3.Click

        Dim i, j As Integer
        i = 0
        j = 0

        Dim binaries As String = ""
        Dim colourorg, colourenc As Color
        Dim chars() As Char

        Do
            colourorg = org.GetPixel(i, j)
            colourenc = enc.GetPixel(i, j)
            If colourenc = colourorg Then
                binaries = binaries & "1"
            Else
                binaries = binaries & "0"
            End If

            i = i + 1
            If i >= enc.Width Then
                i = 0
                j = j + 1
            End If

            chars = binaries.ToArray

            If UBound(chars) > 8 Then
                If chars(UBound(chars)) = "1" And binaries(UBound(chars) - 1) = "1" _
                    And chars(UBound(chars) - 2) = "1" And binaries(UBound(chars) - 3) = "1" _
                    And chars(UBound(chars) - 4) = "1" And binaries(UBound(chars) - 5) = "1" _
                    And chars(UBound(chars) - 6) = "1" And binaries(UBound(chars) - 7) = "1" _
                    Then

                    'Array.Resize(chars, UBound(chars) - 8)
                    Exit Do

                End If
            End If

        Loop Until j >= enc.Height

        Dim finalmessage As String = ""

        For i = 1 To (UBound(chars) + 1) / 8
            Dim str As String
            str = chars((i - 1) * 8) & chars((i - 1) * 8 + 1) & chars((i - 1) * 8 + 2) & _
                chars((i - 1) * 8 + 3) & chars((i - 1) * 8 + 4) & chars((i - 1) * 8 + 5) & _
                chars((i - 1) * 8 + 6) & chars((i - 1) * 8 + 7)
            Dim dec As Long = BinaryToDecimal(str)
            Dim character As Char
            character = Chr(dec)
            finalmessage = finalmessage & character
        Next

        TextBox1.Text = finalmessage

    End Sub

End Class