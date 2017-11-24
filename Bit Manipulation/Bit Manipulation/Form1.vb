Public Class Form1

    Dim Byte1 As Byte
    Dim Bit1 As Integer

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click

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

        Label1.Text = binarycode

        Dim resolution As Integer = Math.Sqrt(binarycode.Length)
        Dim res As Integer = 200 / resolution
        Dim bmp As Bitmap = New Bitmap(res * resolution, res * resolution)
        Dim img As Bitmap = New Bitmap(My.Resources.taxibot_test_rig_controller)
        Dim scale, scaleX, scaleY As Single
        scaleX = bmp.Width / img.Width
        scaleY = bmp.Height / img.Height
        If scaleX > scaleY Then
            scale = scaleX
        Else
            scale = scaleY
        End If
        Label2.Text = scale
        Dim g As Graphics = Graphics.FromImage(bmp)
        Dim ints() As Char
        ints = binarycode.ToArray
        Dim w, h As Integer
        w = 0
        h = 0

        For i = 0 To UBound(ints)
            If ints(i) = "0" Then
                Dim colour As Color
                Dim he, wi As Integer
                he = h / scale
                wi = w / scale
                If he >= img.Height Or he <= 0 Then
                    colour = Color.Black
                Else
                    colour = img.GetPixel(wi, he)
                End If
                Dim colourbrush As New SolidBrush(colour)
                g.FillRectangle(colourbrush, w, h, res, res)
            Else
                g.FillRectangle(Brushes.White, w, h, res, res)
            End If
            w = w + res
            If w >= bmp.Width Then
                w = 0
                h = h + res
            ElseIf h >= bmp.Height Then
                Exit For
            End If
        Next

        PictureBox1.Image = bmp

    End Sub

    Private Sub Timer1_Tick(sender As System.Object, e As System.EventArgs) Handles Timer1.Tick

        TextBox1.Text = TextBox1.Text & "A"
        Button1.PerformClick()

    End Sub

    Private Sub Form1_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

        Timer1.Interval = 5

        Dim binaries() As Char
        binaries = DecimalToBinary(Asc("°")).ToArray
        If binaries.Length > 8 Then
            MsgBox("Binary too long")
            'Exit Sub
        ElseIf binaries.Length <= 7 Then
            Dim l As Integer = binaries.Length
            Array.Resize(binaries, 80)
            'For j = 0 To 6
            '    binaries(6 - j + 1) = binaries(6 - j)
            'Next
            'For j = 0 To (8 - l - 1)
            '    binaries(j) = "0"
            'Next
        Else
            MsgBox("Other error.")
            'Exit Sub
        End If
        Dim sp As String = ""
        For i = 0 To UBound(binaries)
            sp = sp & binaries(i)
        Next
        Label1.Text = binaries

    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        If Timer1.Enabled = True Then
            Timer1.Enabled = False
        Else
            Timer1.Enabled = True
        End If
    End Sub

End Class
