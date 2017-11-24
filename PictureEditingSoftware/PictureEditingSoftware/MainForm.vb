Public Class MainForm

    Dim imgwidth, imgheight, sh, sw, num As Integer
    Dim mousepoint, mouseinpic As Point
    Dim mouseisdown, resizedown, resizeright, resizedownright As Boolean
    Dim pictureboxes(), selectedbox As PictureBox
    Dim images() As Bitmap
    Dim topleft, topright, bottomleft, bottomright As Point
    Dim sqheight, sqwidth, sbn As Integer
    Dim ratio As Single
    Dim scaling As Double
    Dim groupboxes() As GroupBox

    Private Sub MainForm_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        sh = Screen.PrimaryScreen.WorkingArea.Height
        sw = Screen.PrimaryScreen.WorkingArea.Width
        Me.Location = New Point(0, 0)
        Me.Size = New Size(sw, sh)
        TextBox1.Location = New Point(sw - 120, 12)
        TextBox2.Location = New Point(sw - 120, 38)
        Label3.Location = New Point(sw - 199, 15)
        Label4.Location = New Point(sw - 196, 41)
        num = 1
        mouseisdown = False
        resizedown = False
        resizeright = False
        resizedownright = False
        TextBox1.Text = "1000"
        TextBox2.Text = "1000"
    End Sub

    Private Sub MainForm_Paint(sender As System.Object, e As System.Windows.Forms.PaintEventArgs) Handles MyBase.Paint
        sqheight = imgheight
        sqwidth = imgwidth
        ratio = 1
        Do
            sqheight = ratio * sqheight
            sqwidth = ratio * sqwidth
            ratio = ratio - 0.01
        Loop Until sqheight < 0.6 * sh And sqwidth < 0.6 * sw
        Label5.Text = ratio
        topleft = New Point((sw - sqwidth) * 0.5, (sh - sqheight) * 0.5)
        topright = New Point((sw + sqwidth) * 0.5, (sh - sqheight) * 0.5)
        bottomleft = New Point((sw - sqwidth) * 0.5, (sh + sqheight) * 0.5)
        bottomright = New Point((sw + sqwidth) * 0.5, (sh + sqheight) * 0.5)
        e.Graphics.DrawLine(Pens.Black, topleft, topright)
        e.Graphics.DrawLine(Pens.Black, bottomleft, bottomright)
        e.Graphics.DrawLine(Pens.Black, topleft, bottomleft)
        e.Graphics.DrawLine(Pens.Black, topright, bottomright)
    End Sub

    Private Sub TextBox1_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox1.TextChanged
        imgheight = Val(TextBox1.Text)
        If IsNumeric(TextBox1) And IsNumeric(TextBox2) Then
            scaling = sqheight / imgheight
        End If
        Label5.Text = scaling
        Me.Refresh()
    End Sub

    Private Sub TextBox2_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox2.TextChanged
        imgwidth = Val(TextBox2.Text)
        If IsNumeric(TextBox1) And IsNumeric(TextBox2) Then
            scaling = sqheight / imgheight
        End If
        Label5.Text = scaling
        Me.Refresh()
    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        If IsNumeric(TextBox1.Text) And IsNumeric(TextBox2.Text) Then
            ' Displays a SaveFileDialog so the user can save the Image
            Dim openFileDialog1 As New OpenFileDialog
            openFileDialog1.Filter = "Bitmap|*.bmp|JPEG|*.jpg|GIF|*.gif|TIFF|*.tif;*.tiff|PNG|*.PNG|All Files|*.*"
            openFileDialog1.Title = "Load Image File"
            Dim response As DialogResult
            response = openFileDialog1.ShowDialog()
            If response = DialogResult.Cancel Then
                Exit Sub
            End If
            ' If the file name is not an empty string open it for saving.
            If openFileDialog1.FileName <> "" Then
                ' Saves the Image via a FileStream created by the OpenFile method.
                Dim picturebox1 As New PictureBox()
                Dim img1 As New Bitmap(openFileDialog1.FileName)
                ReDim Preserve images(0 To num - 1)
                images(num - 1) = img1
                'picturebox1.Size = img1.Size
                picturebox1.Width = img1.Width * (sqheight / imgheight)
                picturebox1.Height = img1.Height * (sqheight / imgheight)
                picturebox1.Image = img1
                picturebox1.Name = "PicBox" & num
                picturebox1.Location = New Point(sw * 0.5, sh * 0.5)
                picturebox1.Cursor = Cursors.SizeAll
                picturebox1.SizeMode = PictureBoxSizeMode.StretchImage
                Dim rightclickmenu As New ContextMenu
                rightclickmenu.MenuItems.Add("Remove")
                rightclickmenu.MenuItems.Add("Rotate 90 Clockwise")
                rightclickmenu.MenuItems.Add("Rotate 90 Anti-Clockwise")
                rightclickmenu.MenuItems.Add("Flip Horizontally")
                rightclickmenu.MenuItems.Add("Flip Vertically")
                AddHandler rightclickmenu.MenuItems(0).Click, AddressOf removeclick
                AddHandler rightclickmenu.MenuItems(1).Click, AddressOf rotateclock
                AddHandler rightclickmenu.MenuItems(2).Click, AddressOf rotateanti
                AddHandler rightclickmenu.MenuItems(3).Click, AddressOf fliphori
                AddHandler rightclickmenu.MenuItems(4).Click, AddressOf flipvert
                picturebox1.ContextMenu = rightclickmenu
                ReDim Preserve pictureboxes(0 To num - 1)
                pictureboxes(num - 1) = picturebox1
                AddHandler picturebox1.MouseDown, AddressOf PictureBoxDown
                AddHandler picturebox1.MouseMove, AddressOf PictureBoxMove
                AddHandler picturebox1.MouseUp, AddressOf PictureBoxUp
                AddHandler picturebox1.Resize, AddressOf PictureBoxResize
                Me.Controls.Add(picturebox1)

                Dim gb1 As New GroupBox
                gb1.RightToLeft = Windows.Forms.RightToLeft.No
                gb1.Text = picturebox1.Name
                gb1.Size = New Size(216, 128)
                If num <= 5 Then
                    gb1.Location = New Point(12, 114 + ((num - 1) * (gb1.Size.Height + 5)))
                Else
                    gb1.Location = New Point(Me.ClientRectangle.Width - 12 - gb1.Width, 114 + ((num - 6) * (gb1.Size.Height + 5)))
                End If
                Dim tb1 As New TextBox
                Dim tb2 As New TextBox
                Dim tb3 As New TextBox
                Dim tb4 As New TextBox
                Dim lb1 As New Label
                Dim lb2 As New Label
                Dim lb3 As New Label
                Dim lb4 As New Label
                Dim bt1 As New Button
                lb1.Location = New Point(6, 22)
                lb2.Location = New Point(9, 48)
                lb3.Location = New Point(6, 74)
                lb4.Location = New Point(9, 100)
                lb1.Text = "Height:"
                lb2.Text = "Width:"
                lb3.Text = "X:"
                lb4.Text = "Y:"
                tb1.Location = New Point(53, 19)
                tb2.Location = New Point(53, 45)
                tb3.Location = New Point(53, 71)
                tb4.Location = New Point(53, 97)
                tb1.Text = picturebox1.Height
                tb2.Text = picturebox1.Width
                tb3.Text = picturebox1.Location.X - topleft.X
                tb4.Text = picturebox1.Location.Y - topleft.Y
                bt1.Location = New Point(159, 45)
                bt1.Size = New Size(46, 46)
                bt1.Text = "Set"
                bt1.Name = "Button" & num
                AddHandler bt1.Click, AddressOf bt1click
                AddHandler gb1.MouseHover, AddressOf gb1mousemove
                gb1.Controls.Add(tb1)
                gb1.Controls.Add(tb2)
                gb1.Controls.Add(tb3)
                gb1.Controls.Add(tb4)
                gb1.Controls.Add(lb1)
                gb1.Controls.Add(lb2)
                gb1.Controls.Add(lb3)
                gb1.Controls.Add(lb4)
                gb1.Controls.Add(bt1)
                Me.Controls.Add(gb1)
                ReDim Preserve groupboxes(0 To num - 1)
                groupboxes(num - 1) = gb1

                num = num + 1
            End If
        Else
            MsgBox("Choose Picture Size First")
        End If

    End Sub

    Private Sub PictureBoxDown(sender As System.Object, e As System.Windows.Forms.MouseEventArgs)
        If e.Button = Windows.Forms.MouseButtons.Left Then
            mouseinpic.X = mousepoint.X - sender.location.x
            mouseinpic.Y = mousepoint.Y - sender.location.y

            If sender.Cursor = Cursors.SizeAll Then
                mouseisdown = True
            ElseIf mouseinpic.X > sender.width * 0.95 And mouseinpic.Y > sender.height * 0.95 Then
                resizedownright = True
            ElseIf mouseinpic.X > sender.width * 0.95 Then
                resizeright = True
            ElseIf mouseinpic.Y > sender.height * 0.95 Then
                resizedown = True
            End If

            For i = 0 To UBound(pictureboxes)
                pictureboxes(i).BorderStyle = BorderStyle.None
            Next
            sender.BorderStyle = BorderStyle.FixedSingle


            selectedbox = sender

            Dim nm1 As String = ""
            For i = 6 To sender.name.length - 1
                nm1 = nm1 & sender.name(i)
            Next
            sbn = Val(nm1)
            Label6.Text = sbn
        ElseIf e.Button = Windows.Forms.MouseButtons.Right Then

            selectedbox = sender

            Dim nm1 As String = ""
            For i = 6 To sender.name.length - 1
                nm1 = nm1 & sender.name(i)
            Next
            sbn = Val(nm1)
            Label6.Text = sbn

        End If
    End Sub

    Private Sub PictureBoxMove(sender As System.Object, e As System.Windows.Forms.MouseEventArgs)
        mousepoint.X = e.X + sender.location.x
        mousepoint.Y = e.Y + sender.location.y

        If mouseisdown Then
            'If mousepoint.X - mouseinpic.X <= topleft.X Then
            '    sender.location = New Point(topleft.X, mousepoint.Y - mouseinpic.Y)
            'End If
            'If mousepoint.Y - mouseinpic.Y <= topleft.Y Then
            '    sender.location = New Point(mousepoint.X - mouseinpic.X, topleft.Y)
            'End If
            'If mousepoint.X + sender.width - mouseinpic.X >= bottomright.X Then
            '    sender.location = New Point(bottomright.X - sender.width, mousepoint.Y - mouseinpic.Y)
            'End If
            'If mousepoint.Y + sender.height - mouseinpic.Y >= bottomright.Y Then
            '    sender.location = New Point(mousepoint.X - mouseinpic.X, bottomright.Y - sender.height)
            'End If
            'If mousepoint.X - mouseinpic.X >= topleft.X And mousepoint.Y - mouseinpic.Y >= topleft.Y And mousepoint.X + sender.width - mouseinpic.X <= bottomright.X And mousepoint.Y + sender.height - mouseinpic.Y <= bottomright.Y Then
            '    sender.location = New Point(mousepoint.X - mouseinpic.X, mousepoint.Y - mouseinpic.Y)
            'End If

            sender.location = New Point(mousepoint.X - mouseinpic.X, mousepoint.Y - mouseinpic.Y)

        End If

        If resizedown Or resizedownright Or resizeright Then

            If resizedown Then
                sender.size = New Size(sender.size.width, mousepoint.Y - sender.location.y)
            ElseIf resizedownright Then
                sender.size = New Size(mousepoint.X - sender.location.x, mousepoint.Y - sender.location.y)
            ElseIf resizeright Then
                sender.size = New Size(mousepoint.X - sender.location.x, sender.size.height)
            End If

            If sender.size.height > sqheight Then
                sender.size = New Size(sender.size.width * (sqheight / sender.size.height), sender.size.height * (sqheight / sender.size.height))
            ElseIf sender.size.width > sqwidth Then
                sender.size = New Size(sender.size.width * (sqwidth / sender.size.width), sender.size.height * (sqwidth / sender.size.width))
            End If

        End If

        If e.X > sender.width * 0.95 And e.Y > sender.height * 0.95 Then
            sender.Cursor = Cursors.SizeNWSE
        ElseIf e.X > sender.width * 0.95 Then
            sender.Cursor = Cursors.SizeWE
        ElseIf e.Y > sender.height * 0.95 Then
            sender.Cursor = Cursors.SizeNS
        Else
            sender.Cursor = Cursors.SizeAll
        End If

        If sbn >= 1 Then
            scaling = imgheight / sqheight
            Dim x As Integer = (pictureboxes(sbn - 1).Location.X - topleft.X) * scaling
            Dim y As Integer = (pictureboxes(sbn - 1).Location.Y - topleft.Y) * scaling
            groupboxes(sbn - 1).Controls(2).Text = x
            groupboxes(sbn - 1).Controls(3).Text = y
        End If

    End Sub

    Private Sub PictureBoxUp(sender As System.Object, e As System.Windows.Forms.MouseEventArgs)
        mouseisdown = False
        resizedown = False
        resizeright = False
        resizedown = False
        resizedownright = False
    End Sub

    Private Sub MainForm_MouseMove(sender As System.Object, e As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseMove
        mousepoint.X = e.X
        mousepoint.Y = e.Y
        If resizedown Then
            selectedbox.Size = New Size(selectedbox.Size.Width, mousepoint.Y - selectedbox.Location.Y)
        ElseIf resizedownright Then
            selectedbox.Size = New Size(mousepoint.X - selectedbox.Location.X, mousepoint.Y - selectedbox.Location.Y)
        ElseIf resizeright Then
            selectedbox.Size = New Size(mousepoint.X - selectedbox.Location.X, selectedbox.Size.Height)
        End If

    End Sub

    Private Sub MainForm_MouseUp(sender As System.Object, e As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseUp
        mouseisdown = False
        resizedown = False
        resizeright = False
        resizedown = False
        resizedownright = False
    End Sub

    Private Sub PictureBoxResize(sender As System.Object, e As System.EventArgs)

        scaling = imgheight / sqheight

        If sbn >= 1 Then
            Dim height As Integer = pictureboxes(sbn - 1).Height * scaling
            Dim width As Integer = pictureboxes(sbn - 1).Width * scaling
            groupboxes(sbn - 1).Controls(0).Text = height
            groupboxes(sbn - 1).Controls(1).Text = width
        End If

    End Sub

    Private Sub removeclick(sender As System.Object, e As System.EventArgs)
        selectedbox.Hide()
        groupboxes(sbn - 1).Hide()
    End Sub

    Private Sub rotateclock(sender As System.Object, e As System.EventArgs)

        Dim image1 As Bitmap = New Bitmap(selectedbox.Image)
        image1.RotateFlip(RotateFlipType.Rotate90FlipNone)

        Dim newimage As Bitmap = New Bitmap(image1, selectedbox.Height, selectedbox.Width)

        'scaling = imgheight / sqheight
        'For i = 0 To newimage.Width - 1
        '    For j = 0 To newimage.Height - 1
        '        newimage.SetPixel(i, j, image1.GetPixel(i * scaling, j * scaling))
        '    Next
        'Next

        selectedbox.Size = New Size(newimage.Width, newimage.Height)
        selectedbox.Image = newimage

    End Sub

    Private Sub rotateanti(sender As System.Object, e As System.EventArgs)

        Dim image1 As Bitmap = New Bitmap(selectedbox.Image)
        image1.RotateFlip(RotateFlipType.Rotate270FlipNone)

        Dim newimage As Bitmap = New Bitmap(image1, selectedbox.Height, selectedbox.Width)

        'scaling = imgheight / sqheight
        'For i = 0 To newimage.Width - 1
        '    For j = 0 To newimage.Height - 1
        '        newimage.SetPixel(i, j, image1.GetPixel(i * scaling, j * scaling))
        '    Next
        'Next

        selectedbox.Size = New Size(newimage.Width, newimage.Height)
        selectedbox.Image = newimage

    End Sub

    Private Sub fliphori(sender As System.Object, e As System.EventArgs)

        Dim image1 As Bitmap = New Bitmap(selectedbox.Image)
        image1.RotateFlip(RotateFlipType.RotateNoneFlipX)

        Dim newimage As Bitmap = New Bitmap(image1, selectedbox.Width, selectedbox.Height)

        'scaling = imgheight / sqheight
        'For i = 0 To newimage.Width - 1
        '    For j = 0 To newimage.Height - 1
        '        newimage.SetPixel(i, j, image1.GetPixel(i * scaling, j * scaling))
        '    Next
        'Next

        selectedbox.Size = New Size(newimage.Width, newimage.Height)
        selectedbox.Image = newimage

    End Sub

    Private Sub flipvert(sender As System.Object, e As System.EventArgs)

        Dim image1 As Bitmap = New Bitmap(selectedbox.Image)
        image1.RotateFlip(RotateFlipType.RotateNoneFlipY)

        Dim newimage As Bitmap = New Bitmap(image1, selectedbox.Width, selectedbox.Height)

        'scaling = imgheight / sqheight
        'For i = 0 To newimage.Width - 1
        '    For j = 0 To newimage.Height - 1
        '        newimage.SetPixel(i, j, image1.GetPixel(i * scaling, j * scaling))
        '    Next
        'Next

        selectedbox.Size = New Size(newimage.Width, newimage.Height)
        selectedbox.Image = newimage

    End Sub

    Private Sub Button3_Click(sender As System.Object, e As System.EventArgs) Handles Button3.Click

        On Error Resume Next
        Dim finalimage As New Bitmap(imgwidth, imgheight)
        Dim g As Graphics = Graphics.FromImage(finalimage)

        g.FillRectangle(Brushes.White, 0, 0, imgwidth, imgheight)

        scaling = sqheight / imgheight



        ' Displays a SaveFileDialog so the user can save the Image
        Dim saveFileDialog1 As New SaveFileDialog
        saveFileDialog1.Filter = "Bitmap|*.bmp|JPEG|*.jpg|GIF|*.gif|TIFF|*.tif;*.tiff|PNG|*.PNG|All Files|*.*"
        saveFileDialog1.Title = "Load Image File"
        Dim response As DialogResult
        response = saveFileDialog1.ShowDialog()
        If response = DialogResult.Cancel Then
            Exit Sub
        ElseIf response = Windows.Forms.DialogResult.OK Then
            LoadingBar.Show()
            LoadingBar.ProgressBar1.Maximum = 100
            LoadingBar.ProgressBar1.Minimum = 0
            LoadingBar.ProgressBar1.Value = 0
            LoadingBar.ProgressBar1.Value = 30

            For i = 0 To UBound(images)
                Dim locationx, locationy As Single
                Dim sizex, sizey As Single
                scaling = sqheight / imgheight
                locationx = (pictureboxes(i).Location.X - topleft.X) / scaling
                locationy = (pictureboxes(i).Location.Y - topleft.Y) / scaling
                sizex = (pictureboxes(i).Width) / scaling
                sizey = (pictureboxes(i).Height) / scaling

                g.DrawImage(images(i), locationx, locationy, sizex, sizey)

                LoadingBar.ProgressBar1.Value = LoadingBar.ProgressBar1.Value + (50 / UBound(images))

            Next

        End If

        ' If the file name is not an empty string open it for saving.
        If saveFileDialog1.FileName <> "" Then

            ' Saves the Image via a FileStream created by the OpenFile method.
            finalimage.Save(saveFileDialog1.FileName)
            LoadingBar.ProgressBar1.Value = 90

        End If
        LoadingBar.ProgressBar1.Value = 100
        Resume Next

        LoadingBar.Hide()

    End Sub

    Private Sub bt1click(sender As System.Object, e As System.EventArgs)

        If sbn >= 1 Then

            Dim nm1 As String = ""
            For i = 6 To sender.name.length - 1
                nm1 = nm1 & sender.name(i)
            Next
            sbn = Val(nm1)
            Label6.Text = sbn

            scaling = imgheight / sqheight

            If IsNumeric(groupboxes(sbn - 1).Controls(0).Text) And IsNumeric(groupboxes(sbn - 1).Controls(1).Text) Then
                pictureboxes(sbn - 1).Size = New Size(Val(groupboxes(sbn - 1).Controls(1).Text) / scaling, Val(groupboxes(sbn - 1).Controls(0).Text) / scaling)
            End If

            If IsNumeric(groupboxes(sbn - 1).Controls(2).Text) And IsNumeric(groupboxes(sbn - 1).Controls(3).Text) Then
                pictureboxes(sbn - 1).Location = New Point((Val(groupboxes(sbn - 1).Controls(2).Text) / scaling) + topleft.X, (Val(groupboxes(sbn - 1).Controls(3).Text) / scaling) + topleft.Y)
            End If

        End If

    End Sub

    Private Sub gb1mousemove(sender As System.Object, e As System.EventArgs)

        'Dim nm1 As String = ""
        'For i = 6 To sender.text.length - 1
        '    nm1 = nm1 & sender.text(i)
        'Next
        'sbn = Val(nm1)
        'Label6.Text = sbn

    End Sub

    Private Sub MainForm_MouseDown(sender As System.Object, e As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseDown
        On Error Resume Next
        For i = 0 To UBound(pictureboxes)
            pictureboxes(i).BorderStyle = BorderStyle.None
        Next
        Resume Next
    End Sub

End Class
