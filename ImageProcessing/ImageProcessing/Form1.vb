Public Class Form1

    Dim colorbox As PictureBox = New PictureBox
    Dim picbox As PictureBox = New PictureBox
    Dim processbox As PictureBox = New PictureBox
    Dim x, y, n, m As Integer
    Dim accuracy1 As Double
    Dim mouseisdown As Boolean = False
    Dim locations(,) As Boolean
    Dim loc(0 To 0) As Point
    Dim pointloc(0 To 0) As Array
    Dim averagex, averagey As Integer
    Dim res As Integer
    Dim groupdist As Integer
    Dim locs(0 To 5, 0 To 0) As Point

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ListBox1.Items.Add("koh_light_low")
        ListBox1.Items.Add("images")
        ListBox1.Items.Add("lego_keyring_torch_2")
        ListBox1.Items.Add("UV_LED_63850")
        ListBox1.Items.Add("Diver1")
        ListBox1.Items.Add("Diver2")
        ListBox1.Items.Add("Diver3")
        ListBox1.Items.Add("Car1")
        ListBox1.Items.Add("Car2")
        ListBox1.Items.Add("Car3")
        ListBox1.Items.Add("Car4")

        Me.StartPosition = FormStartPosition.CenterScreen
        picbox.Image = My.Resources.koh_light_low
        picbox.Location = New Point(0, 0)
        picbox.Height = picbox.Image.Height
        picbox.Width = picbox.Image.Width
        Me.Controls.Add(picbox)
        Me.Height = (picbox.Height) + ((DisplayRectangle.Height - ClientRectangle.Height) * 2)
        Me.Width = (picbox.Width * 2) + ((DisplayRectangle.Width - ClientRectangle.Width) * 2)
        AddHandler picbox.Paint, AddressOf paint_picbox
        accuracy1 = 0.9
        TrackBar1.SetRange(900, 1000)
        TrackBar1.RightToLeftLayout = RightToLeftLayout
        processbox.Location = New Point(picbox.Width, 0)
        processbox.Height = My.Resources.koh_light_low.Height
        processbox.Width = My.Resources.koh_light_low.Width
        processbox.BorderStyle = BorderStyle.FixedSingle
        Me.Controls.Add(processbox)
        AddHandler processbox.Paint, AddressOf paint_processbox
        n = 0
        m = 0
    End Sub

    Sub imageprocess(ByVal accuracy As Double)

        Dim btmcheck As New Bitmap(picbox.Image)
        Dim btmnew As New Bitmap(btmcheck.Width, btmcheck.Height)
        'ReDim locations(0 To btmcheck.Width - 1, 0 To btmcheck.Height - 1)
        n = 0
        ReDim loc(1)

        For i = res To (btmcheck.Width - 1) / res
            For j = res To (btmcheck.Height - 1) / res
                Dim colorcheck As Color = btmcheck.GetPixel(res * i, res * j)
                If colorcheck.GetBrightness > accuracy Then
                    btmnew.SetPixel(res * i, res * j, Color.Black)
                    'locations(i, j) = True  
                    If btmcheck.GetPixel(res * (i + 1), res * (j)).GetBrightness > accuracy And _
                    btmcheck.GetPixel(res * (i - 1), res * (j)).GetBrightness > accuracy And _
                    btmcheck.GetPixel(res * (i), res * (j + 1)).GetBrightness > accuracy And _
                    btmcheck.GetPixel(res * (i), res * (j - 1)).GetBrightness > accuracy Then

                        loc(n) = New Point(res * i, res * j)
                        ReDim Preserve loc(n + 2)
                        n = n + 1
                    End If
                Else
                    'locations(i, j) = False
                End If
            Next
        Next

        'For i = 0 To UBound(locations) / 2 - 1
        '    For j = 0 To UBound(locations) / 2 - 1
        '        If locations(i, j) Then
        '            btmnew.SetPixel(i, j, Color.Black)
        '        End If
        '    Next
        'Next

        For t = 0 To UBound(loc) - 1
            btmnew.SetPixel(loc(t).X, loc(t).Y, Color.Green)
        Next t


        processbox.Image = btmnew

    End Sub

    Private Sub trackbar1_down(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TrackBar1.MouseDown

        mouseisdown = True

    End Sub

    Private Sub trackbar1_scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TrackBar1.MouseMove

        If mouseisdown Then
            accuracy1 = (TrackBar1.Value) / 1000

            Call imageprocess(accuracy1)

            Call pointmarker()

            Label1.Text = accuracy1

        End If

    End Sub

    Private Sub trackbar1_scroll2(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TrackBar1.MouseUp

        mouseisdown = False

    End Sub

    Private Sub ListBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListBox1.SelectedIndexChanged

        If ListBox1.SelectedItem = ListBox1.Items.Item(0) Then
            picbox.Image = My.Resources.koh_light_low
        ElseIf ListBox1.SelectedItem = ListBox1.Items.Item(1) Then
            picbox.Image = My.Resources.images
        ElseIf ListBox1.SelectedItem = ListBox1.Items.Item(2) Then
            picbox.Image = My.Resources.lego_keyring_torch_2
        ElseIf ListBox1.SelectedItem = ListBox1.Items.Item(3) Then
            picbox.Image = My.Resources.UV_LED_63850
        ElseIf ListBox1.SelectedItem = ListBox1.Items.Item(4) Then
            picbox.Image = My.Resources.scuba_diver_shining_torch_by_table_coral_sami_sarkis
        ElseIf ListBox1.SelectedItem = ListBox1.Items.Item(5) Then
            picbox.Image = My.Resources.www_gettyimages_com_detail_84430601
        ElseIf ListBox1.SelectedItem = ListBox1.Items.Item(6) Then
            picbox.Image = My.Resources.diver_shining_torch_by_sunlight_sami_sarkis
        ElseIf ListBox1.SelectedItem = ListBox1.Items.Item(7) Then
            picbox.Image = My.Resources.headlights_2resize
        ElseIf ListBox1.SelectedItem = ListBox1.Items.Item(8) Then
            picbox.Image = My.Resources.LongExposureTrails
        ElseIf ListBox1.SelectedItem = ListBox1.Items.Item(9) Then
            picbox.Image = My.Resources.car_headlights
        ElseIf ListBox1.SelectedItem = ListBox1.Items.Item(10) Then
            picbox.Image = My.Resources.SuperStock_4102_15814
        End If

    End Sub

    Sub pointmarker()

        averagex = 0
        averagey = 0
       
        For n1 = 0 To UBound(loc)


            averagex = averagex + loc(n1).X
            averagey = averagey + loc(n1).Y

        Next n1

        If UBound(loc) <= 1 Then
        Else
            averagex = averagex / (UBound(loc) - 1)
            averagey = averagey / (UBound(loc) - 1)
        End If

        Label2.Text = averagex
        Label3.Text = averagey

        picbox.Height = picbox.Image.Height
        picbox.Width = picbox.Image.Width
        processbox.Location = New Point(picbox.Image.Width, 0)
        processbox.Height = picbox.Image.Height
        processbox.Width = picbox.Image.Width
        Me.Height = (picbox.Height) + ((DisplayRectangle.Height - ClientRectangle.Height) * 2)
        Me.Width = (picbox.Width * 2) + ((DisplayRectangle.Width - ClientRectangle.Width) * 2)

        Me.Refresh()

    End Sub

    Sub paint_picbox(ByVal sender As System.Object, ByVal e As PaintEventArgs)

        e.Graphics.DrawLine(Pens.Red, New Point(averagex + 10, averagey), New Point(averagex - 10, averagey))
        e.Graphics.DrawLine(Pens.Red, New Point(averagex, averagey + 10), New Point(averagex, averagey - 10))

    End Sub

    Sub paint_processbox(ByVal sender As System.Object, ByVal e As PaintEventArgs)

        e.Graphics.DrawLine(Pens.Red, New Point(averagex + 10, averagey), New Point(averagex - 10, averagey))
        e.Graphics.DrawLine(Pens.Red, New Point(averagex, averagey + 10), New Point(averagex, averagey - 10))

        For r = 0 To UBound(locs, 2)
            e.Graphics.DrawLine(Pens.Salmon, New Point(locs(5, r).X + 2, locs(5, r).Y), New Point(locs(5, r).X - 2, locs(5, r).Y))
            e.Graphics.DrawLine(Pens.Salmon, New Point(locs(5, r).X, locs(5, r).Y + 2), New Point(locs(5, r).X, locs(5, r).Y - 2))
        Next

    End Sub


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        If IsNumeric(TextBox1.Text) Then
            res = TextBox1.Text
        Else
            MsgBox("Enter a valid number!")
            Exit Sub
        End If


        Call imageprocess(TrackBar1.Value / 1000)

        Call pointmarker()

        Call splitter()

        Me.Refresh()

    End Sub


    Sub splitter()

        Dim r As Integer
        Dim u, dw, le, ri As Boolean
        Dim btmtest As Bitmap = New Bitmap(picbox.Image)
        r = 0
        u = False
        dw = False
        le = False
        ri = False

            For h = 0 To UBound(loc)

                Dim up As Point = New Point(loc(h).X, loc(h).Y - res)
                Dim down As Point = New Point(loc(h).X, loc(h).Y + res)
                Dim left As Point = New Point(loc(h).X - res, loc(h).Y)
                Dim right As Point = New Point(loc(h).X + res, loc(h).Y)
                For f = 0 To UBound(loc)
                    If loc(f).X = up.X And loc(f).Y = up.Y Then
                        u = True
                        locs(1, r) = loc(f)
                    ElseIf loc(f).X = down.X And loc(f).Y = down.Y Then
                        dw = True
                        locs(2, r) = loc(f)
                    ElseIf loc(f).X = left.X And loc(f).Y = left.Y Then
                        le = True
                        locs(3, r) = loc(f)
                ElseIf loc(f).X = right.X And loc(f).Y = right.Y Then
                    ri = True
                    locs(4, r) = loc(f)
                    End If
                Next

            If u And dw And le And ri Then
                locs(5, r) = loc(h)
                r = r + 1
                ReDim Preserve locs(0 To 5, 0 To r)
                Label1.Text = "Yes"
            End If

            Next


    End Sub



End Class
