Public Class Form5
    Dim xn, yn As Single
    Dim imd As Boolean
    Dim zx As Boolean
    Dim lb(0 To 9) As Label
    Private Sub Form5_Load(sender As Object, e As EventArgs) Handles Me.Load
        Call Lbjz()
        imd = False
        Me.Height = Label10.Top + Label10.Height
        Me.Width = Label1.Width
        ho = Label1.Top
        hn = Me.Height
        zx = True
        zrd = True
        Call Xian()
        Timer1.Enabled = True
        'cpp
        Me.Height = Me.PictureBox1.Height + (Me.Label1.Height - 2) * zr_j(zr_t)
        If zr_t = 0 Then Me.Height = Me.PictureBox1.Height - 2
        hn = Me.Height
        Me.Tag = Me.Width & ":" & Me.Height
    End Sub
    Private Sub PictureBox1_MouseDown(sender As Object, e As MouseEventArgs) Handles PictureBox1.MouseDown
        xn = e.X : yn = e.Y
        imd = True
    End Sub
    Private Sub PictureBox1_MouseMove(sender As Object, e As MouseEventArgs) Handles PictureBox1.MouseMove
        If imd Then
            Me.Left = Me.Left + e.X - xn
            Me.Top = Me.Top + e.Y - yn
        End If
    End Sub
    Private Sub Timer1_zrtick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Call Xian()
    End Sub
    Private Sub Timer2_zrtick(sender As Object, e As EventArgs) Handles Timer2.Tick
        If zrd = True Then
            zrtick += 1
        Else zrtick -= 1
        End If
        'zrtick调整，防止重复操作
        If zrtick < 0 Then zrtick = 0
        If zrtick > 20 Then zrtick = 20
        Me.Height = ho + (hn - ho) * zrtick / 20
        If zrtick = 20 Or zrtick = 0 Then Timer2.Enabled = False
    End Sub
    Private Sub PictureBox1_MouseUp(sender As Object, e As MouseEventArgs) Handles PictureBox1.MouseUp
        imd = False
        If oftk = True Then
            If Me.Left < 0 Then
                Me.Left = 0
            ElseIf Me.Left > sw - Me.Width Then
                Me.Left = sw - Me.Width
            End If
            If Me.Top < 0 Then
                Me.Top = 0
            ElseIf Me.Top > sh - Me.Height Then
                Me.Top = sh - Me.Height
            End If
        End If
    End Sub
    Private Sub PictureBox1_mouseClick(sender As Object, e As MouseEventArgs) Handles PictureBox1.MouseClick
        If e.Button = MouseButtons.Right Then
            zrd = Not (zrd)
            If Timer2.Enabled = False Then
                If zrd = True Then
                    zrtick = 0
                    Timer2.Enabled = True
                Else
                    zrtick = 20
                    Timer2.Enabled = True
                End If
            End If
        End If
    End Sub
    Sub Lbjz()
        lb(0) = Label1 : lb(1) = Label2 : lb(2) = Label3 : lb(3) = Label4 : lb(4) = Label5 : lb(5) = Label6 : lb(6) = Label7 : lb(7) = Label8 : lb(8) = Label9 : lb(9) = Label10
    End Sub
    Sub Xian()
        If ch(9) = 1 Then
            For i = 1 To 10
                If zr_(zr_t, i, 1) = "#define" Then
                    lb(i - 1).Text = zr_c(zr_ci)
                Else
                    lb(i - 1).Text = zr_(zr_t, i, 1) & " " & zr_(zr_t, i, 2)
                End If
            Next
        Else
        End If
    End Sub
End Class