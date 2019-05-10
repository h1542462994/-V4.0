Public Class Form4
    Dim i As Integer
    Dim j As Integer
    Dim lb(0 To 7) As Label
    Private Sub Form4_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call Lbjz()
        Call Xian()
        imd = False
        zd = True
        Me.Height = Label8.Top + Label8.Height
        Me.Width = Label1.Width
        hio = Label1.Top
        hin = Label8.Top + Label8.Height
        Timer1.Enabled = True
    End Sub
    Dim xn, yn As Single
    Dim imd As Boolean
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
    Private Sub M_Click(sender As Object, e As EventArgs) Handles Label1.Click, Label2.Click, Label3.Click, Label4.Click, Label5.Click, Label6.Click, Label7.Click, Label8.Click
        '颜色调整
        If PictureBox1.BackColor = Color.White Then
            PictureBox1.BackColor = Color.Black
            Me.BackColor = Color.Black
            For i = 0 To 7
                lb(i).BackColor = Color.White
                lb(i).ForeColor = Color.Black
            Next
        Else
            PictureBox1.BackColor = Color.White
            Me.BackColor = Color.White
            For i = 0 To 7
                lb(i).BackColor = Color.Black
                lb(i).ForeColor = Color.White
            Next
        End If
    End Sub
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Call Xian()
    End Sub
    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick
        If zd = True Then
            tick += 1
        Else tick -= 1
        End If
        'tick调整，防止重复操作
        If tick < 0 Then tick = 0
        If tick > 20 Then tick = 20
        Me.Height = hio + (hin - hio) * tick / 20
        If tick = 20 Or tick = 0 Then Timer2.Enabled = False
    End Sub
    Private Sub PictureBox1_mouseClick(sender As Object, e As MouseEventArgs) Handles PictureBox1.MouseClick
        If e.Button = MouseButtons.Right Then
            zd = Not (zd)
            If Timer2.Enabled = False Then
                If zd = True Then
                    tick = 0
                    Timer2.Enabled = True
                Else
                    tick = 20
                    Timer2.Enabled = True
                End If
            End If
        End If
    End Sub
    Sub Lbjz()
        lb(0) = Label1
        lb(1) = Label2
        lb(2) = Label3
        lb(3) = Label4
        lb(4) = Label5
        lb(5) = Label6
        lb(6) = Label7
        lb(7) = Label8
    End Sub
    Sub Xian()
        For i = 0 To 7
            lb(i).Text = iClass(vb_cday, i + 1)
        Next
    End Sub
End Class