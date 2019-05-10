Public Class Form2
#Region "定义"
    Dim i As Integer
    Dim imd As Boolean
    Dim xn As Single, yn As Single  'me.move 中用到
    Dim xe As Single, ye As Single
    Dim tick As Integer
    '与有关时钟
    Dim hx As String, hl As String
    Dim mx As String, ml As String
    Dim sx As String
    Dim q As Integer
    Dim t(0 To 4) As String : Dim tu(0 To 4) As String : Dim lo(0 To 4) As Boolean
    Dim pb(0 To 1, 0 To 3) As PictureBox
    Dim eh As Integer
#End Region
#Region "代码"
    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Width = Label2.Left + Label2.Width
        Call Pbjz()
        PictureBox9.Image = Image.FromFile(F(0, 0, 2))
        hl = G(Hour(Now)) : ml = G(Minute(Now))
        tu(1) = Mid(hl, 1, 1) : tu(2) = Mid(hl, 2, 1)
        tu(3) = Mid(ml, 1, 1) : tu(4) = Mid(ml, 2, 1)
        For i = 1 To 4
            pb(0, i - 1).Image = Image.FromFile(F(Val(tu(i)), 0, 1)) '因数据类型没设好，曾出现了bug
            pb(1, i - 1).Image = Image.FromFile(F(Val(tu(i) + 1) Mod M(i), 0, 1))
        Next i
        sx = Second(Now)
        Label1.Text = G(sx)
        Timer1.Enabled = True
        Timer2.Enabled = True
        eh = Label2.Top + Label2.Height
    End Sub
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        hx = G(Hour(Now)) : mx = G(Minute(Now))
        sx = Second(Now)
        Label1.Text = G(sx)
        If hx <> hl Or mx <> ml Then
            t(1) = Mid(hx, 1, 1) : t(2) = Mid(hx, 2, 1)
            t(3) = Mid(mx, 1, 1) : t(4) = Mid(mx, 2, 1)
            For i = 1 To 4
                If t(i) <> tu(i) Then lo(i) = True
            Next i
            q = 0 : Timer2.Enabled = True
            hl = G(Hour(Now)) : ml = G(Minute(Now))
            tu(1) = Mid(hl, 1, 1) : tu(2) = Mid(hl, 2, 1)
            tu(3) = Mid(ml, 1, 1) : tu(4) = Mid(ml, 2, 1)
        End If
    End Sub
    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick
        q = q + 1
        If q <= 4 Then
            For i = 1 To 4
                If lo(i) = True Then
                    pb(0, i - 1).Top = pb(0, i - 1).Top - eh / 4
                    pb(1, i - 1).Top = pb(1, i - 1).Top - eh / 4
                End If
            Next i
        End If
        If q = 4 Then
            For i = 1 To 4
                If lo(i) = True Then
                    pb(0, i - 1).Image = Image.FromFile(F(Val(tu(i)), 0, 1))
                    pb(1, i - 1).Image = Image.FromFile(F(Val(tu(i) + 1) Mod M(i), 0, 1))
                    pb(0, i - 1).Top = 0 : pb(1, i - 1).Top = eh
                End If
            Next i
            For i = 1 To 4
                lo(i) = False
            Next i
            Timer2.Enabled = False
        End If
    End Sub
    '以下部分是加载部分
    Private Sub M_down(sender As Object, e As MouseEventArgs) Handles PictureBox1.MouseDown, PictureBox2.MouseDown, PictureBox3.MouseDown, PictureBox4.MouseDown, PictureBox9.MouseDown, Me.MouseDown, Label1.MouseDown
        imd = True
        xn = e.X : yn = e.Y
    End Sub
    Private Sub M_move(sender As Object, e As MouseEventArgs) Handles PictureBox1.MouseMove, PictureBox2.MouseMove, PictureBox3.MouseMove, PictureBox4.MouseMove, PictureBox9.MouseMove, Me.MouseMove, Label1.MouseMove
        If imd Then
            Me.Left = Me.Left + e.X - xn
            Me.Top = Me.Top + e.Y - yn
        End If
    End Sub
    Private Sub M_up(sender As Object, e As MouseEventArgs) Handles PictureBox1.MouseUp, PictureBox2.MouseUp, PictureBox3.MouseUp, PictureBox4.MouseUp, PictureBox9.MouseUp, Me.MouseUp, Label1.MouseUp
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
    Function F(X As Integer, Y As Integer, z As Integer) As String
        If z = 1 Then
            F = Application.StartupPath & "\内部文件\UWP\" & Mid(Str(X), 2, 1) & "number.bmp"
        ElseIf z = 2 Then
            F = Application.StartupPath & "\内部文件\UWP\冒号.bmp"
        Else F = ""
        End If
    End Function
    Function G(X As String) As String
        If Val(X) >= 0 And Val(X) <= 9 Then '曾出现了小bug,因为用字符比较
            G = 0 & X
        Else : G = X
        End If
    End Function
    Function M(X As Integer) As Integer
        If X = 1 Then
            M = 3
        ElseIf X = 3 Then
            M = 6
        Else : M = 10
        End If
    End Function
    Sub Pbjz()
        pb(0, 0) = PictureBox1 : pb(0, 1) = PictureBox2 : pb(0, 2) = PictureBox3 : pb(0, 3) = PictureBox4
        pb(1, 0) = PictureBox5 : pb(1, 1) = PictureBox6 : pb(1, 2) = PictureBox7 : pb(1, 3) = PictureBox8
    End Sub
#End Region
End Class