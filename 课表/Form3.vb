Public Class Form3
#Region "<变量>"
    Dim i, j, k As Integer
    Dim xn, yn As Single
    Dim imd As Boolean
    Dim b1 As String
#End Region
#Region "<窗体加载>"
    Private Sub Form3_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Width = Label2.Left + Label2.Width
        PictureBox2.Left = Width - PictureBox2.Width
        PictureBox1.BringToFront()
        Call Dq()
        Call Xianshi()
        Timer1.Enabled = True
        If Hour(Now) >= 17 And vb_cday = vb_sday Then
            vb_cday = (vb_sday) Mod 7 + 1
            Form10.ComboBox7.Text = vb_cday
        End If
    End Sub
    Private Sub M_down(sender As Object, e As MouseEventArgs) Handles PictureBox1.MouseDown, Label1.MouseDown, Label2.MouseDown, Label3.MouseDown, Me.MouseDown, Label4.MouseDown
        imd = True
        xn = e.X : yn = e.Y
    End Sub
    Private Sub M_move(sender As Object, e As MouseEventArgs) Handles PictureBox1.MouseMove, Label1.MouseMove, Label2.MouseMove, Label3.MouseMove, Me.MouseMove, Label4.MouseMove
        If imd Then
            Me.Left = Me.Left + e.X - xn
            Me.Top = Me.Top + e.Y - yn
        End If
    End Sub
    Private Sub M_up(sender As Object, e As MouseEventArgs) Handles PictureBox1.MouseUp, Label1.MouseUp, Label2.MouseUp, Label3.MouseUp, Me.MouseUp, Label4.MouseUp
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
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Call Xianshi()
    End Sub
#End Region
#Region "<核心代码>"
    Sub Xianshi()
#Region "<-1内核，处理信息>"
        Dim tn As Integer
        tn = 100 * Hour(Now) + Minute(Now) + 1
        If vb_sday = 1 Then
            i = 1
        ElseIf vb_sday >= 2 And vb_sday <= 5 Then
            i = 2
        ElseIf vb_sday = 6 Then
            i = 3
        Else : i = 4
        End If
        For j = 1 To 50
            If ti(i, j) >= tn Then Exit For
        Next j
        Label1.Text = " " & tm(i, j - 1, 1) & ":" & tm(i, j - 1, 2) & "-" & tm(i, j, 1) & ":" & tm(i, j, 2)
        bt = Val(((60 * Hour(Now) + Minute(Now) + 1 / 60 * Second(Now)) - (tm(i, j - 1, 1) * 60 + tm(i, j - 1, 2))) / ((60 * tm(i, j, 1) + tm(i, j, 2)) - (tm(i, j - 1, 1) * 60 + tm(i, j - 1, 2))))
        If bt <= 0 Then bt = 0
        PictureBox1.Width = Me.Width * bt + 1
        Dim u As Integer
        Dim a1 As String, a2 As String
        If Mid(tm(i, j - 1, 3), 1, 5) = "Class" Then
            u = Val(Mid(tm(i, j - 1, 3), 6, 1))
            a1 = iClass(vb_cday, u)
        Else
            a1 = tm(i, j - 1, 3)
        End If
        a2 = CStr(Int(bt * 1000) / 10)
#End Region'a1显示内容,a2进度,b1储存a1
#Region "<-1显示>"
#Region "<---2变色块>"
        If ofbs = 0 Then
            PictureBox1.BackColor = Color.SpringGreen
            PictureBox2.BackColor = Color.Tomato
        ElseIf ofbs = 1 Then
            PictureBox2.BackColor = Color.DeepSkyBlue
            If tm(i, j - 1, 4) = "1" Then
                If a2 < 95 Then PictureBox1.BackColor = Color.Orange Else PictureBox1.BackColor = Color.OrangeRed
            ElseIf tm(i, j - 1, 4) = "2" Then
                If a2 < 80 Then PictureBox1.BackColor = Color.SpringGreen Else PictureBox1.BackColor = Color.Yellow
            ElseIf tm(i, j - 1, 4) = "3" Then
                PictureBox1.BackColor = Color.Violet
            Else
                PictureBox1.BackColor = Color.Silver
            End If
        Else
            PictureBox2.BackColor = Color.DeepSkyBlue
            If tm(i, j - 1, 4) = "1" Then
                PictureBox1.BackColor = Color.FromArgb(255, 255 - (255 - 69) * a2 / 100, 0) 'Yellow>OrangeRed
            ElseIf tm(i, j - 1, 4) = "2" Then
                PictureBox1.BackColor = Color.FromArgb(255 * a2 / 100, 255, 127 - 127 * a2 / 100) 'SpringGreen>Yellow
            ElseIf tm(i, j - 1, 4) = "3" Then
                PictureBox1.BackColor = Color.FromArgb(216 + (238 - 216) * a2 / 100, 143 - (143 - 130) * a2 / 100, 216 + (238 - 216) * a2 / 100) 'Thistle>Violet
            Else
                PictureBox1.BackColor = Color.Silver
            End If
        End If
#End Region
#Region "<---2显示及后续处理>"
        If a1 <> b1 Then
            If Mid(tm(i, j - 1, 3), 1, 5) = "Class" Then
                '控制form4收缩
                If zd = True Then
                    zd = False
                    tick = 20
                    Form4.Timer2.Enabled = True
                End If
                '上课重排选项
                If ofcp Then
                    Call Pdq()
                    Call Skcp()
                End If
            ElseIf j >= 3 And Mid(tm(i, j - 2, 3), 1, 5) = "Class" Then '上一项事项是上课的时候
                '控制form4展开
                If zd = False Then
                    zd = True
                    tick = 0
                    Form4.Timer2.Enabled = True
                End If
            End If
        End If
        If a1 <> b1 Then
            If tm(i, j - 1, 3) = "Class9" And vb_cday = vb_sday Then vb_cday = (vb_sday) Mod 7 + 1
            Form10.ComboBox7.Text = vb_cday
        End If
        If Len(a1) <= 3 Then
            Label3.BringToFront()
        Else
            Label4.BringToFront()
        End If
        Label3.Text = a1
        Label4.Text = a1
        If ch(21) = 0 Then
            Label2.Text = a2
        Else
            Label2.Text = CStr(Int((1 - bt) * 1000) / 10)
        End If
        b1 = a1 '赋值
        PictureBox2.BringToFront()
        PictureBox1.BringToFront()
#End Region
    End Sub
#End Region
#End Region
End Class