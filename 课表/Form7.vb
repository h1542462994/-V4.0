Imports System.IO
Public Class Form7
#Region "<定义>"
    Dim pb(0 To 5) As PictureBox
    Dim i As Integer
    Dim xn, yn As Single
    Dim imd As Integer
    Dim l(0 To 5), t(0 To 5) As Single
    Dim tick, tick2 As Integer
#End Region
#Region "<窗体>"
    Private Sub Form7_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call Pbjz()
        sw = My.Computer.Screen.WorkingArea.Width
        sh = My.Computer.Screen.WorkingArea.Height
        For i = 0 To 5
            pb(i).Image = Image.FromFile(Application.StartupPath & "\内部文件\图2\" & (i + 1) & ".bmp")
        Next
        imd = -1
        Label1.Top = My.Computer.Screen.WorkingArea.Height / 10
        Label1.Left = (My.Computer.Screen.WorkingArea.Width - Label1.Width) / 2
        Button1.Top = Label1.Top + Label1.Height
        Button1.Left = Label1.Left
        Button2.Top = Button1.Top
        Button2.Left = Button1.Left + Button1.Width
        For i = 0 To 5
            lp(i) = Int(pb(i).Left / My.Computer.Screen.WorkingArea.Width * 1000) / 10
            tp(i) = Int(pb(i).Top / My.Computer.Screen.WorkingArea.Height * 1000) / 10
        Next
        Call Csh()
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim MyStream As New FileStream(ljc_p, FileMode.Create)
        Dim writer1 As New StreamWriter(MyStream, System.Text.Encoding.GetEncoding("gb2312")) '文字编码格式，非常重要，否则会出现乱码。
        For i = 0 To 5
            writer1.WriteLine(lp(i) & ";" & tp(i))
        Next
        writer1.Close()
        '防止高频操作
        Button1.Text = "3s"
        tick = 3
        Button1.Enabled = False
        Timer1.Enabled = True
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Call Csh()
        '防止高频操作
        Button2.Text = "3s"
        tick2 = 3
        Button2.Enabled = False
        Timer2.Enabled = True
    End Sub
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        tick = tick - 1
        Button1.Text = tick & "s"
        If tick = 0 Then
            Timer1.Enabled = False
            Button1.Enabled = True
            Button1.Text = "保存设置"
        End If
    End Sub
    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick
        tick2 = tick2 - 1
        Button2.Text = tick2 & "s"
        If tick2 = 0 Then
            Timer2.Enabled = False
            Button2.Enabled = True
            Button2.Text = "不保存"
        End If
    End Sub
    Sub Pbjz()
        pb(0) = PictureBox1 : pb(1) = PictureBox2 : pb(2) = PictureBox3 : pb(3) = PictureBox4 : pb(4) = PictureBox5 ： pb(5) = PictureBox7
    End Sub
    Private Sub M_MouseDown(sender As Object, e As MouseEventArgs) Handles PictureBox1.MouseDown, PictureBox2.MouseDown, PictureBox3.MouseDown, PictureBox4.MouseDown, PictureBox5.MouseDown, PictureBox7.MouseDown
        imd = sender.Tag
        '获取鼠标位置
        xn = Cursor.Position.X
        yn = Cursor.Position.Y
        l(imd) = pb(imd).Left
        t(imd) = pb(imd).Top
        pb(imd).BringToFront()
        lp(imd) = Int(l(imd) / My.Computer.Screen.WorkingArea.Width * 10000) / 100
        tp(imd) = Int(t(imd) / My.Computer.Screen.WorkingArea.Height * 10000) / 100
        Label1.Text = "窗口:" & F(imd) & "   Left:" & pb(imd).Left & "(" & lp(imd) & "%)" & "   Top:" & pb(imd).Top & "(" & tp(imd) & "%)"
    End Sub
    Private Sub M_MouseMove(sender As Object, e As MouseEventArgs) Handles PictureBox1.MouseMove, PictureBox2.MouseMove, PictureBox3.MouseMove, PictureBox4.MouseMove, PictureBox5.MouseMove, PictureBox7.MouseMove
        If imd <> -1 Then
            Dim xe, ye As Single
            xe = Cursor.Position.X
            ye = Cursor.Position.Y
            pb(imd).Left = l(imd) + xe - xn
            pb(imd).Top = t(imd) + ye - yn
            lp(imd) = Int(pb(imd).Left / My.Computer.Screen.WorkingArea.Width * 10000) / 100
            tp(imd) = Int(pb(imd).Top / My.Computer.Screen.WorkingArea.Height * 10000) / 100
            Label1.Text = "窗口:" & F(imd) & "   Left:" & pb(imd).Left & "(" & lp(imd) & "%)" & "   Top:" & pb(imd).Top & "(" & tp(imd) & "%)"
        End If
    End Sub
    Private Sub M_MouseUp(sender As Object, e As MouseEventArgs) Handles PictureBox1.MouseUp, PictureBox2.MouseUp, PictureBox3.MouseUp, PictureBox4.MouseUp, PictureBox5.MouseUp, PictureBox7.MouseUp
        If oftk Then
            If pb(imd).Left < 0 Then
                pb(imd).Left = 0
                lp(imd) = 0
            ElseIf pb(imd).Left + pb(imd).Width > sw Then
                pb(imd).Left = sw - pb(imd).Width
                lp(imd) = Int(100 * (100 - pb(imd).Width / sw * 100)) / 100
            End If
            If pb(imd).Top < 0 Then
                pb(imd).Top = 0
                tp(imd) = 0
            ElseIf pb(imd).Top > sh - pb(imd).Height Then
                pb(imd).Top = sh - pb(imd).Height
                tp(imd) = Int(100 * (100 - pb(imd).Height / sh * 100)) / 100
            End If
        End If
        Label1.Text = "窗口:" & F(imd) & "   Left:" & pb(imd).Left & "(" & lp(imd) & "%)" & "   Top:" & pb(imd).Top & "(" & tp(imd) & "%)"
        imd = -1
    End Sub
#End Region
#Region "<函数、过程>"
    Function F(ByVal x As Integer) As String
        If x = 0 Then
            F = "时间显示器"
        ElseIf x = 1 Then
            F = "进度条"
        ElseIf x = 2 Then
            F = "悬挂课表"
        ElseIf x = 3 Then
            F = "值日生"
        ElseIf x = 4 Then
            F = "设置"
        ElseIf x = 5 Then
            F = "倒计时"
        Else
            F = "####"
        End If
    End Function
    Sub Csh()
        '窗体位置初始化
        Call Pdq()
        For i = 0 To 5
            pb(i).Left = sw * (lp(i) / 100) : pb(i).Top = sh * (tp(i) / 100)
        Next
    End Sub
#End Region
End Class