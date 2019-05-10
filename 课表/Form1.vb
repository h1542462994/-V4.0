Imports System.IO
Public Class Form1
#Region "<定义>"
    Dim i, j, k As Integer
    Dim pb(0 To 8) As PictureBox '控件栏目
    Dim lc, lo As Single '主界面的两个位置（zdr.left,zds.left）
    Dim mi(0 To 8) As Integer '是否鼠标经过
    Dim mold(0 To 8) As Integer '记录之前的状态
    Dim zdt(0 To 4) As Boolean '记录当窗体界面打开时隐藏的窗口
    Dim mo As Boolean '记录指针是否在里面，里面为True
    Dim tick As Integer
    Dim mx1, my1 As Integer '记录指针的位置
    Dim ys As Integer '收缩的延时
    Dim imd As Boolean
    Dim xn, yn As Single
    Dim fl, ft As Single
    Dim h, h0 As Integer
    Dim m, m0 As Integer
    Dim sfx As Single
#End Region
#Region "<窗体加载>"
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        启动窗口.Show()
        ys = 75 '(延时计算公式：ys*40(Timer1.Interval) ms，即25对应1ms)
        imd = False
        sfx = PictureBox6.Width / 100
        '窗体初始化
        Call Pbjz()
        Me.Height = pb(7).Top + pb(7).Height
        Me.Width = pb(0).Width
        Me.Top = (sh - Me.Height) / 2
        lc = sw - 2
        lo = sw - Me.Width
        Me.Left = lc
        '加载图片
        For i = 0 To 8
            pb(i).Image = Image.FromFile(Lj(i, qd(i), mi(i))) '加载图片的写法
        Next i
        mo = False
        Timer5.Enabled = True
        Timer4.Enabled = True
        Timer2.Enabled = True
        Timer3.Enabled = True
    End Sub
    Private Sub M_mousemove(sender As Object, e As EventArgs) Handles PictureBox1.MouseMove, PictureBox2.MouseMove, PictureBox3.MouseMove, PictureBox4.MouseMove, PictureBox5.MouseMove, PictureBox6.MouseMove, PictureBox7.MouseMove, PictureBox8.MouseMove, PictureBox9.MouseMove
        '加载图标
        For i = 0 To 8
            If sender Is pb(i) Then
                mi(i) = 1
            Else
                mi(i) = 0
            End If
            If mi(i) <> mold(i) Then pb(i).Image = Image.FromFile(Lj(i, qd(i), mi(i))) '加载图片的写法
        Next
        For i = 0 To 8
            mold(i) = mi(i)
        Next
    End Sub
    Private Sub M_mouseclick(sender As Object, e As EventArgs) Handles PictureBox2.MouseClick, PictureBox3.MouseClick, PictureBox4.MouseClick, PictureBox5.MouseClick, PictureBox6.MouseClick, PictureBox7.MouseClick, PictureBox8.MouseClick, PictureBox9.MouseClick
        '执行各个按钮的功能
        If File.Exists(ljc_p) Then
            Call Pdq()
        End If
        For i = 0 To 8
            If sender Is pb(i) And i <> 0 Then
                If Not (qd(6) = 1 And i <> 6 And i <> 7) Then
                    qd(i) = 1 - qd(i)
                    Select Case i
                        'Case 0
                        '    If qd(0) = 1 Then
                        '        Me.Height = pb(0).Height - 2
                        '    Else Me.Height = pb(7).Top + pb(7).Height
                        '    End If
                        '    Me.Top = (My.Computer.Screen.WorkingArea.Height - Me.Height) / 2
                        Case 1 : If qd(1) = 1 Then Form2.Show() Else Form2.Hide()
                            Form2.Left = sw * (lp(0) / 100) : Form2.Top = sh * (tp(0) / 100)
                        Case 2 : If qd(2) = 1 Then Form3.Show() Else Form3.Hide()
                            Form3.Left = sw * (lp(1) / 100) : Form3.Top = sh * (tp(1) / 100)
                        Case 3 : If qd(3) = 1 Then Form4.Show() Else Form4.Hide()
                            Form4.Left = sw * (lp(2) / 100) : Form4.Top = sh * (tp(2) / 100)
                        Case 4 : If qd(4) = 1 Then Form5.Show() Else Form5.Hide()
                            Form5.Left = sw * (lp(3) / 100) : Form5.Top = sh * (tp(3) / 100)
                        Case 5 : If qd(5) = 1 Then Form10.Show() Else Form10.Hide()
                            Form10.Left = sw * (lp(4) / 100) : Form10.Top = sh * (tp(4) / 100)
                        Case 6 : If qd(6) = 1 Then Form7.Show() : Call Tz() Else Form7.Hide() : Call Tzhy()
                            Form7.Top = 0 : Form7.Left = 0
                            Form7.Width = sw
                            Form7.Height = sh
                            Form7.PictureBox1.Width = Form2.Width : Form7.PictureBox1.Height = Form2.Height
                            Form7.PictureBox2.Width = Form3.Width : Form7.PictureBox2.Height = Form3.Height
                            Form7.PictureBox3.Width = Form4.Width : Form7.PictureBox3.Height = hin
                            Form7.PictureBox4.Width = Form5.Width : Form7.PictureBox4.Height = hn
                            Form7.PictureBox5.Width = Form10.Label55.Left + Form10.Label55.Width : Form7.PictureBox5.Height = Form10.Label55.Top + Form10.Label55.Height
                            Form7.PictureBox7.Width = Form6.Width : Form7.PictureBox7.Height = Form6.Height
                        Case 7
                            End
                        Case 8
                            If qd(8) = 1 Then Form6.Show() Else Form6.Hide()
                            Form6.Left = sw * (lp(5) / 100) : Form6.Top = sh * (tp(5) / 100)
                    End Select
                End If
                pb(i).Image = Image.FromFile(Lj(i, qd(i), mi(i))) '加载图片的写法
            End If
        Next
        For i = 0 To 8
            'pb(i).Image = Image.FromFile(lj(i, qd(i), mi(i))) '加载图片的写法
        Next
    End Sub
#End Region
#Region "<窗体移动>"
    Private Sub PictureBox1_MouseDown(sender As Object, e As MouseEventArgs) Handles PictureBox1.MouseDown
        imd = True
        xn = e.X : yn = e.Y
        fl = Left : ft = Top
    End Sub
    Private Sub PictureBox1_MouseMove(sender As Object, e As MouseEventArgs) Handles PictureBox1.MouseMove
        If imd Then
            Left = Left + e.X - xn : Top = Top + e.Y - yn
        End If
    End Sub
    Private Sub PictureBox1_MouseUp(sender As Object, e As MouseEventArgs) Handles PictureBox1.MouseUp
        imd = False
        If Left >= (sw - Width) / 2 Then
            lc = sw - 2
            lo = sw - Me.Width
            Left = lo
        Else
            lc = -Me.Width + 2
            lo = 0
            Left = lo
        End If
        If Math.Abs(Left - fl) <= 20 And Math.Abs(Top - ft) <= 20 Then
            qd(0) = 1 - qd(0)
            If qd(0) = 1 Then
                Me.Height = pb(0).Height - 2
                Me.Top = Me.Top + 3.5 * pb(0).Height
            Else Me.Height = pb(7).Top + pb(7).Height
                Me.Top = Me.Top - 3.5 * pb(0).Height
            End If
        Else
            '图标加载
            For i = 0 To 8
                mi(i) = 0
                If mi(i) <> mold(i) Then pb(i).Image = Image.FromFile(Lj(i, qd(i), mi(i))) '加载图片的写法
            Next
            For i = 0 To 8
                mold(i) = mi(i)
            Next
        End If
        If Top < 0 Then
            Top = 0
            i = 0 : pb(i).Image = Image.FromFile(Lj(i, qd(i), mi(i))) '加载图片的写法
        ElseIf Top > sh - Height Then
            Top = sh - Height
        End If
    End Sub
#End Region
#Region "<窗体收缩，展开>"
    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick '实现窗体收缩展开的较好方法
        '获取鼠标位置
        If imd = False Then '防止拖动时图标闪烁的现象
            mx1 = System.Windows.Forms.Cursor.Position.X
            my1 = System.Windows.Forms.Cursor.Position.Y
            If Left < sw / 2 Then
                If Not (mx1 <= lo + Width And (my1 >= Me.Top And my1 <= Me.Top + Me.Height)) Then
                    '图标加载
                    For i = 0 To 8
                        mi(i) = 0
                        If mi(i) <> mold(i) Then pb(i).Image = Image.FromFile(Lj(i, qd(i), mi(i))) '加载图片的写法
                    Next
                    For i = 0 To 8
                        mold(i) = mi(i)
                    Next
                End If
                If Not (mx1 <= lo + Width And (my1 >= Me.Top And my1 <= Me.Top + Me.Height)) And mo = True And Timer1.Enabled = False Then
                    mo = False
                    tick = 0
                    Timer1.Enabled = True
                ElseIf mx1 <= lc - 2 + Width And (my1 >= Me.Top And my1 <= Me.Top + Me.Height) And mo = False And Timer1.Enabled = False Then
                    mo = True
                    tick = 0
                    Timer1.Enabled = True
                ElseIf mx1 <= lo + Width And (my1 >= Me.Top And my1 <= Me.Top + Me.Height) And mo = False And Timer1.Enabled = True And tick < ys Then
                    mo = True
                    tick = 0
                    Timer1.Enabled = False
                End If
            Else
                If Not (mx1 >= lo And (my1 >= Me.Top And my1 <= Me.Top + Me.Height)) Then
                    '图标加载
                    For i = 0 To 8
                        mi(i) = 0
                        If mi(i) <> mold(i) Then pb(i).Image = Image.FromFile(Lj(i, qd(i), mi(i))) '加载图片的写法
                    Next
                    For i = 0 To 8
                        mold(i) = mi(i)
                    Next
                End If
                If Not (mx1 >= lo And (my1 >= Me.Top And my1 <= Me.Top + Me.Height)) And mo = True And Timer1.Enabled = False Then
                    mo = False
                    tick = 0
                    Timer1.Enabled = True
                ElseIf mx1 >= lc - 2 And (my1 >= Me.Top And my1 <= Me.Top + Me.Height) And mo = False And Timer1.Enabled = False Then
                    mo = True
                    tick = 0
                    Timer1.Enabled = True
                ElseIf mx1 >= lo And (my1 >= Me.Top And my1 <= Me.Top + Me.Height) And mo = False And Timer1.Enabled = True And tick < ys Then
                    mo = True
                    tick = 0
                    Timer1.Enabled = False
                End If
            End If
        End If
        '设置块关闭时的刷新
        If zes Then
            zes = False : pb(5).Image = Image.FromFile(lj(5, qd(5), mi(5)))
        End If
    End Sub
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick '实现收缩展开的核心代码
        If mo = True Then
            tick += 1
            Me.Left = lc - (lc - lo) * tick / 5
            If tick = 5 Then Timer1.Enabled = False
        Else
            tick += 1
            If tick >= ys And tick <= ys + 5 Then
                Me.Left = lo + (lc - lo) * (tick - ys) / 5
            End If
            If tick = ys + 5 Then Timer1.Enabled = False
        End If
    End Sub

    Private Sub M_mouseclick(sender As Object, e As MouseEventArgs) Handles PictureBox8.MouseClick, PictureBox7.MouseClick, PictureBox6.MouseClick, PictureBox5.MouseClick, PictureBox4.MouseClick, PictureBox3.MouseClick, PictureBox2.MouseClick, PictureBox9.MouseClick

    End Sub

    Private Sub M_mousemove(sender As Object, e As MouseEventArgs) Handles PictureBox8.MouseMove, PictureBox7.MouseMove, PictureBox6.MouseMove, PictureBox5.MouseMove, PictureBox4.MouseMove, PictureBox3.MouseMove, PictureBox2.MouseMove, PictureBox1.MouseMove, PictureBox9.MouseMove

    End Sub
#End Region
#Region "<数据更新>"
    Private Sub Timer3_Tick(sender As Object, e As EventArgs) Handles Timer3.Tick
        vbday = Weekday(Now, vbMonday)
        If vbday <> vbold Then
            vb_cday = vbday
            vb_sday = vbday
            Form10.Label33.Text = "星期" & vbday
            Form10.ComboBox6.Text = vb_sday
            Form10.ComboBox6.Enabled = True
            Form10.ComboBox7.Text = vb_cday
            Form10.ComboBox7.Enabled = True
            Call Dq()
            Call Dq1()
            Form10.ComboBox1.Enabled = True
            ch(7) = 0
            Form10.PictureBox15.Image = Image.FromFile(Application.StartupPath & "\内部文件\图4\off.bmp")
            Form10.ComboBox3.Enabled = True
            ch(8) = 0
            Form10.PictureBox17.Image = Image.FromFile(Application.StartupPath & "\内部文件\图4\off.bmp")
            If ch(24) = 1 Then
                Call Zd_Show()
            End If
        End If
        vbold = vbday
    End Sub
    Private Sub Timer5_Tick(sender As Object, e As EventArgs) Handles Timer5.Tick
        If ofzd Then
            Me.TopMost = True
            Form2.TopMost = True
            Form3.TopMost = True
            Form4.TopMost = True
            Form5.TopMost = True
        End If
    End Sub
#End Region
#Region "<函数，过程>"
    Function Lj(ByVal x As Integer, ByVal y As Integer, ByVal z As Integer) As String
        Lj = Application.StartupPath & "\内部文件\图1\" & x & (2 * y + z) & ".bmp"
    End Function
    Function Ljo(x As Integer) As String
        If x = 0 Then
            Ljo = Application.StartupPath & "\内部文件\图4\off.bmp"
        ElseIf x = 1 Then
            Ljo = Application.StartupPath & "\内部文件\图4\on.bmp"
        Else
            Ljo = Application.StartupPath & "\内部文件\图4\off-f.bmp"
        End If
    End Function
    Sub Pbjz()
        pb(0) = PictureBox1
        pb(1) = PictureBox2
        pb(2) = PictureBox3
        pb(3) = PictureBox4
        pb(4) = PictureBox5
        pb(5) = PictureBox6
        pb(6) = PictureBox7
        pb(7) = PictureBox8
        pb(8) = PictureBox9
    End Sub
    Sub Tz()
        Form2.Hide()
        Form3.Hide()
        Form4.Hide()
        Form5.Hide()
        Form10.Hide()
        Form6.Hide()
    End Sub
    Sub Tzhy()
        If qd(1) = 1 Then Form2.Show() : Form2.Left = sw * (lp(0) / 100) : Form2.Top = sh * (tp(0) / 100)
        If qd(2) = 1 Then Form3.Show() : Form3.Left = sw * (lp(1) / 100) : Form3.Top = sh * (tp(1) / 100)
        If qd(3) = 1 Then Form4.Show() : Form4.Left = sw * (lp(2) / 100) : Form4.Top = sh * (tp(2) / 100)
        If qd(4) = 1 Then Form5.Show() : Form5.Left = sw * (lp(3) / 100) : Form5.Top = sh * (tp(3) / 100)
        If qd(5) = 1 Then Form10.Show() : Form10.Left = sw * (lp(4) / 100) : Form10.Top = sh * (tp(4) / 100)
        If qd(8) = 1 Then Form6.Show() : Form6.Left = sw * (lp(5) / 100) : Form6.Top = sh * (tp(5) / 100)
    End Sub
    Private Sub Timer4_Tick(sender As Object, e As EventArgs) Handles Timer4.Tick
        If File.Exists(ljto_dm) And ofdm = True Then
            h = Hour(Now) : m = Minute(Now)
            If m <> m0 Or h <> h0 Then
                For k = 1 To 6
                    If hie(k) = h And mie(k) = m Then
                        If dmb(vb_cday, k) = True And k < 5 Then
                            Call Dms()
                        ElseIf dmb(Qm(vb_cday), k) = True And k >= 5 Then
                            Call Dms()
                        End If
                    End If
                Next
            End If
            h0 = h : m0 = m
        Else
            ch(10) = 0
            ofdm = False
        End If
        If Not (File.Exists(ljto_dm)) Then Form10.PictureBox18.Image = Image.FromFile(Ljo(2)) : Form10.Label25.ForeColor = Color.Gray
    End Sub '点名器
    Function Qm(x As Integer) As Integer
        If x = 1 Then
            Qm = 7
        Else qm = x - 1
        End If
    End Function
#End Region
End Class