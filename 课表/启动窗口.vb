Imports System.IO
Imports System.Text
Imports System.Threading
Imports Microsoft.Win32
Imports System.Net
Public Class 启动窗口
#Region "<定义>"
    Dim imd As Boolean
    Dim xn, yn As Single
    Dim px As Single : Dim py As Single
    Dim runcode As Integer
    Dim runtext As String
    Dim srd As Boolean
    Dim ljxjt As String
#End Region
#Region "<窗口>"
#Region "<--窗体加载>"
    Private Sub 启动窗口_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Label2.Text = "时间显示器加载程序 V" & bbh
        sw = My.Computer.Screen.WorkingArea.Width
        sh = My.Computer.Screen.WorkingArea.Height
        Left = (sw - Width) / 2 : Top = (sh - Height) / 2
        imd = False
    End Sub
    Private Sub 启动窗口_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        px = Label8.Left : py = Label8.Top
        Label3.Text = "第一步：检查更新"
        Application.DoEvents()
        Dim ljupc As String = "ftp://192.168.2.195/update/program/kbupdate/Version.txt"
        Dim ljupct As String = Application.StartupPath & "\Version.txt"
        Try
            My.Computer.Network.DownloadFile(ljupc, ljupct, "", "", False, 100, True)
            If File.Exists(ljupct) Then
                Dim re1 As New StreamReader(ljupct, Encoding.GetEncoding("gb2312"))
                Dim s As String = re1.ReadLine
                Label9.Text = "更新版本 " & s
                If s <> bbh Then
                    srd = False : runcode = 2
                    Call Run()
                End If
                re1.Close()
            End If
        Catch ex As Exception

        End Try
        Label3.Text = "第二步：窗体加载及数据初始化"
        Application.DoEvents()
        ProgressBar1.Value = 20
        'Thread.Sleep(100)
        vbday = Weekday(Now, vbMonday)
        Form10.Label33.Text = "星期" & vbday
        vbold = vbday
        vb_cday = vbday
        vb_sday = vbday
        Form10.Label33.Text = "星期" & vbday
        Form10.ComboBox6.Text = vb_sday
        Form10.ComboBox7.Text = vb_cday
        For i = 1 To 7
            Form10.ComboBox6.Items.Add(i)
            Form10.ComboBox7.Items.Add(i)
        Next
        Application.DoEvents()
        Label3.Text = "第三步：检查内部文件"
        Application.DoEvents()
        ProgressBar1.Value = 40
        'Thread.Sleep(100)
        Call Xj()
        Application.DoEvents()
        Label3.Text = "第四步：检查基础配置"
        Application.DoEvents()
        ProgressBar1.Value = 60
        'Thread.Sleep(100)
        Call Xj2()
        Application.DoEvents()
        Label3.Text = "第五步：加载配置"
        Panel1.Visible = False
        Application.DoEvents()
        ProgressBar1.Value = 80
        'Thread.Sleep(100)
        Call CJz()
        If File.Exists(ljc_s) Then
            Call Cr() '加载-数据源
        End If
        If File.Exists(ljc(0)) Then
            Call Copr() '加载-不透明度
        End If
        Application.DoEvents()
        ProgressBar1.Value = 100
        Thread.Sleep(100)
        Application.DoEvents()
        For i = 0 To 20
            ch(i) = 0
        Next
        Call Dq()
        Call Dq1()
        Call Zrjz()
        Call Xian()
        Call Pdq()
        Call Cojz()
        Call Mojz()
        Form1.Show()
        Me.Hide()
    End Sub
#End Region
#Region "<--窗体设计>"
    Private Sub M_MouseDown(sender As Object, e As MouseEventArgs) Handles Me.MouseDown, Label2.MouseDown
        imd = True
        xn = e.X : yn = e.Y
    End Sub
    Private Sub M_MouseMove(sender As Object, e As MouseEventArgs) Handles Me.MouseMove, Label2.MouseMove
        If imd Then
            Left = Left + e.X - xn
            Top = Top + e.Y - yn
        End If
    End Sub
    Private Sub M_MouseUp(sender As Object, e As MouseEventArgs) Handles Me.MouseUp, Label2.MouseUp
        imd = False
    End Sub
    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click
        End
    End Sub
    Private Sub Label1_MouseMove(sender As Object, e As MouseEventArgs) Handles Label1.MouseMove
        Label1.ForeColor = Color.Red
    End Sub
    Private Sub Label1_MouseLeave(sender As Object, e As EventArgs) Handles Label1.MouseLeave
        Label1.ForeColor = Color.White
    End Sub
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged, ComboBox2.SelectedIndexChanged
        If ComboBox1.Text = "" Or ComboBox2.Text = "" Then
            Label6.BackColor = Color.Tomato
        Else
            Label6.BackColor = Color.FromArgb(0, 192, 0)
        End If
    End Sub
    Private Sub Label6_Click(sender As Object, e As EventArgs) Handles Label6.Click
        If Label6.BackColor <> Color.Tomato Then srd = True
        Dim mtx As New FileStream(ljc_s, FileMode.Create)
        Dim we1 As New StreamWriter(mtx, System.Text.Encoding.GetEncoding("gb2312"))
        we1.WriteLine(ComboBox1.Text)
        we1.WriteLine(ComboBox2.Text)
        we1.Close()
    End Sub
#End Region
#Region "<--过程与函数>"
#Region "<----#Run>"
    Sub Xj()
        If File.Exists(ljxj) Then
            Dim re1 As New StreamReader(ljxj, Encoding.GetEncoding("gb2312"))
            Dim ljxt As String
            Dim s As String
            Do Until re1.EndOfStream
                s = re1.ReadLine
                ljxt = Application.StartupPath & "\内部文件" & s
                If File.Exists(ljxt) = False Then
                    runcode = 1
                    runtext = "未找到文件[" & s & "],请联系开发人员或者将备份文件拷贝到相应位置"
                    Call Run()
                End If
            Loop
            re1.Close()
        Else
            runcode = 1
            runtext = "未找到文件[\内部文件\文件校检.txt],请联系开发人员或者将备份文件拷贝到相应位置"
            Call Run()
        End If
    End Sub
    Sub Xj2()
        Dim trun As Boolean
        trun = True
        If File.Exists(ljc_s) Then
            Dim r1 As New StreamReader(ljc_s, Encoding.GetEncoding("gb2312"))
            i = 0
            Do Until r1.EndOfStream
                i += 1
                ljxjt = Application.StartupPath & "\数据\" & r1.ReadLine & ".txt"
                If File.Exists(ljxjt) = False Then runcode = 0 : trun = False : Exit Do
            Loop
            If i <> 2 Then runcode = 0 : trun = False
            If trun = False Then
                r1.Close() : Call Run()
            End If
        Else
            runcode = 0 : Call Run()
        End If
    End Sub
    Sub Run()
        If runcode = 0 Then
            Panel1.Left = px : Panel1.Top = py
            Panel1.Visible = True
            srd = False
            Call Jz()
            Do While Not srd  'pause
                Application.DoEvents()
                Thread.Sleep(15)
            Loop
        ElseIf runcode = 1 Then
            Panel2.Left = px : Panel2.Top = py
            Panel2.Visible = True
            Label7.Text = runtext
            Do While True 'pause
                Application.DoEvents()
                Thread.Sleep(15)
            Loop
        ElseIf runcode = 2 Then
            Panel3.Left = px : Panel3.Top = py
            Panel3.Visible = True
            Try
                Dim lju(1) As String
                Dim ljut(1) As String
                lju(0) = "ftp://192.168.2.195/update/program/kbupdate/UpdateDiary.txt"
                ljut(0) = Application.StartupPath & "\UpdateDiary.txt"
                lju(1) = "ftp://192.168.2.195/update/program/kbupdate/Controls.txt"
                ljut(1) = Application.StartupPath & "\Controls.txt"
                My.Computer.Network.DownloadFile(lju(0), ljut(0), "", "", False, 100, True)
                My.Computer.Network.DownloadFile(lju(1), ljut(1), "", "", False, 100, True)
                Dim re2 As New StreamReader(ljut(0), Encoding.GetEncoding("gb2312"))
                Dim s As String = re2.ReadToEnd
                RichTextBox1.Text = s
                re2.Close()
            Catch ex As Exception

            End Try
            Do While Not srd 'pause
                Application.DoEvents()
                Thread.Sleep(15)
            Loop
        End If
    End Sub
    Sub Jz()
        ljxjt = Application.StartupPath & "\数据\"
        For Each mfl In Directory.GetFiles(ljxjt)
            Dim mfn As String = Mid(Path.GetFileName(mfl), 1, Len(Path.GetFileName(mfl)) - 4)
            If Mid(mfn, 1, 2) = "时间" Then
                ComboBox1.Items.Add(mfn.ToString)
            ElseIf Mid(mfn, 1, 2) = "课表" Then
                ComboBox2.Items.Add(mfn.ToString)
            End If
        Next
    End Sub
    Private Sub Label11_Click(sender As Object, e As EventArgs) Handles Label11.Click
        '>>new class
        Try
            Shell(Application.StartupPath & "\更新程序.exe", AppWinStyle.NormalFocus, False)
            End
        Catch ex As Exception

        End Try

    End Sub
    Private Sub Label12_Click(sender As Object, e As EventArgs) Handles Label12.Click
        srd = True
        Panel3.Visible = False
    End Sub
#End Region
#Region "<----加载数据>"
    Sub Cr()
        '加载配置-数据源（初始化）
        Dim writer As New StreamReader(ljc_s, Encoding.GetEncoding("gb2312"))
        Dim st As String
        st = writer.ReadLine()
        lj1 = Application.StartupPath & "\数据\" & st & ".txt“ : Form10.ComboBox1.Text = st
        st = writer.ReadLine()
        lj2 = Application.StartupPath & "\数据\" & st & ".txt“ : Form10.ComboBox3.Text = st
        st = writer.ReadLine()
        writer.Close()
    End Sub '数据源
    Sub Copr()
        '加载数据-不透明度（初始化）
        Dim op As Single
        Dim re1 As New StreamReader(ljc(0), Encoding.GetEncoding("gb2312"))
        Dim st As String
        st = re1.ReadLine()
        re1.Close()
        op = st / 100
        Form2.Opacity = op
        Form3.Opacity = op
        Form4.Opacity = op
        Form5.Opacity = op
        Form7.Opacity = op
        Form10.Opacity = op
        Form10.Label2.Text = st
        Form10.TrackBar1.Value = op * 100
        Form6.Opacity = op
    End Sub '不透明度
    Sub Mojz()
        If File.Exists(ljc(13)) Then
            Call pdq()
            Dim s As String
            Dim re1 As New StreamReader(ljc(13), Encoding.GetEncoding("gb2312"))
            s = re1.ReadLine()
            If s = "1" Then
                ch(4) = 1
                Form10.PictureBox12.Image = Image.FromFile(ljo(ch(4)))
                qd(1) = 1
                Form2.Show()
                Form2.Left = sw * (lp(0) / 100) : Form2.Top = sh * (tp(0) / 100)
            End If
            s = re1.ReadLine()
            If s = "1" Then
                ch(5) = 1
                Form10.PictureBox13.Image = Image.FromFile(ljo(ch(5)))
                qd(2) = 1
                Form3.Show()
                Form3.Left = sw * (lp(1) / 100) : Form3.Top = sh * (tp(1) / 100)
            End If
            s = re1.ReadLine()
            If s = "1" Then
                ch(6) = 1
                Form10.PictureBox14.Image = Image.FromFile(ljo(ch(6)))
                qd(3) = 1
                Form4.Show()
                Form4.Left = sw * (lp(2) / 100) : Form4.Top = sh * (tp(2) / 100)
            End If
            s = re1.ReadLine()
            If s = "1" Then
                ch(22) = 1
                Form10.PictureBox23.Image = Image.FromFile(Ljo(ch(22)))
                qd(4) = 1
                Form5.Show()
                Form5.Left = sw * (lp(3) / 100) : Form5.Top = sh * (tp(3) / 100)
            End If
            s = re1.ReadLine()
            If s = "1" Then
                ch(23) = 1
                Form10.PictureBox24.Image = Image.FromFile(Ljo(ch(23)))
                qd(8) = 1
                Form6.Show()
                Form6.Left = sw * (lp(5) / 100) : Form6.Top = sh * (tp(5) / 100)
            End If
            re1.Close()
        End If
    End Sub '默认启动项
    Sub CJz()
        ljc(0) = Application.StartupPath & "\配置\设置-不透明度.txt"
        ljc(1) = Application.StartupPath & "\配置\设置-开机启动.txt"
        ljc(2) = Application.StartupPath & "\配置\设置-窗口贴靠.txt"
        ljc(3) = Application.StartupPath & "\配置\设置-窗口缩放.txt"
        ljc(4) = Application.StartupPath & "\配置\设置-上课模式.txt"
        ljc(5) = Application.StartupPath & "\配置\设置-上课模式子项.txt"
        ljc(6) = Application.StartupPath & "\配置\设置-定时置顶.txt"
        ljc(7) = Application.StartupPath & "\配置\设置-值日生表开关.txt"
        ljc(8) = Application.StartupPath & "\配置\设置-点名器.txt"
        ljc(9) = Application.StartupPath & "\配置\设置-点名器路径.txt"
        ljc(10) = Application.StartupPath & "\配置\设置-点名器设置.txt"
        ljc(11) = Application.StartupPath & "\配置\设置-上课重排.txt"
        ljc(12) = Application.StartupPath & "\配置\设置-变色模式.txt"
        ljc(13) = Application.StartupPath & "\配置\设置-默认启动项.txt"
        ljc(14) = Application.StartupPath & "\配置\设置-临时时间表.txt"
        ljc(15) = Application.StartupPath & "\配置\设置-临时课表.txt"
        ljc(16) = Application.StartupPath & "\配置\设置-点名器时间.txt"
        ljc(17) = Application.StartupPath & "\配置\设置-窗口缩放.txt"
        ljc(18) = Application.StartupPath & "\配置\设置-倒计时设置.txt"
        ljc(19) = Application.StartupPath & "\配置\设置-进度条剩余时间.txt"
        ljc(20) = Application.StartupPath & "\配置\设置-值日生表自动切换.txt"
        ljc(21) = Application.StartupPath & "\配置\设置-ZdQh.txt"
        ljc(22) = Application.StartupPath & "\配置\设置-隐藏进度.txt"
    End Sub '13项配置的路径
    Sub Xian()
        Form4.Label1.Text = iClass(vb_cday, 1)
        Form4.Label2.Text = iClass(vb_cday, 2)
        Form4.Label3.Text = iClass(vb_cday, 3)
        Form4.Label4.Text = iClass(vb_cday, 4)
        Form4.Label5.Text = iClass(vb_cday, 5)
        Form4.Label6.Text = iClass(vb_cday, 6)
        Form4.Label7.Text = iClass(vb_cday, 7)
        Form4.Label8.Text = iClass(vb_cday, 8)
    End Sub '课表窗口
    Sub Zrjz()
        If File.Exists(ljc(7)) Then
            Dim re1 As New StreamReader(ljc(7), Encoding.GetEncoding("gb2312"))
            ch(9) = Val(re1.ReadLine)
            re1.Close()
            Form10.PictureBox16.Image = Image.FromFile(Ljo(ch(9)))
        End If
        If ch(9) = 1 Then
            If File.Exists(ljc_zrs) Then
                Dim re2 As New StreamReader(ljc_zrs, Encoding.GetEncoding("gb2312"))
                Dim s As String = re2.ReadLine
                re2.Close()
                lj3 = Application.StartupPath & "\数据\" & s & ".txt"
                ljb = Application.StartupPath & "\数据\" & "值日班长" & Mid(s, 5, Len(s) - 4) & ".txt"
                If File.Exists(lj3) Then
                    Call Zrdq()
                    Form10.ComboBox4.Text = s
                    Form10.ComboBox8.Enabled = True
                Else
                    ch(9) = 0
                End If
            Else
                ch(9) = 0
                Form10.PictureBox16.Image = Image.FromFile(Ljo(ch(2)))
            End If
        End If
        If Not File.Exists(lj3) Then
            Form10.PictureBox16.Image = Image.FromFile(Ljo(2))
        End If
    End Sub '值日生表初始化
    Sub Cksf()
        Call SetTag(Form2)
        Call SetTag(Form3)
        Call SetTag(Form4)
        Call SetTag(Form5)
        Call SetTag(Form6)
    End Sub '窗口缩放初始化
    Sub Cojz()
        Application.DoEvents()
        '开机启动
        Dim a As RegistryKey = My.Computer.Registry.CurrentUser.OpenSubKey("Software\Microsoft\Windows\CurrentVersion\Run")
        If a.GetValue("课表", "null").ToString <> "null" Then '存在数据
            ch(0) = 1
            Form10.PictureBox1.Image = Image.FromFile(Ljo(ch(0)))
        End If
        '窗口贴靠
        If File.Exists(ljc(2)) Then
            Dim re1 As New StreamReader(ljc(2), Encoding.GetEncoding("gb2312"))
            ch(1) = Val(re1.ReadLine)
            oftk = (ch(1) = 1)
            re1.Close()
            Form10.PictureBox9.Image = Image.FromFile(Ljo(ch(1)))
        End If
        '窗口置顶
        If File.Exists(ljc(6)) Then
            Dim re2 As New StreamReader(ljc(6), Encoding.GetEncoding("gb2312"))
            ch(3) = Val(re2.ReadLine)
            ofzd = (ch(3) = 1)
            re2.Close()
            Form10.PictureBox11.Image = Image.FromFile(Ljo(ch(3)))
        End If
        '变色模式
        If File.Exists(ljc(12)) Then
            Dim re6 As New StreamReader(ljc(12), Encoding.GetEncoding("gb2312"))
            ofbs = re6.ReadLine()
            re6.Close()
            If ofbs = 0 Then
                Form3.PictureBox2.BackColor = Color.Tomato
                Form10.PictureBox21.Image = Image.FromFile(Ljo(3))
            ElseIf ofbs = 1 Then
                Form3.PictureBox2.BackColor = Color.DeepSkyBlue
                ch(12) = 1
                Form10.PictureBox20.Image = Image.FromFile(Ljo(ch(12)))
            ElseIf ofbs = 2 Then
                Form3.PictureBox2.BackColor = Color.DeepSkyBlue
                ch(12) = 1 : ch(13) = 1
                Form10.PictureBox20.Image = Image.FromFile(Ljo(ch(12)))
                Form10.PictureBox21.Image = Image.FromFile(Ljo(ch(13)))
            End If
        End If
        '上课重排
        If File.Exists(ljc(11)) Then
            Dim re7 As New StreamReader(ljc(11), Encoding.GetEncoding("gb2312"))
            ch(11) = Val(re7.ReadLine)
            re7.Close()
            ofcp = (ch(11) = 1)
            Form10.PictureBox19.Image = Image.FromFile(Ljo(ch(11)))
        End If
        '倒计时label
        For Each tb As Label In Form6.Controls
            djslb(tb.Tag) = tb
        Next
#Region "<点名器>"
        Call Tjz()
        If File.Exists(ljc(8)) Then
            Dim re5 As New StreamReader(ljc(8), System.Text.Encoding.GetEncoding("gb2312"))
            ch(10) = Val(re5.ReadLine)
            ofdm = (ch(10) = 1)
            re5.Close()
        End If
        For Each ce As CheckBox In Form10.Panel11.Controls
            Dim k As Integer
            k = Val(ce.Tag)
            i = (k - 1) \ 6 + 1
            j = (k - 1) Mod 6 + 1
            dmc(i, j) = ce
        Next
        For i = 1 To 7
            For j = 1 To 6
                dmb(i, j) = dmc(i, j).Checked
            Next
        Next
        If File.Exists(ljc(10)) Then
            Dim re4 As New StreamReader(ljc(10), System.Text.Encoding.GetEncoding("gb2312"))
            Dim s As String
            Do Until re4.EndOfStream
                s = re4.ReadLine
                If Mid(s, 1, 1) = "#" Then
                    i = Val(Mid(s, 2, 1))
                    j = 0
                Else
                    j += 1
                    dmc(i, j).Checked = (s = "True")
                    dmb(i, j) = dmc(i, j).Checked
                End If
            Loop
            re4.Close()
        End If '文件加载
        If File.Exists(ljc(9)) Then
            Dim re3 As New StreamReader(ljc(9), System.Text.Encoding.GetEncoding("gb2312"))
            ljto_dm = re3.ReadLine
            re3.Close()
        End If
        For Each tc As TextBox In Form10.Panel13.Controls
            i = tc.Tag \ 2 + 1
            j = tc.Tag Mod 2 + 1
            dmte(i, j) = tc
        Next
        i = 1 : j = 0
        If File.Exists(ljc(16)) Then
            Dim re8 As New StreamReader(ljc(16), System.Text.Encoding.GetEncoding("gb2312"))
            Do Until re8.EndOfStream
                Dim s As Integer = Val(re8.ReadLine)
                j += 1
                If j = 3 Then j = 1 : i += 1
                If j = 1 Then hie(i) = s Else mie(i) = s
            Loop
            re8.Close()
        End If
        For i = 1 To 6
            dmte(i, 1).Text = hie(i)
            dmte(i, 2).Text = mie(i)
        Next
        If File.Exists(ljto_dm) Then Form10.PictureBox18.Image = Image.FromFile(Ljo(ch(10))) Else Form10.PictureBox18.Image = Image.FromFile(Ljo(2))
#End Region
        '帮助界面初始化加载
        If File.Exists(ljh) Then
            Dim re2 As New StreamReader(ljh, Encoding.GetEncoding("gb2312"))
            i = -1
            Dim s As String
            Do Until re2.EndOfStream
                s = re2.ReadLine()
                If Mid(s, 1, 1) = "#" Then
                    i += 1
                Else
                    If hsr(i) = "" Then hsr(i) = s Else hsr(i) += vbCrLf & s
                End If
            Loop
            re2.Close()
        End If
        htd = 1
        Form10.RichTextBox1.Text = hsr(htd)
        Application.DoEvents()
        '倒计时
        If File.Exists(ljc(18)) Then
            Dim re10 As New StreamReader(ljc(18), Encoding.GetEncoding("gb2312"))
            i = 0
            Do Until re10.EndOfStream
                i += 1
                Dim ts() As String = re10.ReadLine.Split(";")
                djstime(i) = ts(0) : djstitle(i) = ts(1)
            Loop
            If i <> 0 Then
                Form6.Height = Form6.Label1.Height * i
            Else
                Form6.Height = Form6.Label1.Height
            End If
            re10.Close()
        Else
            Form6.Height = Form6.Label1.Height
        End If
        '进度条剩余进度
        If File.Exists(ljc(19)) Then
            Dim re11 As New StreamReader(ljc(19), Encoding.GetEncoding("gb2312"))
            ch(21) = Val(re11.ReadLine)
            Form10.PictureBox22.Image = Image.FromFile(Ljo(ch(21)))
            re11.Close()
        End If
        '值日生表自动切换
        If File.Exists(ljc(21)) Then
            If ch(9) = 1 Then
                Dim re12 As New StreamReader(ljc(21))
                ch(24) = re12.ReadLine
                Form10.PictureBox25.Image = Image.FromFile(Ljo(ch(24)))
                ch(25) = re12.ReadLine
                Form10.PictureBox26.Image = Image.FromFile(Ljo(ch(25)))
                re12.Close()
                Call Zd_Read()
                If File.Exists(ljc(20)) And ch(24) = 1 Then
                    Call Zd_Show()
                ElseIf Not File.Exists(ljc(20)) Then
                    ch(24) = 0
                    Form10.PictureBox25.Image = Image.FromFile(Ljo(2))
                End If
            End If
        End If
        '窗口缩放
        xssfold = 100
        Call Cksf()
        If File.Exists(ljc(17)) Then
            Dim re9 As New StreamReader(ljc(17), Encoding.GetEncoding("gb2312"))
            xssfold = Val(re9.ReadLine) : xssf = xssfold
            re9.Close()
            If xssfold <> 100 Then
                Dim tb As Single = xssfold / 100
                Dim d0 As Integer
                Form10.Label46.Text = xssfold
                Call SetControls(tb, Form2)
                Call SetControls(tb, Form3)
                If Form4.Height = hio Then d0 = 0 Else d0 = 1
                Call SetControls(tb, Form4)
                hio = Form4.Label1.Top
                hin = Form4.Label8.Top + Form4.Label8.Height
                If d0 = 0 Then Form4.Height = hio Else Form4.Height = hin
                Call SetControls(tb, Form5)
                ho = Form5.Label1.Top
                hn = Form5.Height
                Call SetControls(tb, Form6)
                Form10.Label13.Text = CStr(xssf)
                Form10.Label46.Text = CStr(xssfold)
                Form10.TrackBar2.Value = xssf / 5
            End If
        Else
            xssf = 100
            Form10.Label13.Text = CStr(xssf)
            Form10.Label46.Text = CStr(xssfold)
            Form10.TrackBar2.Value = xssf / 5
        End If
        Application.DoEvents()
    End Sub '加载可能需要的配置
#End Region
    Function Ljo(ByVal x As Integer) As String
        If x = 0 Then
            Ljo = Application.StartupPath & "\内部文件\图4\off.bmp"
        ElseIf x = 1 Then
            Ljo = Application.StartupPath & "\内部文件\图4\on.bmp"
        Else
            Ljo = Application.StartupPath & "\内部文件\图4\off-f.bmp"
        End If
    End Function
#End Region
#End Region
End Class