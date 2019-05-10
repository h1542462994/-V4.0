Imports System.IO
Imports System.Text
Imports System.Globalization
Module Module1
#Region "<定义>"
    '保留变量
    Public i, j, k As Integer
    '数据
    Public tm(0 To 4, 0 To 50, 0 To 4) As String
    Public ti(0 To 4, 0 To 50) As Integer
    Public iClass(0 To 7, 0 To 9) As String '三项表示数据
    Public vbold As Integer
    Public vbday As Integer '实际的星期
    Public vb_cday As Integer '课表的星期
    Public vb_sday As Integer '时间表的星期
    Public bt As Single '进度条进度
    'Public be As Boolean
    '设置
    Public lp(0 To 5), tp(0 To 5) As Single
    Public ofzr As Boolean = False '是否启用值日生表
    Public ofdm As Boolean = False '是否打开点名器
    Public ofcp As Boolean = False '是否上课重排
    Public ofbs As Integer  '变色模式(0不变色，1骤变色，2平滑变色)
    Public ofzd As Boolean = False '定时置顶(5s)
    Public oftk As Boolean = False  '是否贴靠
    Public op As Single '不透明度
    Public ch(26) As Integer 'form10
    Public xssf As Integer '窗体缩放
    Public xssfold As Integer
    '窗体的公共变量
    Public sw As Integer
    Public sh As Integer
    Public tick As Integer '课表的tick
    Public hio, hin As Single '课表
    Public zd As Boolean '课表
    Public zrtick As Integer '值日生表的tick
    Public ho, hn As Single '值日生表
    Public zrd As Boolean '值日生表
    Public qd(0 To 8) As Integer '是否为选中状态(主窗体)
    Public zes As Boolean = False
    '点名器
    Public dmc(7, 6) As CheckBox
    Public hie(6), mie(6) As Integer '子模块-点名器时间设置
    Public dmt As Integer
    Public dmte(6, 2) As TextBox
    Public dmb(7, 6) As Boolean
    '帮助
    Public hsr(3) As String
    Public htd As Integer = 1
    Public hlb(3) As Label
    '倒计时
    Public djstitle(10) As String
    Public djstime(10) As String
    Public djslb(10) As Label
    '值日生表自动切换
    Public Woy As Integer
    Public CulInfo As New CultureInfo("zh-CN")
    Dim cal As Calendar = CulInfo.Calendar
    Public zdrule(7, 1) As Integer
    Public bbh As String = "4.3.1732.0"
#End Region
#Region "定义，初始化路径"
    Public lj1 As String '时间表
    Public lj2 As String '课表
    Public lj3 As String '值日生表
    Public ljb As String '值日班长
    Public lj4 As String '临时课表
    Public lj5 As String '临时时间表
    Public ljxj As String = Application.StartupPath & "\内部文件\文件校检.txt" '文件校检机制
    Public ljc_p As String = Application.StartupPath & "\配置\配置.txt" '配置-各个窗口的位置
    Public ljc_s As String = Application.StartupPath & "\配置\" & "配置数据源.txt" '配置-表的数据源
    Public ljc_zrs As String = Application.StartupPath & "\配置\" & "配置值日数据源.txt" '配置-表的数据源
    Public ljc(22) As String '配置-19项
    Public ljto_dm As String
    Public ljh As String = Application.StartupPath & "\内部文件\" & "帮助.txt"
    Public ljrule As String
#End Region
#Region "新加内容（值日生表）"
    Public zr_i As Integer
    Public zr_j(20) As Integer
    Public zr_b(20) As String
    Public zr_(20, 20, 2) As String
    Public zr_t As Integer
    Public zr_c(30) As String
    Public zr_ci As Integer
    Dim b_zrb As Boolean '判断是否需要选择值日班长
#End Region
#Region "<数据加载>"
    Public Sub Dq()
#Region "<初始化>"
        For i = 1 To 4
            For j = 1 To 50
                For k = 1 To 4
                    tm(i, j, k) = ""
                Next
            Next
        Next
#End Region
#Region "时间表加载"
        Dim writer1 As New StreamReader(lj1, Encoding.GetEncoding("gb2312")) '文字编码格式，非常重要，否则会出现乱码。
        Dim s As String
        i = 0
        s = writer1.ReadLine()
        Do Until IsNothing(s)
            If Mid(s, 1, 2) = "vb" Then
                i += 1
                j = 0
            Else
                j += 1
                k = 1
                For e = 1 To Len(s)
                    If Mid(s, e, 1) = ";" Then
                        k += 1
                    Else
                        tm(i, j, k) = tm(i, j, k) & Mid(s, e, 1)
                    End If
                Next
            End If
            s = writer1.ReadLine
        Loop
        writer1.Close()
        For i = 1 To 4
            For j = 1 To 50
                ti(i, j) = Val(tm(i, j, 1) & tm(i, j, 2))
            Next
        Next
#End Region
    End Sub
    Public Sub Dq1()
#Region "课表加载"
        Dim writer2 As New StreamReader(lj2, Encoding.GetEncoding("gb2312")) '文字编码格式，非常重要，否则会出现乱码。
        Dim s As String
        i = 0
        s = writer2.ReadLine()
        Do Until IsNothing(s)
            If Mid(s, 1, 2) = "vb" Then
                i += 1
                j = 0
            Else
                j += 1
                iClass(i, j) = s
            End If
            s = writer2.ReadLine()
        Loop
        writer2.Close()
#End Region
    End Sub
    '值日生表加载
    Public Sub Zrdq()
        b_zrb = False
#Region "值日生表加载"
        Dim re1 As New StreamReader(lj3, Encoding.GetEncoding("gb2312")) '文字编码格式，非常重要，否则会出现乱码。
        Dim s As String
        zr_i = 0
        For i = 1 To 20
            For j = 1 To 20
                For k = 1 To 2
                    zr_(i, j, k) = ""
                Next
            Next
        Next
        s = re1.ReadLine()
        Do Until IsNothing(s)
            If Mid(s, 1, 7) = "#define" Then b_zrb = True
            If Mid(s, 1, 2) = "vb" Then
                zr_i += 1
                zr_b(zr_i) = s
                zr_j(zr_i) = 0
            Else
                zr_j(zr_i) += 1
                k = 1
                For e = 1 To Len(s)
                    If Mid(s, e, 1) = ";" Then
                        k += 1
                    Else
                        zr_(zr_i, zr_j(zr_i), k) &= Mid(s, e, 1)
                    End If
                Next
            End If
            s = re1.ReadLine()
        Loop
        re1.Close()
        Form10.ComboBox8.Items.Clear()
        For i = 1 To zr_i
            Form10.ComboBox8.Items.Add(Mid(zr_b(i), 3, Len(zr_b(i)) - 2))
        Next
        Form10.ComboBox8.Text = Mid(zr_b(1), 3, Len(zr_b(1)) - 2)
        zr_t = 1
        Form10.ComboBox9.Items.Clear()
        Form10.ComboBox9.Enabled = False
        If b_zrb = True Then '需要选择值日班长
            If Not (File.Exists(ljb)) Then
                Form10.ComboBox9.Text = "找不到相应的文件"
                Exit Sub
            End If
            Form10.ComboBox9.Enabled = True
#End Region
#Region "值日班长加载"
            Dim re2 As New StreamReader(ljb, Encoding.GetEncoding("gb2312")) '文字编码格式，非常重要，否则会出现乱码。
            'Dim s As String
            i = 0
            zr_ci = 1
            s = re2.ReadLine()
            Do Until IsNothing(s)
                i += 1
                zr_c(i) = s
                Form10.ComboBox9.Items.Add(zr_c(i))
                s = re2.ReadLine
            Loop
            re2.Close()
            Form10.ComboBox9.Text = zr_c(1)
            Form10.Label24.ForeColor = Color.Black
        Else
            Form10.ComboBox9.Text = "没有设置"
            Form10.Label24.ForeColor = Color.Gray
        End If
#End Region
    End Sub
    '加载配置-窗体布局
    Function Ljo(x As Integer) As String
        If x = 0 Then
            Ljo = Application.StartupPath & "\内部文件\图4\off.bmp"
        ElseIf x = 1 Then
            Ljo = Application.StartupPath & "\内部文件\图4\on.bmp"
        ElseIf x = 2 Then
            Ljo = Application.StartupPath & "\内部文件\图4\off-f.bmp"
        Else
            Ljo = Application.StartupPath & "\内部文件\图4\on-f.bmp"
        End If
    End Function
    Public Sub Pdq()
        If File.Exists(ljc_p) Then
            Dim re1 As New StreamReader(ljc_p, System.Text.Encoding.GetEncoding("gb2312")) '文字编码格式，非常重要，否则会出现乱码。
            Dim s As String
            Dim st As String
            k = -1
            s = re1.ReadLine()
            Do Until IsNothing(s)
                st = ""
                k += 1
                For i = 1 To Len(s)
                    If Mid(s, i, 1) <> ";" Then
                        st = st & Mid(s, i, 1)
                    Else
                        lp(k) = Val(st)
                        st = ""
                    End If
                Next
                tp(k) = Val(st)
                s = re1.ReadLine()
            Loop
            re1.Close()
        Else
            lp(0) = 84.89 : tp(0) = 4.8
            lp(1) = 2.65 : tp(1) = 3.75
            lp(2) = 93.54 : tp(2) = 14.13
            lp(3) = 2.7 : tp(3) = 16.44
            lp(4) = 55.15 : tp(4) = 19.9
        End If
    End Sub
    Public Sub Skcp()
        Form2.Left = sw * (lp(0) / 100) : Form2.Top = sh * (tp(0) / 100)
        Form3.Left = sw * (lp(1) / 100) : Form3.Top = sh * (tp(1) / 100)
        Form4.Left = sw * (lp(2) / 100) : Form4.Top = sh * (tp(2) / 100)
        Form5.Left = sw * (lp(3) / 100) : Form5.Top = sh * (tp(3) / 100)
    End Sub
    Public Sub Zrp()
        zr_t = Form10.ComboBox8.SelectedIndex + 1
        Form5.Height = Form5.PictureBox1.Height + (Form5.Label1.Height - 2) * zr_j(zr_t)
        If zr_t = 0 Then Form5.Height = Form5.PictureBox1.Height - 2
        hn = Form5.Height
    End Sub
    Public Sub Opc()
        Form2.Opacity = op
        Form3.Opacity = op
        Form4.Opacity = op
        Form5.Opacity = op
        Form10.Opacity = op
        Form6.Opacity = op
    End Sub
    Public Sub Dms()
        If File.Exists(ljto_dm) Then
            dmt = 0 : 打开点名器.Timer1.Enabled = True
            打开点名器.Show()
        End If
    End Sub
    Public Sub Tjz()
        hie(1) = 6 : mie(1) = 50
        hie(2) = 10 : mie(2) = 18
        hie(3) = 12 : mie(3) = 45
        hie(4) = 15 : mie(4) = 35
        hie(5) = 17 : mie(5) = 40
        hie(6) = 20 : mie(6) = 28
    End Sub
    Public Sub Zd_Read()
        If File.Exists(ljc(20)) Then
            Dim re1 As New StreamReader(ljc(20), Encoding.GetEncoding("gb2312"))
            Dim se As String = re1.ReadLine
            ljrule = Application.StartupPath & "\数据\" & se & ".txt"
            Form10.ComboBox12.Text = se
            If File.Exists(ljrule) Then
                Dim re2 As New StreamReader(ljrule, Encoding.GetEncoding("gb2312"))
                se = re2.ReadLine
                If se = "#zdrule=1" Then
                    Dim est0() As String = re2.ReadLine.Split(";")
                    Dim est1() As String = re2.ReadLine.Split(";")
                    For i = 1 To 7
                        zdrule(i, 0) = Val(est0(i - 1))
                        zdrule(i, 1) = Val(est1(i - 1))
                    Next
                Else
                End If
                re2.Close()
                Form10.PictureBox25.Image = Image.FromFile(Ljo(ch(24)))
            Else
                ch(24) = 0
                Form10.PictureBox25.Image = Image.FromFile(ljo(2))
            End If
            re1.Close()
        End If
    End Sub
    Public Sub Zd_Show()
        Woy = cal.GetWeekOfYear(Now, CalendarWeekRule.FirstDay, DayOfWeek.Sunday) Mod 2
        If ch(25) = 1 Then
            Woy = 1 - Woy
        End If
        zr_t = zdrule(vb_cday, Woy)
        Form10.ComboBox8.SelectedIndex = zr_t - 1
        Call Zrp()
    End Sub
#End Region
#Region "<新-窗体缩放>"
    Public Sub SetTag(ByVal obj As Object)
        For Each con As Control In obj.Controls
            con.Tag = con.Width & ":" & con.Height & ":" & con.Left & ":" & con.Top & ":" & con.Font.Size
            If con.Controls.Count > 0 Then
                SetTag(con)
            End If
        Next
        obj.Tag = obj.Width & ":" & obj.Height & ":" & obj.Left & ":" & obj.Top & ":" & obj.Font.Size
    End Sub
    Public Sub SetControls(newb As Single, obj As Object)
        For Each con As Control In obj.Controls
            con.AutoSize = False
            Dim mytag() As String = con.Tag.ToString.Split(":")
            con.Width = mytag(0) * newb
            con.Height = mytag(1) * newb
            con.Left = mytag(2) * newb
            con.Top = mytag(3) * newb
            ' 计算字体缩放比例， 缩放字体
            Dim currentSize As Single = (mytag(1) * newb * mytag(4)) / mytag(1)
            con.Font = New Font(con.Font.Name, currentSize,
            con.Font.Style, con.Font.Unit)
            '如果是容器控件， 则递归继续缩放
            If con.Controls.Count > 0 Then
                SetControls(newb, con)
            End If
        Next
        Dim mytag1() As String = obj.Tag.ToString.Split(":")
        obj.Width = mytag1(0) * newb
        obj.Height = mytag1(1) * newb
    End Sub
#End Region
End Module