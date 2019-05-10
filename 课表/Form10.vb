Imports System.IO
Imports System.Text
Imports Microsoft.Win32
Public Class Form10
#Region "<定义>"
    Dim pb(4) As PictureBox
    Dim i, j, k As Integer
    Dim qrd(4) As Integer
    Dim mi(4) As Integer
    Dim mold(4) As Integer
    Dim mo As Boolean '判断窗体是否在里面
    Dim mix, miy As Single
    Dim l(3) As Single
    Dim imd As Boolean
    Dim bx, by As Single
    Dim pab(4) As Panel
    Dim ljt As String
    Dim zp As Integer
    Dim tb(8) As TextBox
    Dim djstd As Integer = 0
    Dim es(4) As String
    Dim est(3) As Integer
#End Region
#Region "<窗体>"
#Region "<窗体加载>"
    Private Sub Form10_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        '窗口缩放
        'BackColor = Color.LightSteelBlue
        'TrackBar2.Value = 10
        xssf = TrackBar2.Value * 5
        Label13.Text = CStr(xssf)
        xssfold = xssf
        Width = Label55.Left + Label55.Width : Height = Label55.Top + Label55.Height  '窗体大小
        tb(0) = TextBox1 : tb(1) = TextBox2 : tb(2) = TextBox3 : tb(3) = TextBox4 : tb(4) = TextBox5 : tb(5) = TextBox6 : tb(6) = TextBox7 : tb(7) = TextBox8 : tb(8) = TextBox9
        For i = 0 To 4
            qrd(i) = 0
            mold(i) = 0
        Next
        qrd(0) = 1
        With Panel1
            l(0) = Me.Left
            l(1) = Me.Left + .Width
            l(2) = Me.Top + .Top
            l(3) = Me.Top + .Top + .Height
        End With
        PictureBox8.Image = Image.FromFile(Application.StartupPath & "\内部文件\图1\71.bmp")
        Call Csh()
        For i = 0 To 4
            pb(i).Image = Image.FromFile(Lj(i, qrd(i), mi(i)))
        Next
        mo = True
        Call Szcsh()
        Label61.Text = "课表版本号 " & bbh & " 专利"
        Label68.Text = Label61.Text : Label69.Text = Label61.Text : Label70.Text = Label61.Text
    End Sub
#End Region
#Region "<窗体设计>"
    Private Sub M_mousemove(sender As Object, e As MouseEventArgs)
        For i = 0 To 4
            If i = Val(sender.tag) Then
                mi(i) = 1
                If mo = False Then
                    mo = True
                    pb(i).Image = Image.FromFile(Lj(i, qrd(i), mi(i)))
                End If
            Else
                mi(i) = 0
            End If
            If mi(i) <> mold(i) Then pb(i).Image = Image.FromFile(Lj(i, qrd(i), mi(i)))
        Next
        For i = 0 To 4
            mold(i) = mi(i)
        Next

    End Sub
    Private Sub M_mouseclick(sender As Object, e As MouseEventArgs)
        For i = 0 To 4
            If i = Val(sender.Tag) Then
                qrd(i) = 1
                pab(i).BringToFront()
            Else
                qrd(i) = 0
            End If
            pb(i).Image = Image.FromFile(Lj(i, qrd(i), mi(i)))
        Next
    End Sub
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        mix = Cursor.Position.X
        miy = Cursor.Position.Y
        If mix < l(0) Or mix > l(1) Or miy < l(2) Or miy > l(3) Then
            For i = 0 To 4
                mi(i) = 0
                If mi(i) <> mold(i) Then pb(i).Image = Image.FromFile(Lj(i, qrd(i), mi(i)))
            Next
            mo = False
        End If
    End Sub
    Private Sub PictureBox8_Click(sender As Object, e As EventArgs) Handles PictureBox8.Click
        Me.Hide()
        qd(5) = 0
        zes = True
    End Sub
#End Region
#Region "<窗体移动>"
    Private Sub Me_MouseDown(sender As Object, e As MouseEventArgs) Handles PictureBox7.MouseDown, Panel2.MouseDown, Panel3.MouseDown, Panel4.MouseDown, Panel5.MouseDown, Panel6.MouseDown, MyBase.MouseDown
        imd = True
        bx = e.X : by = e.Y
    End Sub
    Private Sub Me_MouseMove(sender As Object, e As MouseEventArgs) Handles PictureBox7.MouseMove, Panel2.MouseMove, Panel3.MouseMove, Panel4.MouseMove, Panel5.MouseMove, Panel6.MouseMove, MyBase.MouseMove
        If imd Then
            Left = Left + e.X - bx
            Top = Top + e.Y - by
        End If
    End Sub
    Private Sub Me_MouseUp(sender As Object, e As MouseEventArgs) Handles PictureBox7.MouseUp, Panel2.MouseUp, Panel3.MouseUp, Panel4.MouseUp, Panel5.MouseUp, Panel6.MouseUp, MyBase.MouseUp
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
#End Region
#End Region
#Region "<函数、过程>"
    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick
        With Panel1
            l(0) = Me.Left
            l(1) = Me.Left + .Width
            l(2) = Me.Top + .Top
            l(3) = Me.Top + .Top + .Height
        End With
    End Sub
    Function Lj(x As Integer, y As Integer, z As Integer) As String
        Lj = Application.StartupPath & "\内部文件\图4\" & x & (2 * y + z) & ".bmp"
    End Function
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
    Function G(x As String) As Boolean
        '判断文件(临时时间表)是否符合要求
        G = True
        Dim re1 As New StreamReader(x, Encoding.GetEncoding("gb2312")) '文字编码格式，非常重要，否则会出现乱码。
        Dim s As String
        Dim st As String
        i = 0
        s = re1.ReadLine() '-----
        Do Until IsNothing(s) '-----
            i += 1
            If i > 50 Then G = False : Exit Function
            j = 1
            st = ""
            For k = 1 To Len(s)
                If Mid(s, k, 1) = ";" Then
                    j += 1
                    st = ""
                Else st = st & Mid(s, k, 1)
                    If j <= 2 Then
                        If Mid(s, k, 1) < "0" Or Mid(s, k, 1) > "9" Then
                            G = False : Exit Function
                        End If
                    End If
                End If
            Next
            If j <> 4 And j <> 3 Then G = False : Exit Function
            s = re1.ReadLine() '-----
        Loop '------
        re1.Close() '------
    End Function
    Function F(x As String) As Boolean
        '判断文件(临时课表)是否符合要求
        F = True
        Dim re1 As New StreamReader(x, Encoding.GetEncoding("gb2312"))
        Dim s As String
        i = 0
        Do Until re1.EndOfStream
            i += 1
            s = re1.ReadLine
        Loop
        If i <> 9 Then F = False
        re1.Close()
    End Function
    Sub Csh()
        For Each pbt As PictureBox In Panel1.Controls
            pb(pbt.Tag) = pbt
            AddHandler pbt.MouseMove, AddressOf M_mousemove
            AddHandler pbt.MouseClick, AddressOf M_mouseclick
        Next
        pab(0) = Panel2 : pab(1) = Panel3 : pab(2) = Panel4 : pab(3) = Panel5 : pab(4) = Panel6
        For i = 0 To 4
            pab(i).Left = Label56.Left : pab(i).Top = Label56.Top
            pab(i).Width = Label57.Left + Label57.Width - Label56.Left : pab(i).Height = Label57.Top + Label57.Height - Label56.Top
        Next
        pab(0).BringToFront()
        For Each lbt As Label In Panel16.Controls
            hlb(lbt.Tag) = lbt
            AddHandler lbt.Click, AddressOf HLabel_click
        Next

    End Sub
    Sub Cc()
        Dim msr As New FileStream(ljc_s, FileMode.Create)
        Dim we1 As New StreamWriter(msr, Encoding.GetEncoding("gb2312")) '文字编码格式，非常重要，否则会出现乱码。
        we1.WriteLine(ComboBox1.Text)
        we1.WriteLine(ComboBox3.Text)
        we1.Close()
    End Sub
    Sub Cczr()
        Dim mr1 As New FileStream(ljc_zrs, FileMode.Create)
        Dim we1 As New StreamWriter(mr1, Encoding.GetEncoding("gb2312"))
        we1.WriteLine(ComboBox4.Text)
        we1.Close()
    End Sub
    Sub Szcsh()
        '文件加载
        ljt = Application.StartupPath & "\数据\"
        ComboBox1.Items.Clear()
        If ch(7) = 0 Then ComboBox1.Enabled = True
        ComboBox2.Items.Clear() : ComboBox2.Enabled = True
        ComboBox3.Items.Clear()
        If ch(8) = 0 Then ComboBox3.Enabled = True
        ComboBox4.Items.Clear() : ComboBox4.Enabled = True
        ComboBox5.Items.Clear() : ComboBox5.Enabled = True
        ComboBox12.Items.Clear()
        For Each mfl In Directory.GetFiles(ljt)
            Dim mfn As String = Mid(Path.GetFileName(mfl), 1, Len(Path.GetFileName(mfl)) - 4)
            If Mid(mfn, 1, 2) = "时间" Then
                ComboBox1.Items.Add(mfn.ToString)
            ElseIf Mid(mfn, 1, 2) = "课表" Then
                ComboBox3.Items.Add(mfn.ToString)
                If lj5 = "" Then lj5 = Application.StartupPath & "\数据\" & ComboBox3.Text & ".txt"
            ElseIf Mid(mfn, 1, 2) = "临时" Then
                If G(mfl) = True Then
                    ComboBox2.Items.Add(mfn.ToString)
                End If
                If lj4 = "" Then lj4 = Application.StartupPath & "\数据\" & ComboBox2.Text & ".txt" ： ComboBox2.Text = mfn
            ElseIf Mid(mfn, 1, 2) = "TC" Then
                If F(mfl) = True Then
                    ComboBox5.Items.Add(mfn.ToString)
                End If
                If lj5 = "" Then lj5 = Application.StartupPath & "\数据\" & ComboBox5.Text & ".txt" : ComboBox5.Text = mfn
            ElseIf Mid(mfn, 1, 4) = "值日生表" Then
                ComboBox4.Items.Add(mfn.ToString)
                If lj3 = "" Then
                    lj3 = Application.StartupPath & "\数据\" & ComboBox4.Text & ".txt"
                    ComboBox4.Text = mfn
                    ljb = Application.StartupPath & "\数据\" & "值日班长" & Mid(ComboBox3.Text, 5, Len(ComboBox3.Text) - 4) & ".txt"
                End If
            ElseIf Mid(mfn, 1, 2) = "ZD" Then
                ComboBox12.Items.Add(mfn.ToString)
            End If
        Next
        If ComboBox2.Items.Count = 0 Then
            ComboBox2.Enabled = False
            ComboBox2.Text = "没有此类文件"
            Label15.ForeColor = Color.Gray
            PictureBox15.Image = Image.FromFile(Ljo(2))
        Else
            Label15.ForeColor = Color.Black
            PictureBox15.Image = Image.FromFile(Ljo(ch(7)))
        End If
        If ComboBox5.Items.Count = 0 Then
            ComboBox5.Enabled = False
            ComboBox5.Text = "没有此类文件"
            Label19.ForeColor = Color.Gray
            PictureBox17.Image = Image.FromFile(Ljo(2))
        Else
            Label19.ForeColor = Color.Black
            PictureBox17.Image = Image.FromFile(Ljo(ch(8)))
        End If
        If ComboBox4.Items.Count = 0 Then
            ComboBox4.Enabled = False
            ComboBox4.Text = "没有此类文件"
            Label18.ForeColor = Color.Gray
            PictureBox16.Image = Image.FromFile(Ljo(2))
            Label23.ForeColor = Color.Gray
            Label24.ForeColor = Color.Gray
        Else
            Label18.ForeColor = Color.Black
            PictureBox16.Image = Image.FromFile(Ljo(ch(9)))
            Label23.ForeColor = Color.Black
            Label24.ForeColor = Color.Black
        End If
        If ComboBox12.Items.Count = 0 Then
            ComboBox12.Enabled = False
            ComboBox12.Text = "没有此类文件"
            Label90.ForeColor = Color.Gray
        Else
            ComboBox12.Enabled = True
            Label90.ForeColor = Color.Black
        End If
    End Sub
    Sub Lsjz()
        Dim re1 As New StreamReader(lj4, Encoding.GetEncoding("gb2312")) '文字编码格式，非常重要，否则会出现乱码。
        Dim s As String
        Dim il As Integer
        '清空
        If vb_sday = 1 Then
            il = 1
        ElseIf vb_sday >= 2 And vb_sday <= 5 Then
            il = 2
        ElseIf vb_sday = 6 Then
            il = 3
        Else : il = 4
        End If
        For i = 1 To 50
            For j = 1 To 4
                tm(il, i, j) = ""
            Next
        Next
        i = 0
        s = re1.ReadLine() '-----
        Do Until IsNothing(s) '-----
            i += 1
            j = 1
            For k = 1 To Len(s)
                If Mid(s, k, 1) = ";" Then
                    j += 1
                Else
                    tm(il, i, j) = tm(il, i, j) & Mid(s, k, 1)
                End If
            Next
            s = re1.ReadLine() '-----
        Loop '------
        re1.Close() '------
        For i = 1 To 4
            For j = 1 To 50
                ti(i, j) = Val(tm(i, j, 1) & tm(i, j, 2))
            Next
        Next
    End Sub
    Sub Lcjz()
        Dim re1 As New StreamReader(lj5, Encoding.GetEncoding("gb2312"))
        i = 0
        Do Until re1.EndOfStream
            i += 1
            iClass(vb_cday, i) = re1.ReadLine
        Loop
        re1.Close()
    End Sub
    Sub Djs()
        If djstd = 0 And DjsB(TextBox22.Text) Then
            '新建
            For i = 1 To 10
                If djstime(i) = "" Then Exit For
            Next
            If i <= 10 Then
                djstime(i) = $"{es(1)}/{es(2)}/{es(3)}"
                djstitle(i) = es(4)
            End If
        ElseIf djstd = 1 And DjsB(TextBox22.text) Then
            '修改
            i = ListBox1.SelectedIndex + 1
            djstime(i) = $"{es(1)}/{es(2)}/{es(3)}"
            djstitle(i) = es(4)
        End If
        Dim ei As Integer = ListBox1.SelectedIndex
        ListBox1.Items.Clear()
        For i = 1 To 10
            If djstime(i) = "" Then Exit For
        Next
        For j = 1 To i - 1
            ListBox1.Items.Add(j & " " & djstime(j) & " " & djstitle(j))
        Next
        If djstd = 0 Then
            ListBox1.SelectedIndex = ListBox1.Items.Count - 1
        ElseIf djstd = 1 Then
            ListBox1.SelectedIndex = ei
        End If
    End Sub
    Sub DjsW()
        '写入
        For i = 1 To 10
            If djstime(i) = "" Then Exit For
        Next
        i -= 1
        Dim mr1 As New FileStream(ljc(18), FileMode.Create)
        Dim we1 As New StreamWriter(mr1, Encoding.GetEncoding("gb2312"))
        For j = 1 To i
            we1.WriteLine(djstime(j) & ";" & djstitle(j))
        Next
        we1.Close()
    End Sub
    Function DjsGetNum() As Integer
        For i = 1 To 10
            If djstime(i) = "" Then DjsGetNum = i - 1 : Exit Function
        Next
        DjsGetNum = 10
    End Function
    Function DjsB(x As String) As Boolean
        For i = 1 To 4
            es(i) = ""
        Next
        For i = 1 To 3
            est(i) = 0
        Next
        DjsB = True
        j = 1
        For i = 1 To Len(x)
            If j <= 2 And Mid(x, i, 1) = "/" Then
                j += 1
            ElseIf j = 3 And (Mid(x, i, 1) = ";" Or Mid(x, i, 1) = "；") Then
                j += 1
            Else
                If (Mid(x, i, 1) >= "0" And Mid(x, i, 1) <= "9" And j <= 3) Or j = 4 Then
                    es(j) += Mid(x, i, 1)
                Else
                    DjsB = False : Exit Function
                End If
            End If
        Next
        If Val(es(1)) < 1959 Or Val(es(1)) > 2049 Then DjsB = False : Exit Function
        If Not IsDate(es(1) & "/" & es(2) & "/" & es(3)) Then DjsB = False : Exit Function
    End Function
#End Region
#Region "<项目>"
#Region "<加载>"
#End Region
#Region "<常用>"
    Private Sub TrackBar1_Scroll(sender As Object, e As EventArgs) Handles TrackBar1.Scroll
        op = TrackBar1.Value / 100
        Call Opc()
        Label2.Text = Int(op * 100)
        Dim mr As New FileStream(ljc(0), FileMode.Create)
        Dim we1 As New StreamWriter(mr, Encoding.GetEncoding("gb2312"))
        we1.WriteLine(CStr(Int(op * 100)))
        we1.Close()
    End Sub '不透明度
    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
        '开机启动
        ch(0) = 1 - ch(0)
        PictureBox1.Image = Image.FromFile(Ljo(ch(0)))
        If ch(0) = 1 Then
            Dim a As RegistryKey = My.Computer.Registry.CurrentUser.CreateSubKey("Software\Microsoft\Windows\CurrentVersion\Run")
            a.SetValue("课表", Application.StartupPath & "\课表.exe")
        Else
            Dim a As RegistryKey = My.Computer.Registry.CurrentUser.CreateSubKey("Software\Microsoft\Windows\CurrentVersion\Run")
            a.DeleteValue("课表")
        End If
    End Sub '开机启动
    Private Sub PictureBox9_Click(sender As Object, e As EventArgs) Handles PictureBox9.Click
        '窗口贴靠
        ch(1) = 1 - ch(1)
        PictureBox9.Image = Image.FromFile(Ljo(ch(1)))
        oftk = (ch(1) = 1)
        Dim mr As New FileStream(ljc(2), FileMode.Create)
        Dim we1 As New StreamWriter(mr, Encoding.GetEncoding("gb2312"))
        we1.WriteLine(CStr(ch(1)))
        we1.Close()
    End Sub '窗口贴靠
    Private Sub PictureBox11_Click(sender As Object, e As EventArgs) Handles PictureBox11.Click
        '定时置顶
        ch(3) = 1 - ch(3)
        PictureBox11.Image = Image.FromFile(Ljo(ch(3)))
        ofzd = (ch(3) = 1)
        Dim mr As New FileStream(ljc(6), FileMode.Create)
        Dim we1 As New StreamWriter(mr, Encoding.GetEncoding("gb2312"))
        we1.WriteLine(CStr(ch(3)))
        we1.Close()
    End Sub '定时置顶
    Private Sub PictureBox12_Click(sender As Object, e As EventArgs) Handles PictureBox12.Click, PictureBox13.Click, PictureBox14.Click, PictureBox23.Click, PictureBox24.Click
        '默认启动项
        i = Val(sender.Tag)
        ch(i) = 1 - ch(i)
        PictureBox12.Image = Image.FromFile(Ljo(ch(4)))
        PictureBox13.Image = Image.FromFile(Ljo(ch(5)))
        PictureBox14.Image = Image.FromFile(Ljo(ch(6)))
        PictureBox23.Image = Image.FromFile(Ljo(ch(22)))
        PictureBox24.Image = Image.FromFile(Ljo(ch(23)))
        Dim mr As New FileStream(ljc(13), FileMode.Create)
        Dim we1 As New StreamWriter(mr, Encoding.GetEncoding("gb2312"))
        we1.WriteLine(CStr(ch(4)))
        we1.WriteLine(CStr(ch(5)))
        we1.WriteLine(CStr(ch(6)))
        we1.WriteLine(CStr(ch(22)))
        we1.WriteLine(CStr(ch(23)))
        we1.Close()
    End Sub '默认启动项
    Private Sub TrackBar2_Scroll(sender As Object, e As EventArgs) Handles TrackBar2.Scroll
        '窗口缩放
        xssf = TrackBar2.Value * 5
        Label13.Text = CStr(xssf)
    End Sub   '窗口缩放
    Private Sub Label35_Click(sender As Object, e As EventArgs) Handles Label35.Click
        If xssf <> xssfold Then
            Dim d0 As Integer
            Dim tb As Single = xssf / 100
            xssfold = xssf
            Label46.Text = xssfold
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
            For i = 1 To 10
                If djstime(i) = "" Then Exit For
            Next
            i -= 1
            If i >= 1 Then
                Form6.Height = Form6.Label1.Height * i
            Else
                Form6.Height = Form6.Label1.Height
            End If
            Dim mr As New FileStream(ljc(17), FileMode.Create)
            Dim we1 As New StreamWriter(mr, Encoding.GetEncoding("gb2312"))
            we1.WriteLine(Label13.Text)
            we1.Close()
        End If
    End Sub  '窗口缩放
    Private Sub PictureBox19_Click(sender As Object, e As EventArgs) Handles PictureBox19.Click
        ch(11) = 1 - ch(11)
        PictureBox19.Image = Image.FromFile(Ljo(ch(11)))
        ofcp = (ch(11) = 1)
        Dim mr As New FileStream(ljc(11), FileMode.Create)
        Dim we1 As New StreamWriter(mr, Encoding.GetEncoding("gb2312"))
        we1.WriteLine(ch(11))
        we1.Close()
    End Sub '上课重排
    Private Sub Label32_Click(sender As Object, e As EventArgs) Handles Label32.Click
        zp = 6
        Call Zcd()
    End Sub '帮助界面
#End Region
#Region "<文件加载>"
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        lj1 = Application.StartupPath & "\数据\" & ComboBox1.Text & ".txt"
        Call Dq() : Call Cc()
    End Sub
    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        lj2 = Application.StartupPath & "\数据\" & ComboBox3.Text & ".txt"
        Call Dq1() : Call Cc()
    End Sub
    Private Sub PictureBox15_Click(sender As Object, e As EventArgs) Handles PictureBox15.Click
        If File.Exists(lj4) Then
            ch(7) = 1 - ch(7)
            PictureBox15.Image = Image.FromFile(ljo(ch(7)))
            'tofls = (ch(7) = 1)
            If ch(7) = 1 Then
                If g(lj4) = True Then
                    Call Lsjz()
                    ComboBox1.Enabled = False : ComboBox6.Enabled = False
                Else
                    ch(7) = 0
                    PictureBox15.Image = Image.FromFile(ljo(ch(7)))
                    ComboBox1.Enabled = True : ComboBox6.Enabled = True : ComboBox6.Text = CStr(vb_sday)
                End If
            Else
                Call Dq()
                ComboBox1.Enabled = True : ComboBox6.Enabled = True : ComboBox6.Text = CStr(vb_sday)
            End If
        Else
            PictureBox15.Image = Image.FromFile(Ljo(2))
        End If
    End Sub
    Private Sub PictureBox17_Click(sender As Object, e As EventArgs) Handles PictureBox17.Click
        If File.Exists(lj5) Then
            ch(8) = 1 - ch(8)
            PictureBox17.Image = Image.FromFile(Ljo(ch(8)))
            If ch(8) = 1 Then
                If F(lj5) = True Then
                    Call Lcjz()
                    ComboBox3.Enabled = False : ComboBox7.Enabled = False
                Else
                    ch(8) = 0
                    PictureBox17.Image = Image.FromFile(Ljo(2))
                    ComboBox3.Enabled = True : ComboBox7.Enabled = True : ComboBox7.Text = CStr(vb_cday)
                End If
            Else
                Call Dq1()
                ComboBox3.Enabled = True : ComboBox7.Enabled = True : ComboBox7.Text = CStr(vb_cday)
            End If
        Else
            PictureBox17.Image = Image.FromFile(Ljo(2))
        End If
    End Sub
    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        lj4 = Application.StartupPath & "\数据\" & ComboBox2.Text & ".txt"
        If ch(7) = 1 Then
            If G(lj4) = True Then
                Call Lsjz()
                ComboBox1.Enabled = False
            Else
                ch(7) = 0
                PictureBox15.Image = Image.FromFile(ljo(ch(7)))
                ComboBox1.Enabled = True
            End If
        End If
        If Not (File.Exists(lj4)) Then PictureBox15.Image = Image.FromFile(Ljo(2))
    End Sub
    Private Sub Label17_Click(sender As Object, e As EventArgs) Handles Label17.Click
        zp = 1
        Call Zcd()
    End Sub
    Private Sub Label34_Click(sender As Object, e As EventArgs) Handles Label34.Click
        Call Szcsh()
    End Sub
    Private Sub ComboBox5_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox5.SelectedIndexChanged
        lj5 = Application.StartupPath & "\数据\" & ComboBox5.Text & ".txt"
        If ch(8) = 1 Then
            If F(lj5) = True Then
                Call Lcjz()
                ComboBox3.Enabled = False
            Else
                ch(8) = 0
                PictureBox17.Image = Image.FromFile(Ljo(ch(8)))
                ComboBox3.Enabled = True
            End If
        End If
    End Sub
    '值日生表加载
    Private Sub PictureBox16_Click(sender As Object, e As EventArgs) Handles PictureBox16.Click
        ch(9) = 1 - ch(9)
        PictureBox16.Image = Image.FromFile(Ljo(ch(9)))
        If ch(9) = 1 Then
            ofzr = True
            If File.Exists(lj3) Then
                Call Zrdq()
            Else
                ch(9) = 0
                PictureBox16.Image = Image.FromFile(Ljo(2))
            End If
        Else
        End If
        If ch(9) = 0 Then
            ComboBox8.Enabled = False : ComboBox8.Text = "现在不可用" : Label23.ForeColor = Color.Gray
            ComboBox9.Enabled = False : ComboBox9.Text = "现在不可用" : Label24.ForeColor = Color.Gray
            Form5.Height = Form5.PictureBox1.Height - 2
            hn = Form5.Height
            PictureBox25.Image = Image.FromFile(Ljo(2)) : Label89.ForeColor = Color.Gray
        Else
            ComboBox8.Enabled = True : Label23.ForeColor = Color.Black
            Form5.Height = Form5.PictureBox1.Height + (Form5.Label1.Height - 2) * zr_j(zr_t)
            hn = Form5.Height
            zrd = True : zrtick = 0
            PictureBox25.Image = Image.FromFile(Ljo(ch(24))) : Label89.ForeColor = Color.Black
        End If
        Dim mr1 As New FileStream(ljc(7), FileMode.Create)
        Dim we1 As New StreamWriter(mr1, Encoding.GetEncoding("gb2312"))
        we1.WriteLine(CStr(ch(9)))
        we1.Close()
        Form5.Tag = Form5.Width / xssf * 100 & ":" & Form5.Height / xssf * 100
    End Sub
    Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox4.SelectedIndexChanged
        lj3 = Application.StartupPath & "\数据\" & ComboBox4.Text & ".txt"
        ljb = Application.StartupPath & "\数据\" & "值日班长" & Mid(ComboBox4.Text, 5, Len(ComboBox4.Text) - 4) & ".txt"
        If ch(9) = 1 Then
            Call Zrdq()
            ComboBox8.Items.Clear()
            For i = 1 To zr_i
                ComboBox8.Items.Add(Mid(zr_b(i), 3, Len(zr_b(i)) - 2))
            Next
            ComboBox8.Text = Mid(zr_b(1), 3, Len(zr_b(1)) - 2)
        End If
        Call Cczr()
    End Sub
#End Region
#Region "<星期,值日生>"
    Private Sub ComboBox6_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox6.SelectedIndexChanged
        Call Dq()
        vb_sday = Val(ComboBox6.Text)
    End Sub
    Private Sub ComboBox7_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox7.SelectedIndexChanged
        Call Dq1()
        vb_cday = Val(ComboBox7.Text)
    End Sub
    Private Sub ComboBox8_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox8.SelectedIndexChanged
        zr_t = ComboBox6.SelectedIndex + 1
        Call Module1.Zrp()
    End Sub
    Private Sub ComboBox9_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox9.SelectedIndexChanged
        zr_ci = ComboBox9.SelectedIndex + 1
    End Sub
#End Region
#Region "<外部程序>"
    Private Sub Label26_Click(sender As Object, e As EventArgs) Handles Label26.Click
        zp = 2
        Call Zcd()
    End Sub
    Private Sub Label27_Click(sender As Object, e As EventArgs) Handles Label27.Click
        zp = 3
        Call Zcd()
    End Sub
    Private Sub Label47_Click(sender As Object, e As EventArgs) Handles Label47.Click
        zp = 4
        Call Zcd()
    End Sub
    Private Sub PictureBox18_Click(sender As Object, e As EventArgs) Handles PictureBox18.Click
        ch(10) = 1 - ch(10) : ofdm = (ch(10) = 1)
        If File.Exists(ljto_dm) = False Then
            ch(10) = 0 : ofdm = False
        End If
        If Not (File.Exists(ljto_dm)) Then
            PictureBox18.Image = Image.FromFile(Ljo(2))
            Label25.ForeColor = Color.Gray
        Else
            PictureBox18.Image = Image.FromFile(Ljo(ch(10)))
            Label25.ForeColor = Color.Black
        End If
        Dim mr As New FileStream(ljc(8), FileMode.Create)
        Dim we1 As New StreamWriter(mr, Encoding.GetEncoding("gb2312"))
        we1.WriteLine(CStr(ch(10)))
        we1.Close()
    End Sub '点名器
    Private Sub Label80_Click(sender As Object, e As EventArgs) Handles Label80.Click
        zp = 7
        Call Zcd()
    End Sub
    Private Sub PictureBox25_Click(sender As Object, e As EventArgs) Handles PictureBox25.Click
        If ch(9) = 1 And File.Exists(ljrule) Then
            ch(24) = 1 - ch(24)
            PictureBox25.Image = Image.FromFile(Ljo(ch(24)))
            If ch(24) = 1 Then
                Call Zd_Show()
            End If
        Else
            ch(24) = 0
            PictureBox25.Image = Image.FromFile(Ljo(2))
        End If
        Dim mr As New FileStream(ljc(21), FileMode.Create)
        Dim we1 As New StreamWriter(mr, Encoding.GetEncoding("gb2312"))
        we1.WriteLine(CStr(ch(24)))
        we1.WriteLine(CStr(ch(25)))
        we1.Close()
    End Sub
    Private Sub PictureBox26_Click(sender As Object, e As EventArgs) Handles PictureBox26.Click
        ch(25) = 1 - ch(25)
        If File.Exists(ljc(20)) And ch(24) = 1 Then
            Call Zd_Show()
        End If
        PictureBox26.Image = Image.FromFile(Ljo(ch(25)))
        Dim mr As New FileStream(ljc(21), FileMode.Create)
        Dim we1 As New StreamWriter(mr, Encoding.GetEncoding("gb2312"))
        we1.WriteLine(CStr(ch(24)))
        we1.WriteLine(CStr(ch(25)))
        we1.Close()
    End Sub
    Private Sub ComboBox12_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox12.SelectedIndexChanged
        Dim mr As New FileStream(ljc(20), FileMode.Create)
        Dim we1 As New StreamWriter(mr, Encoding.GetEncoding("gb2312"))
        we1.WriteLine(ComboBox12.Text)
        we1.Close()
        Call Zd_Read()
    End Sub
#End Region
#Region "<辅助功能>"
    Private Sub Label30_Click(sender As Object, e As EventArgs) Handles Label30.Click
        zp = 5
        Call Zcd()
    End Sub '系统时间调整
    Private Sub PictureBox20_Click(sender As Object, e As EventArgs) Handles PictureBox20.Click, PictureBox21.Click
        If sender Is PictureBox20 Then
            ch(12) = 1 - ch(12)
            If ch(12) = 0 Then
                ch(13) = 0 : ofbs = 0
            Else
                ofbs = 1
            End If
        Else
            If ch(12) = 1 Then
                ch(13) = 1 - ch(13)
                ofbs = 1 + ch(13)
            End If
        End If
        PictureBox20.Image = Image.FromFile(Ljo(ch(12)))
        PictureBox21.Image = Image.FromFile(Ljo(ch(13)))
        If ch(12) = 0 Then PictureBox21.Image = Image.FromFile(Ljo(2))
        Dim mr As New FileStream(ljc(12), FileMode.Create)
        Dim we1 As New StreamWriter(mr, Encoding.GetEncoding("gb2312"))
        we1.WriteLine(CStr(ch(12) + ch(13)))
        we1.Close()
    End Sub '进度条变色
    Private Sub PictureBox10_Click(sender As Object, e As EventArgs) Handles PictureBox10.Click
        '上课模式
        'ch(2) = 1 - ch(2)
        PictureBox10.Image = Image.FromFile(Ljo(2))
    End Sub '上课模式
    Private Sub PictureBox22_Click(sender As Object, e As EventArgs) Handles PictureBox22.Click
        ch(21) = 1 - ch(21)
        PictureBox22.Image = Image.FromFile(Ljo(ch(21)))
        Dim mr As New FileStream(ljc(19), FileMode.Create)
        Dim we1 As New StreamWriter(mr, Encoding.GetEncoding("gb2312"))
        we1.WriteLine(CStr(ch(21)))
        we1.Close()
    End Sub '进度条剩余百分比
#End Region
#End Region
#Region "<子菜单>"
    Sub Zcd()
        Panel7.Left = 0 : Panel7.Top = PictureBox7.Height
        Panel7.Visible = True : Panel7.BringToFront()
        If zp = 0 Then
        ElseIf zp = 1 Then
            Panel8.BringToFront() : Panel8.Left = Label56.Left : Panel8.Top = Label56.Top + Label56.Height : Panel8.Width = Label57.Left + Label57.Width - Label56.Left : Panel8.Height = Label57.Top + Label57.Height - Label56.Top
            For i = 0 To 8
                tb(i).Text = iClass(vb_cday, i + 1)
            Next
        ElseIf zp = 2 Then
            Panel9.BringToFront() : Panel9.Left = Label56.Left : Panel9.Top = Label56.Top + Label56.Height : Panel9.Width = Label57.Left + Label57.Width - Label56.Left : Panel9.Height = Label57.Top + Label57.Height - Label56.Top
            If File.Exists(ljc(9)) Then
                Dim re1 As New StreamReader(ljc(9), Encoding.GetEncoding("gb2312"))
                ljto_dm = re1.ReadLine
                re1.Close()
                If File.Exists(ljto_dm) And Path.GetFileName(ljto_dm) = "Checkin1175.exe" Then
                    Label41.Text = ljto_dm
                Else
                    Label41.Text = "文件错误"
                End If
            End If
        ElseIf zp = 3 Then
            Label43.Text = "保存"
            Call Chjz()
            Panel10.BringToFront() : Panel10.Left = Label56.Left / 2 : Panel10.Top = Label56.Top + Label56.Height : Panel10.Width = Label57.Left + Label57.Width - Label56.Left : Panel10.Height = Label57.Top + Label57.Height - Label56.Top
            Panel10.Visible = True
        ElseIf zp = 4 Then
            Panel12.BringToFront() : Panel12.Left = Label56.Left : Panel12.Top = Label56.Top + Label56.Height : Panel12.Width = Label57.Left + Label57.Width - Label56.Left : Panel12.Height = Label57.Top + Label57.Height - Label56.Top
            Label54.Text = "保存"
            If File.Exists(ljc(16)) Then
                Dim re8 As New StreamReader(ljc(16), System.Text.Encoding.GetEncoding("gb2312"))
                j = 0 : i = 1
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
        ElseIf zp = 5 Then
            Panel14.BringToFront() : Panel14.Left = Label56.Left : Panel14.Top = Label56.Top + Label56.Height : Panel14.Width = Label57.Left + Label57.Width - Label56.Left : Panel14.Height = Label57.Top + Label57.Height - Label56.Top
            TrackBar3.Value = 0 : Label56.Text = "0"
            ComboBox10.Items.Clear()
            ComboBox11.Items.Clear()
            For i = 0 To 23
                ComboBox10.Items.Add(CStr(i))
            Next
            For i = 0 To 59
                ComboBox11.Items.Add(CStr(i))
            Next
            ComboBox10.Text = CStr(Hour(Now))
            ComboBox11.Text = CStr(Minute(Now))
        ElseIf zp = 6 Then
            Panel15.BringToFront() : Panel15.Left = Label56.Left : Panel15.Top = Label56.Top + Label56.Height : Panel15.Width = Label57.Left + Label57.Width - Label56.Left : Panel15.Height = Label57.Top + Label57.Height - Label56.Top
            htd = 1
            RichTextBox1.Text = hsr(htd)
        ElseIf zp = 7 Then
            Panel18.BringToFront() : Panel18.Left = Label56.Left : Panel18.Top = Label56.Top + Label56.Height : Panel18.Width = Label57.Left + Label57.Width - Label56.Left : Panel18.Height = Label57.Top + Label57.Height - Label56.Top
            '倒计时
            If File.Exists(ljc(18)) Then
                Dim re10 As New StreamReader(ljc(18), Encoding.GetEncoding("gb2312"))
                i = 0
                Do Until re10.EndOfStream
                    i += 1
                    Dim ts() As String = re10.ReadLine.Split(";")
                    djstime(i) = ts(0) : djstitle(i) = ts(1)
                Loop
                ListBox1.Items.Clear()
                For j = 1 To i
                    ListBox1.Items.Add(j & " " & djstime(j) & " " & djstitle(j))
                Next
                re10.Close()
            End If
        End If
    End Sub
    Private Sub Label36_Click(sender As Object, e As EventArgs) Handles Label36.Click
        Panel7.Visible = False
        Panel10.Visible = False
    End Sub
    Private Sub Label36_MouseMove(sender As Object, e As MouseEventArgs) Handles Label36.MouseMove
        Label36.ForeColor = Color.Red
    End Sub
    Private Sub Label36_MouseLeave(sender As Object, e As EventArgs) Handles Label36.MouseLeave
        Label36.ForeColor = Color.White
    End Sub
    '修改课表
    Private Sub Label42_Click(sender As Object, e As EventArgs) Handles Label42.Click
        Dim re1 As New StreamReader(lj2, Encoding.GetEncoding("gb2312"))
        Dim tclass(9) As String
        Dim s As String
        i = 0
        Do Until re1.EndOfStream
            s = re1.ReadLine
            If Mid(s, 1, 2) = "vb" Then
                i += 1
                j = 0
            Else
                j += 1
                If i = vb_cday Then
                    tclass(j) = s
                End If
            End If
        Loop
        re1.Close()
        For i = 0 To 8
            tb(i).Text = tclass(i + 1)
        Next
    End Sub
    Private Sub Label39_Click(sender As Object, e As EventArgs) Handles Label39.Click
        For i = 0 To 8
            iClass(vb_cday, i + 1) = tb(i).Text
        Next
        Panel7.Visible = False
    End Sub
    '点名器
    Private Sub Label40_Click(sender As Object, e As EventArgs) Handles Label40.Click
        OpenFileDialog1.InitialDirectory = Application.StartupPath
        If OpenFileDialog1.ShowDialog = DialogResult.OK Then
            ljto_dm = OpenFileDialog1.FileName
            If File.Exists(ljto_dm) And Path.GetFileName(ljto_dm) = "Checkin1175.exe" Then
                Label41.Text = ljto_dm
                Dim mr1 As New FileStream(ljc(9), FileMode.Create)
                Dim we1 As New StreamWriter(mr1, Encoding.GetEncoding("gb2312"))
                we1.WriteLine(ljto_dm)
                we1.Close()
            Else
                Label41.Text = "文件错误"
                ljto_dm = ""
            End If
        End If
    End Sub
    Private Sub Label54_Click(sender As Object, e As EventArgs) Handles Label54.Click
        For i = 1 To 6
            If Val(dmte(i, 1).Text) < 0 Or Val(dmte(i, 1).Text) > 23 Then Exit Sub
            If Val(dmte(i, 2).Text) < 0 Or Val(dmte(i, 2).Text) > 59 Then Exit Sub
        Next
        Dim mr1 As New FileStream(ljc(16), FileMode.Create)
        Dim we1 As New StreamWriter(mr1, Encoding.GetEncoding("gb2312"))
        For i = 1 To 6
            For j = 1 To 2
                we1.WriteLine(CStr(Val(dmte(i, j).Text)))
            Next
        Next
        we1.Close()
        Label54.Text = "已保存"
    End Sub
    '系统时间调整
    Private Sub TrackBar3_Scroll(sender As Object, e As EventArgs) Handles TrackBar3.Scroll
        Label59.Text = CStr(TrackBar3.Value)
    End Sub
    Private Sub Label58_Click(sender As Object, e As EventArgs) Handles Label58.Click
        Dim eh, em, es As Integer
        eh = Hour(Now) : em = Minute(Now) : es = Second(Now)
        Dim ec As Integer = TrackBar3.Value
        es = es + ec
        Do Until es >= 0 AndAlso es < 60
            If es < 0 Then
                es += 60
                em -= 1
                If em = -1 Then em = 59 : eh -= 1
                If eh = -1 Then eh = 23
            Else
                es -= 60
                em += 1
                If em = 60 Then em = 0 : eh += 1
                If eh = 24 Then eh = 0
            End If
        Loop
        TimeOfDay = CDate(eh & ":" & em & ":" & es)
        Panel7.Visible = False
    End Sub
    Private Sub Label63_Click(sender As Object, e As EventArgs) Handles Label63.Click
        Dim eh, em, es As Integer
        eh = Val(ComboBox10.Text) : em = Val(ComboBox11.Text) : es = Second(Now)
        TimeOfDay = CDate($"{eh}:{em}:{es}")
    End Sub
    '帮助
    Private Sub HLabel_click(sender As Object, e As EventArgs)
        htd = Val(sender.Tag)
        For i = 0 To 3
            If i = htd Then
                hlb(i).ForeColor = Color.Tomato
            Else
                hlb(i).ForeColor = Color.White
            End If
        Next
        RichTextBox1.Text = hsr(htd)
    End Sub
    '倒计时-新建
    Private Sub Label81_Click(sender As Object, e As EventArgs) Handles Label81.Click
        djstd = 0
        Call Djs()
        TextBox22.Text = ""
        If DjsGetNum() = 10 Then Label81.BackColor = Color.DimGray Else Label81.BackColor = Color.LimeGreen

        Call DjsW()

    End Sub
    '倒计时-修改
    Private Sub Label82_Click(sender As Object, e As EventArgs) Handles Label82.Click
        djstd = 1
        Call Djs()
        TextBox22.Text = ""
        Call DjsW()
    End Sub
    Private Sub ListBox1_Click(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged
        i = ListBox1.SelectedIndex + 1
        TextBox22.Text = djstime(i) & "；" & djstitle(i)
    End Sub
    Private Sub Label83_Click(sender As Object, e As EventArgs) Handles Label83.Click
        Dim ei As Integer
        j = ListBox1.SelectedIndex + 1
        i = DjsGetNum()
        If i = j Then
            djstime(i) = ""
            ei = j - 2
        Else
            For k = j To i
                If k <= 9 Then
                    djstime(k) = djstime(k + 1)
                    djstitle(k) = djstitle(k + 1)
                End If
            Next
            djstime(i) = ""
            ei = j - 1
        End If
        i = DjsGetNum()
        ListBox1.Items.Clear()
        If i <> 0 Then
            For j = 1 To i
                ListBox1.Items.Add(j & " " & djstime(j) & " " & djstitle(j))
            Next
            ListBox1.SelectedIndex = ei
        End If
        Call DjsW()
    End Sub
    Private Sub Label84_Click(sender As Object, e As EventArgs) Handles Label84.Click
        i = ListBox1.SelectedIndex + 1
        Dim t As String
        If i > 1 Then
            t = djstime(i) : djstime(i) = djstime(i - 1) : djstime(i - 1) = t
            t = djstitle(i) : djstitle(i) = djstitle(i - 1) : djstitle(i - 1) = t
            For k = 1 To 10
                If djstime(k) = "" Then Exit For
            Next
            k -= 1
            ListBox1.Items.Clear()
            For j = 1 To k
                ListBox1.Items.Add(j & " " & djstime(j) & " " & djstitle(j))
            Next
            ListBox1.SelectedIndex = i - 2
        End If
        Call DjsW()
    End Sub
    Private Sub PictureBox27_Click(sender As Object, e As EventArgs) Handles PictureBox27.Click
        ch(26) = 1 - ch(26)

    End Sub
    Private Sub Label85_Click(sender As Object, e As EventArgs) Handles Label85.Click
        i = ListBox1.SelectedIndex + 1
        Dim t As String
        If i > 0 And i < ListBox1.Items.Count Then
            t = djstime(i) : djstime(i) = djstime(i + 1) : djstime(i + 1) = t
            t = djstitle(i) : djstitle(i) = djstitle(i + 1) : djstitle(i + 1) = t
            For k = 1 To 10
                If djstime(k) = "" Then Exit For
            Next
            k -= 1
            ListBox1.Items.Clear()
            For j = 1 To k
                ListBox1.Items.Add(j & " " & djstime(j) & " " & djstitle(j))
            Next
            ListBox1.SelectedIndex = i
        End If
        Call DjsW()
    End Sub
#End Region
#Region "<子模块，点名器设置>"
    Sub Chjz()
        For Each ce As CheckBox In Panel11.Controls
            Dim k As Integer
            k = Val(ce.Tag)
            i = (k - 1) \ 6 + 1
            j = (k - 1) Mod 6 + 1
            dmc(i, j) = ce
        Next
        If File.Exists(ljc(10)) Then
            Dim re1 As New StreamReader(ljc(10), Encoding.GetEncoding("gb2312"))
            Dim s As String
            Do Until re1.EndOfStream
                s = re1.ReadLine
                If Mid(s, 1, 1) = "#" Then
                    i = Val(Mid(s, 2, 1))
                    j = 0
                Else
                    j += 1
                    dmc(i, j).Checked = (s = "True")
                    dmb(i, j) = dmc(i, j).Checked
                End If
            Loop
            re1.Close()
        End If '文件加载
        For i = 1 To 7
            If i <> 6 Then
                For j = 1 To 6
                    dmb(i, j) = dmc(i, j).Checked
                Next
            Else
                For j = 1 To 6
                    dmb(i, j) = False
                Next
            End If
        Next
    End Sub
    Private Sub Label43_Click(sender As Object, e As EventArgs) Handles Label43.Click
        Dim mr1 As New FileStream(ljc(10), FileMode.Create)
        Dim we1 As New StreamWriter(mr1, Encoding.GetEncoding("gb2312"))
        For i = 1 To 7
            we1.WriteLine("#" & i)
            For j = 1 To 6
                we1.WriteLine(dmc(i, j).Checked.ToString)
            Next
        Next
        we1.Close()
        Label43.Text = "已保存"
    End Sub
    Private Sub Label44_Click(sender As Object, e As EventArgs) Handles Label44.Click
        Call dms()
    End Sub
#End Region
End Class