Public Class Form6
    Dim imd As Boolean
    Dim x0, y0 As Single
    Dim Fd As Date
    Private Sub Form6_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Timer1.Enabled = True
        Call Timer1_Tick(Timer1, e)
        For i = 1 To 10
            AddHandler djslb(i).MouseDown, AddressOf M_MouseDown
            AddHandler djslb(i).MouseMove, AddressOf M_MouseMove
            AddHandler djslb(i).MouseUp, AddressOf M_MouseUp
        Next
    End Sub
    Private Sub M_MouseDown(sender As Object, e As MouseEventArgs) Handles Me.MouseDown
        imd = True
        x0 = e.X : y0 = e.Y
    End Sub
    Private Sub M_MouseMove(sender As Object, e As MouseEventArgs) Handles Me.MouseMove
        If imd Then
            Left = Left + e.X - x0
            Top = Top + e.Y - y0
        End If
    End Sub
    Private Sub M_MouseUp(sender As Object, e As MouseEventArgs) Handles Me.MouseUp
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
        For i = 1 To 10
            If djstime(i) <> "" Then
                j = DateDiff(DateInterval.Day, Now, CDate(djstime(i)))
                djslb(i).Text = djstitle(i) & " " & j & "天"
                If j >= 300 Then
                    djslb(i).ForeColor = Color.DeepSkyBlue
                ElseIf j >= 200 Then
                    k = j - 200
                    djslb(i).ForeColor = Color.FromArgb(0, 255 - (255 - 191) * k / 100, 127 + (255 - 127) * k / 100) 'SpringGreen>DeepSkyBlue
                ElseIf j >= 100 Then
                    k = j - 100
                    djslb(i).ForeColor = Color.FromArgb(255 - 255 * k / 100, 255, 127 * k / 100) 'Yellow>SpringGreen
                ElseIf j >= 0 Then
                    k = j
                    djslb(i).ForeColor = Color.FromArgb(255, 69 + (255 - 69) * k / 100, 0) 'Orangred>Yelloew
                Else
                    djslb(i).ForeColor = Color.Silver
                End If
            Else
                Exit For
            End If
        Next
        i -= 1
        If i >= 1 Then
            Height = Label1.Height * i
        Else
            Height = Label1.Height
            Label1.Text = "No Schedule!"
            Label1.ForeColor = Color.White
        End If
    End Sub
End Class