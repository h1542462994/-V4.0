Imports System.IO
Public Class 打开点名器
    Dim imd As Boolean
    Dim xn, yn As Single
    Private Sub 打开点名器_Load(sender As Object, e As EventArgs) Handles Me.Load
        Left = (sw - Width) / 2 : Top = (sh - Height) / 2
    End Sub
    Private Sub 打开点名器_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        dmt = 0
        Timer1.Enabled = True
    End Sub
    Private Sub M_MouseDown(sender As Object, e As MouseEventArgs) Handles Me.MouseDown, Label1.MouseDown
        imd = True
        xn = e.X : yn = e.Y
    End Sub
    Private Sub M_MouseMove(sender As Object, e As MouseEventArgs) Handles Me.MouseMove, Label1.MouseMove
        If imd Then
            Left = Left + e.X - xn
            Top = Top + e.Y - yn
        End If
    End Sub
    Private Sub M_MouseUp(sender As Object, e As MouseEventArgs) Handles Me.MouseUp, Label1.MouseUp
        imd = False
    End Sub
    Private Sub Label3_Click(sender As Object, e As EventArgs) Handles Label3.Click
        If File.Exists(ljto_dm) Then
            Shell(ljto_dm, vbNormalFocus, False)
        End If
        Me.Hide()
    End Sub
    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click
        Me.Hide()
        Timer1.Enabled = False
    End Sub
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        dmt += 1
        Label3.Text = "是(" & (8 - dmt) & ")"
        If dmt = 8 Then
            Timer1.Enabled = False
            If File.Exists(ljto_dm) Then
                Shell(ljto_dm, vbNormalFocus)
            End If
            Me.Hide()
        End If
    End Sub
End Class