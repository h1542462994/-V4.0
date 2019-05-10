Imports System.IO
Public Class Form9
    Private Sub Form9_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Opacity = 0.7
        Me.Top = 0 : Me.Left = 0
        Me.Height = sh : Me.Width = sw
        Panel1.Left = (sw - Panel1.Width) / 2 : Panel1.Top = (sh - Panel1.Height) / 2
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If File.Exists(lj_ed) Then
            Shell(lj_ed, vbNormalFocus)
        End If
        Me.Hide()
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Hide()
    End Sub
End Class