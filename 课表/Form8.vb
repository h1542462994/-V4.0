Imports System.IO
Public Class Form8
    Dim ch(7, 6) As CheckBox
    Dim i, j As Integer
    Private Sub Form8_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call chjz()
    End Sub
    Sub chjz()
        For Each ce As CheckBox In Panel1.Controls
            Dim t_i, t_j As Integer
            Dim k As Integer
            k = Val(ce.Tag)
            t_i = (k - 1) \ 6 + 1
            t_j = (k - 1) Mod 6 + 1
            ch(t_i, t_j) = ce
        Next
#Region "文件加载"
        If File.Exists(ljdm) Then
            Dim reader As New StreamReader(ljdm, System.Text.Encoding.GetEncoding("gb2312"))
            Dim s As String
            Do Until reader.EndOfStream
                s = reader.ReadLine
                If Mid(s, 1, 1) = "#" Then
                    i = Val(Mid(s, 2, 1))
                    j = 0
                Else
                    j += 1
                    ch(i, j).Checked = (s = "True")
                End If
            Loop
        End If
#End Region
        For i = 1 To 7
            If i <> 6 Then
                For j = 1 To 6
                    dmb(i, j) = ch(i, j).Checked
                Next
            Else
                For j = 1 To 6
                    dmb(i, j) = False
                Next
            End If
        Next
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If File.Exists(lj_ed) Then
            Form9.Show()
        End If
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim MyStream As New FileStream(ljdm, FileMode.Create)
        Dim writer As New StreamWriter(MyStream, System.Text.Encoding.GetEncoding("gb2312"))
        For i = 1 To 7
            If i <> 6 Then
                writer.WriteLine("#" & i)
                For j = 1 To 6
                    writer.WriteLine(dmb(i, j).ToString)
                Next
            End If
        Next
        writer.Close()
    End Sub
End Class