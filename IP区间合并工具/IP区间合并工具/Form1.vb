Imports System.Text

Public Class Form1
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox1.TextLength > 0 Then
            If IO.File.Exists(TextBox1.Text) Then
                Button1.Enabled = False
                Threading.ThreadPool.QueueUserWorkItem(AddressOf go, TextBox1.Text)
            End If
        End If
    End Sub

    Private Sub go(ByVal file As String)
        Dim intervals As New List(Of Interval)
        Dim tmp() As String
        For Each txt In IO.File.ReadLines(file)
            If txt.Length = 0 Then Continue For
            If txt.Contains("-") Then
                tmp = txt.Split("-")
                If tmp.Length = 2 Then
                    intervals.Add(New Interval(IpToLong(tmp(0)), IpToLong(tmp(1))))
                End If
            Else
                intervals.Add(New Interval(IpToLong(txt), IpToLong(txt)))
            End If
        Next
        Dim fi As New IO.FileInfo(file)
        Dim newfile As String = fi.DirectoryName & "\New_" & fi.Name
        Dim IPRes = Merge(intervals)
        Dim appstr As New System.Text.StringBuilder
        For Each item In IPRes
            appstr.AppendLine(LongToIP(item.st) & "-" & LongToIP(item.ed))
        Next
        IO.File.WriteAllText(newfile, appstr.ToString)
        Invoke(Sub()
                   TextBox2.Text = newfile
                   RichTextBox2.Text = YuLan(newfile)
                   Button1.Enabled = True
               End Sub)
    End Sub

    Public Shared Function IpToLong(strIP As String) As Long
        Try
            Dim ip(3) As Long
            Dim s As String() = strIP.Split(".")
            ip(0) = Long.Parse(s(0))
            ip(1) = Long.Parse(s(1))
            ip(2) = Long.Parse(s(2))
            ip(3) = Long.Parse(s(3))
            Return (ip(0) << 24) + (ip(1) << 16) + (ip(2) << 8) + ip(3)
        Catch ex As Exception
            Return 0
        End Try
    End Function

    Public Shared Function LongToIP(longIP As Long) As String
        Dim sb As New StringBuilder
        sb.Append(longIP >> 24)
        sb.Append(".")
        sb.Append((longIP And &HFFFFFF) >> 16)
        sb.Append(".")
        sb.Append((longIP And &HFFFF) >> 8)
        sb.Append(".")
        sb.Append((longIP And &HFF))
        Return sb.ToString()
    End Function

    Public Function Merge(ByVal intervals As List(Of Interval)) As List(Of Interval)
        Dim res As New List(Of Interval)
        If intervals.Count = 0 Then Return res
        intervals = intervals.OrderBy(Function(i) i.st).ToList
        res.Add(intervals(0))
        For i As Long = 1 To intervals.Count - 1
            If intervals(i).st <= res(res.Count - 1).ed Then
                res(res.Count - 1).ed = Math.Max(intervals(i).ed, res(res.Count - 1).ed)
            Else
                res.Add(intervals(i))
            End If
        Next
        Return res
    End Function

    Private Sub TextBox1_DragDrop(sender As Object, e As DragEventArgs) Handles TextBox1.DragDrop
        TextBox1.Text = e.Data.GetData(DataFormats.FileDrop).getvalue(0).ToString
        RichTextBox1.Text = YuLan(TextBox1.Text)
    End Sub

    Private Function YuLan(ByVal file As String) As String
        Dim appstr As New System.Text.StringBuilder
        If IO.File.Exists(file) Then
            Dim x As Integer
            For Each s In IO.File.ReadLines(file)
                If s.Length = 0 Then Continue For
                x += 1
                If x > 100 Then Exit For
                appstr.AppendLine(s)
            Next
        End If
        Return appstr.ToString
    End Function

    Private Sub TextBox1_DragEnter(sender As Object, e As DragEventArgs) Handles TextBox1.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop, False) = True Then
            e.Effect = DragDropEffects.Copy '指针变成+号，表示正在拖入数据
        End If
    End Sub

End Class

Public Class Interval
    Public st As Long
    Public ed As Long

    Public Sub New()
        st = 0
        ed = 0
    End Sub

    Public Sub New(ByVal s As Long, ByVal e As Long)
        st = s
        ed = e
    End Sub
End Class
