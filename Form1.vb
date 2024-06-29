Imports System.IO
Imports System.Text
Imports Markdig


Public Class Form1
    Private Async Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ToolStrip2.Visible = False


        Await WebViewHTML.EnsureCoreWebView2Async(Nothing)
        OpenFileDialog1.Filter = "HTML文件(*.html)|*.html"
        SaveFileDialog1.Filter = "HTML文件(*.html)|*.html"
        OpenFileDialog2.Filter = "MarkDown文件(*.md)|*.md"
        SaveFileDialog2.Filter = "MarkDown文件(*.md)|*.md"
        Me.Text = My.Application.Info.Title + String.Format("版本 V {0}", My.Application.Info.Version.ToString)

    End Sub













    Private closeConfirmed As Boolean = False

    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        ' 如果用户已经确认关闭，则允许关闭
        If closeConfirmed Then
            Return
        End If

        ' 检查RichTextBox1和RichTextBox2是否为空
        If String.IsNullOrWhiteSpace(RichTextBoxHTML.Text) AndAlso String.IsNullOrWhiteSpace(RichTextBoxmarkdown.Text) Then
            ' 如果两个RichTextBox都是空的，允许关闭
            e.Cancel = False
        Else
            ' 如果至少有一个RichTextBox不是空的，显示提醒信息
            Dim result As DialogResult = MessageBox.Show("所有内容未清除，确定要关闭吗?", "确认关闭", MessageBoxButtons.YesNo)
            If result = DialogResult.Yes Then
                ' 如果用户确认关闭，设置标志并再次尝试关闭窗口
                closeConfirmed = True
                Me.Close()
            Else
                ' 如果用户取消关闭，取消关闭事件
                e.Cancel = True
            End If
        End If
    End Sub









    Private Sub 保存HTML文件ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 保存HTML文件ToolStripMenuItem.Click
        ' 显示 SaveFileDialog
        If SaveFileDialog1.ShowDialog() = DialogResult.OK Then
            ' 将 RichTextBox 的内容转换为 HTML 并保存到文件
            Saveashtml(RichTextBoxHTML.Text, SaveFileDialog1.FileName)
        End If
    End Sub

    Private Sub Saveashtml(content As String, filepath As String)
        '将文本内容转换为简单的 html 格式
        Dim htmlcontent As String = "<html><body><pre>" & content.Replace(vbCrLf, "<br>") & "</pre></body></html>"

        '将 html 内容写入到文件
        System.IO.File.WriteAllText(filepath, htmlcontent)
    End Sub







    Private Sub 保存MarkDown文件ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 保存MarkDown文件ToolStripMenuItem.Click
        If SaveFileDialog2.ShowDialog() = DialogResult.OK Then
            ' 保存 RichTextBox2 的内容为 Markdown 文件
            System.IO.File.WriteAllText(SaveFileDialog2.FileName, RichTextBoxmarkdown.Text)
        End If
    End Sub



    Private Sub RichTextBoxHTML_TextChanged_(sender As Object, e As EventArgs) Handles RichTextBoxHTML.TextChanged
        ToolStripLabel1.Text = "字数为：" & RichTextBoxHTML.TextLength
        ToolStripTextBox1.Text = ""
        ToolStripTextBox2.Text = ""
        Dim HTMLcontent As String = RichTextBoxHTML.Text
        WebViewHTML.NavigateToString(HTMLcontent)
        WebBrowser1.DocumentText = HTMLcontent
        Dim rtb As RichTextBox = DirectCast(sender, RichTextBox)
        UpdateLineInfo(rtb)
    End Sub

    Private Sub RichTextBoxmarkdown_TextChanged_1(sender As Object, e As EventArgs) Handles RichTextBoxmarkdown.TextChanged
        ToolStripLabel1.Text = "字数为：" & RichTextBoxmarkdown.TextLength
        ToolStripTextBox1.Text = ""
        ToolStripTextBox2.Text = ""
        Dim Markdowncontent As String = RichTextBoxmarkdown.Text
        Dim markdownHTMLcontent As String = Markdown.ToHtml(Markdowncontent)
        WebViewHTML.NavigateToString(markdownHTMLcontent)
        WebBrowser1.DocumentText = markdownHTMLcontent

        Dim rtb As RichTextBox = DirectCast(sender, RichTextBox)
        UpdateLineInfo(rtb)
    End Sub

    ' 更新label1以显示当前行和总行数
    Private Sub UpdateLineInfo(rtb As RichTextBox)
        ' 获取当前光标的位置
        Dim cursorPosition As Integer = rtb.SelectionStart
        ' 计算当前行
        Dim currentLine As Integer = rtb.GetLineFromCharIndex(cursorPosition)
        ' 计算总行数
        Dim lineCount As Integer = rtb.Lines.Length
        ' 更新label1的文本
        ToolStripLabel2.Text = $"行数信息：当前在第{currentLine + 1}行 /总共 {lineCount}行"
    End Sub

    Private Sub ToolStripButton1_Click(sender As Object, e As EventArgs) Handles ToolStripButton1.Click
        RichTextBoxmarkdown.AppendText("**" + My.Application.Info.ProductName + "**")
    End Sub

    Private Sub ToolStripButton2_Click(sender As Object, e As EventArgs) Handles ToolStripButton2.Click
        RichTextBoxmarkdown.AppendText("*" + My.Application.Info.ProductName + "*")
    End Sub

    Private Sub ToolStripButton3_Click(sender As Object, e As EventArgs) Handles ToolStripButton3.Click
        RichTextBoxmarkdown.AppendText("[" + My.Application.Info.ProductName + "](https://github.com/yilihamujiang365/HTMLandMarkDown-Editor)")
    End Sub


    Private Sub ToolStripButton5_Click(sender As Object, e As EventArgs) Handles ToolStripButton5.Click

        RichTextBoxmarkdown.AppendText(">" & My.Application.Info.Title)
    End Sub

    Private Sub ToolStripButton4_Click(sender As Object, e As EventArgs) Handles ToolStripButton4.Click

        RichTextBoxmarkdown.AppendText("| " + My.Application.Info.ProductName + " | 其他 |
| ------- | ------- |
|       |         |
")
    End Sub

    Private Sub ToolStripButton10_Click(sender As Object, e As EventArgs) Handles ToolStripButton10.Click

        WebViewHTML.GoForward()

    End Sub

    Private Sub ToolStripButton11_Click(sender As Object, e As EventArgs) Handles ToolStripButton11.Click

        WebViewHTML.CoreWebView2.OpenDevToolsWindow()
    End Sub

    Private Sub ToolStripButton12_Click(sender As Object, e As EventArgs) Handles ToolStripButton12.Click
        If ToolStripTextBox1.Text = "" Then
            MessageBox.Show("请在左边输入链接地址！")
        Else
            WebViewHTML.CoreWebView2.Navigate("https://" & ToolStripTextBox1.Text)
        End If


    End Sub

    Private Sub ToolStripButton13_Click(sender As Object, e As EventArgs) Handles ToolStripButton13.Click
        WebBrowser1.GoBack()

    End Sub









    Private Sub ToolStripButton6_Click(sender As Object, e As EventArgs) Handles ToolStripButton6.Click

        RichTextBoxmarkdown.AppendText("![示例图片](http://example.com/example.jpg)")
    End Sub



    Private Sub ToolStripButton7_Click(sender As Object, e As EventArgs) Handles ToolStripButton7.Click
        RichTextBoxmarkdown.AppendText("``` 
这是代码片段
Console.Writrline(""" + My.Application.Info.Title + """)
```")
    End Sub



    Private Sub 退出XToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 退出XToolStripMenuItem.Click
        ' 检查RichTextBox1和RichTextBox2是否为空
        If String.IsNullOrWhiteSpace(RichTextBoxHTML.Text) AndAlso String.IsNullOrWhiteSpace(RichTextBoxmarkdown.Text) Then
            ' 如果两个RichTextBox都是空的，关闭窗体
            Me.Close()
        Else
            ' 如果至少有一个RichTextBox不是空的，显示提醒信息
            MessageBox.Show("请确认所有内容已清除后再关闭。")
        End If
    End Sub

    Private Sub 关于AToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 关于AToolStripMenuItem.Click
        AboutBox1.Show()
    End Sub

    Private Sub 反馈或意见ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 反馈或意见ToolStripMenuItem.Click
        Dim linkToOpen As String = "https://forms.office.com/r/ryf0EZnNS0"
        System.Diagnostics.Process.Start(linkToOpen)
    End Sub

    Private Sub 修改HTML编辑器字体ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 修改HTML编辑器字体ToolStripMenuItem.Click
        If FontDialog1.ShowDialog() = DialogResult.OK Then
            ' 如果用户选择了字体，设置RichTextBox的字体
            RichTextBoxHTML.Font = FontDialog1.Font
        End If
    End Sub

    Private Sub 修改Markdown编辑器字体ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 修改Markdown编辑器字体ToolStripMenuItem.Click
        If FontDialog1.ShowDialog() = DialogResult.OK Then
            ' 如果用户选择了字体，设置RichTextBox的字体
            RichTextBoxmarkdown.Font = FontDialog1.Font
        End If
    End Sub


    Private Sub 新建HTML网页文件ToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles 新建HTML网页文件ToolStripMenuItem1.Click
        ' 显示 OpenFileDialog
        If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
            ' 尝试以 UTF-8 编码读取文件内容
            Try
                Using reader As New StreamReader(OpenFileDialog1.FileName, Encoding.UTF8)
                    RichTextBoxHTML.Text = reader.ReadToEnd()
                End Using
            Catch ex As Exception
                ' 如果 UTF-8 编码失败，尝试使用系统默认编码
                Using reader As New StreamReader(OpenFileDialog1.FileName)
                    RichTextBoxHTML.Text = reader.ReadToEnd()
                End Using
            End Try
        End If
    End Sub

    Private Sub 新建MarkDown文件ToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles 新建MarkDown文件ToolStripMenuItem1.Click
        ' 显示 OpenFileDialog
        If OpenFileDialog2.ShowDialog() = DialogResult.OK Then
            ' 尝试以 UTF-8 编码读取文件内容
            Try
                Using reader As New StreamReader(OpenFileDialog2.FileName, Encoding.UTF8)
                    RichTextBoxmarkdown.Text = reader.ReadToEnd()
                End Using
            Catch ex As Exception
                ' 如果 UTF-8 编码失败，尝试使用系统默认编码
                Using reader As New StreamReader(OpenFileDialog2.FileName)
                    RichTextBoxmarkdown.Text = reader.ReadToEnd()
                End Using
            End Try
        End If
    End Sub

    Private Sub ToolStripButton16_Click(sender As Object, e As EventArgs) Handles ToolStripButton16.Click
        If ToolStripTextBox2.Text = "" Then
            MessageBox.Show("请在左边输入链接地址！")
        Else
            WebBrowser1.Navigate("https://" & ToolStripTextBox2.Text)
        End If

    End Sub

    Private Sub ToolStripButton14_Click(sender As Object, e As EventArgs) Handles ToolStripButton14.Click
        WebBrowser1.Refresh()
    End Sub

    Private Sub ToolStripButton15_Click(sender As Object, e As EventArgs) Handles ToolStripButton15.Click
        WebBrowser1.GoForward()
    End Sub

    Private Sub ToolStripButton9_Click(sender As Object, e As EventArgs) Handles ToolStripButton9.Click
        WebViewHTML.Refresh()
    End Sub

    Private Sub ToolStripButton8_Click(sender As Object, e As EventArgs) Handles ToolStripButton8.Click
        WebViewHTML.GoBack()
    End Sub

    Private Sub 新建HTML网页文件ToolStripMenuItem_Click(sender As Object, e As EventArgs)
        RichTextBoxHTML.Text = ""
    End Sub

    Private Sub 新建MarkDown文件ToolStripMenuItem_Click(sender As Object, e As EventArgs)
        RichTextBoxmarkdown.Text = ""
    End Sub
End Class