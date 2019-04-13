Imports System.Xml

Public Class Form1
    Private pStylesManager As StylesManager
    Private pHtml As HtmlManager
    Private pVbaWindow As VBAWindow
    Private pWindow As Window


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Me.Icon = New System.Drawing.Icon("C:\Users\Tomek\Desktop\mlmh.ico")
        Call testHtml()
        'Call testTransparentFrame()

        With Me.ListBox1
            Call .Items.Add("a")
            Call .Items.Add("b")
            Call .Items.Add("c")
            Call .Items.Add("d")
            Call .Items.Add("e")
            Call .Items.Add("f")
            Call .Items.Add("g")
            Call .Items.Add("h")
            Call .Items.Add("i")
            Call .Items.Add("j")
            Call .Items.Add("k")
            Call .Items.Add("l")
            Call .Items.Add("m")
            Call .Items.Add("n")
            Call .Items.Add("o")
            Call .Items.Add("p")
            Call .Items.Add("q")
            Call .Items.Add("r")
            Call .Items.Add("s")
            Call .Items.Add("t")
            Call .Items.Add("u")
            Call .Items.Add("v")
            Call .Items.Add("w")
            Call .Items.Add("x")
            Call .Items.Add("y")
            Call .Items.Add("z")
        End With
    End Sub


    Private Sub testHtml()
        Call loadStyles()
        Call initializeForm()
        Call loadHtml()

        Me.TextBox1.AutoSize = False
        Me.TextBox1.Size = New System.Drawing.Size(142, 87)

        Call pWindow.SetIcon("C:\Users\Tomek\Desktop\testIcon.ico")
        Call pWindow.Display()
    End Sub

    Private Sub loadStyles()
        Dim stylesLoader As VbaStylesLoader
        Dim styleSets As Collection
        Dim styleSet As VbaStyleSet

        pStylesManager = GetStylesManager()
        stylesLoader = New VbaStylesLoader
        styleSets = stylesLoader.getStyleSets()

        For Each styleSet In styleSets
            Call pStylesManager.LoadStyleSet(styleSet)
        Next

    End Sub

    Private Sub loadHtml()
        Dim pNode As XmlNode
        '---------------------------------------------------------------------------------------------------------------------------------------
        pHtml = New HtmlManager("C:\Users\Tomek\Dropbox\tm\mielk\mielk-ui\mielk-ui\docs\form_basic.xml")
        pNode = pHtml.GetFormNode()
        Call pVbaWindow.insertControls(pNode)
    End Sub

    Private Sub initializeForm()
        pVbaWindow = New VBAWindow()
        pWindow = New Window()
        With pWindow
            Call .SetListener(pVbaWindow)
            Call .SetStylesManager(pStylesManager)
        End With
        Call pVbaWindow.SetVbWindow(pWindow)
    End Sub





    'Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    'End Sub

    'Private Sub ListBox1_DragDrop(sender As Object, e As DragEventArgs) Handles ListBox1.DragDrop
    '    Dim files() As String = e.Data.GetData(DataFormats.FileDrop)
    '    For Each path In files
    '        MsgBox(path)
    '    Next
    'End Sub

    'Private Sub ListBox1_DragEnter(sender As Object, e As DragEventArgs) Handles ListBox1.DragEnter
    '    If e.Data.GetDataPresent(DataFormats.FileDrop) Then
    '        e.Effect = DragDropEffects.Copy
    '    End If
    'End Sub

    'Private Sub ListBox1_DragLeave(sender As Object, e As EventArgs) Handles ListBox1.DragLeave
    'End Sub

    'Private Sub ListBox1_DragOver(sender As Object, e As DragEventArgs) Handles ListBox1.DragOver
    'End Sub

    'Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged

    'End Sub


    Private Sub x(sender As Object, e As EventArgs) Handles Me.Click
        Dim panel As TransparentPanel
        panel = New TransparentPanel
        'Dim panel As Panel
        'panel = New Panel

        Exit Sub

        panel.BackColor = colorFromString("rgb(230, 230, 230)")

        panel.Opacity = 75
        panel.Width = 500
        panel.Height = 500
        panel.setText("testowy tekst")

        Me.Controls.Add(panel)
        panel.BringToFront()


        Dim label As Label
        label = New Label
        label.BackColor = Color.Transparent
        'label.Parent = panel
        label.Text = "text"
        label.Font = New Font("Century Gothic", 13, FontStyle.Bold)
        label.Top = 220
        label.Left = 250
        Call panel.Controls.Add(label)


        'Dim form As TransparentWindow
        'form = New TransparentWindow
        'form.ControlBox = False
        'form.Text = String.Empty
        'form.FormBorderStyle = FormBorderStyle.None

        'Dim lbl As Label
        'lbl = New Label
        'lbl.Text = "test"
        'lbl.Size = New Size(200, 100)
        'lbl.Left = 100
        'lbl.Top = 150
        'Call form.Controls.Add(lbl)
        'form.ShowInTaskbar = False

        'form.TopMost = True
        'Call form.updateOpacity(120, True)
        'form.Show()
    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click, Label2.Click

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged

    End Sub
End Class