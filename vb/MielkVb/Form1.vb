Imports System.Xml

Public Class Form1
    Private pStylesManager As StylesManager
    Private pHtml As HtmlManager
    Private pVbaWindow As VBAWindow
    Private pWindow As Window


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Call testHtml()
        'Call testTransparentFrame()
    End Sub


    Private Sub testHtml()
        Call loadStyles()
        Call initializeForm()
        Call loadHtml()

        Me.TextBox1.AutoSize = False
        Me.TextBox1.Size = New System.Drawing.Size(142, 87)

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

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim form As Window
        form = New Window
        form.Show()
        form.ControlBox = False
        form.Text = String.Empty
        form.FormBorderStyle = FormBorderStyle.None
    End Sub
End Class