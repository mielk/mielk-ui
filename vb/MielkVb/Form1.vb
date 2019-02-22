Public Class Form1
    Private pStylesManager As StylesManager
    Private pWindow As Window

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call loadStyles()
        Call initializeForm()
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

    Private Sub initializeForm()
        pWindow = New Window()
        With pWindow
            Call .SetListener(Me)
            Call .SetStylesManager(pStylesManager)
            Call .Display()
        End With
    End Sub
    
End Class