Public Class VbaStylesLoader

    Public Function getStyleSets() As Collection
        Dim text As String = readText()
        Dim parser As New VbaStylesParser

        Return parser.parse(text)

    End Function

    Private Function readText() As String
        Return My.Computer.FileSystem.ReadAllText("C:\Users\Tomek\Dropbox\tm\mielk\mielk-ui\mielk-ui\docs\test.css")
    End Function

End Class
