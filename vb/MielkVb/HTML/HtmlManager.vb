Imports System.Xml
Imports System.IO

Public Class HtmlManager
    Private pXmlDoc As New XmlDataDocument()
    Private pXmlNode As XmlNode

    Public Sub New(path As String)
        Dim fs As New FileStream(path, FileMode.Open, FileAccess.Read)
        '----------------------------------------------------------------------------------------------------
        Call pXmlDoc.Load(fs)
        pXmlNode = pXmlDoc.ChildNodes(0)
    End Sub

    Public Function GetFormNode() As XmlNode
        Return pXmlNode
    End Function

End Class