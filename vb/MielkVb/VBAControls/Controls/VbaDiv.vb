Public Class VbaDiv
    Implements IVbaControl
    Implements IVbaContainer

    Private pVbNetObject As Div
    Private pContainer As IVbaContainer
    Private pId As String
    Private pCssClass As String
    Private pTest As String


    Public Sub New(parent As IVbaContainer)
        pContainer = parent
    End Sub


    Public Sub LoadFromXml(xmlNode As Xml.XmlNode) Implements IVbaControl.LoadFromXml
        Call loadAttributes(xmlNode)
        pVbNetObject = New Div(pContainer.getVbNetObject, pId)
        Call pVbNetObject.AddStyleClass(pCssClass)
    End Sub

    Private Sub loadAttributes(xmlNode As Xml.XmlNode)
        On Error Resume Next
        pId = xmlNode.Attributes.ItemOf("id").Value
        pCssClass = xmlNode.Attributes.ItemOf("class").Value
        pTest = xmlNode.Attributes.ItemOf("x").Value
        On Error GoTo 0
    End Sub

    Public Sub SetParent(value As IVbaContainer) Implements IVbaControl.SetParent
        pContainer = value
    End Sub

    Public Function getVbNetObject() As Object Implements IVbaContainer.getVbNetObject, IVbaControl.getVbNetObject
        Return pVbNetObject
    End Function

End Class
