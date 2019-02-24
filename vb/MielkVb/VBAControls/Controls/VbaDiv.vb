Public Class VbaDiv
    Implements IVbaControl
    Implements IVbaContainer

    Private pVbNetObject As Div
    Private pContainer As IVbaContainer


    Public Sub New(parent As IVbaContainer)
        pContainer = parent
    End Sub


    Public Sub LoadFromXml(xmlNode As Xml.XmlNode) Implements IVbaControl.LoadFromXml
        Stop
    End Sub

    Public Sub SetParent(value As IVbaContainer) Implements IVbaControl.SetParent
        pContainer = value
    End Sub

    Public Function getVbNetObject() As Object Implements IVbaContainer.getVbNetObject
        Return pVbNetObject
    End Function

End Class
