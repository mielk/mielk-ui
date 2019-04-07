Public Class VbaDiv
    Implements IVbaControl
    Implements IVbaContainer

    Private pVbNetObject As Div
    Private pContainer As IVbaContainer
    Private pId As String
    Private pCssClass As String
    Private pControls As Collection


    Public Sub New(parent As IVbaContainer)
        pContainer = parent
    End Sub


    Public Sub LoadFromXml(xmlNode As Xml.XmlNode) Implements IVbaControl.LoadFromXml
        Call loadAttributes(xmlNode)
        pVbNetObject = New Div(pContainer.getVbNetObject, pId)
        With pVbNetObject
            Call .SetListener(Me)
            Call .SetStyleClasses(pCssClass, True)
        End With
        If xmlNode.ChildNodes.Count Then
            Call loadChildrenControls(xmlNode.ChildNodes)
        End If
    End Sub

    Private Sub loadAttributes(xmlNode As Xml.XmlNode)
        On Error Resume Next
        pId = xmlNode.Attributes.ItemOf("id").Value
        pCssClass = xmlNode.Attributes.ItemOf("class").Value
        On Error GoTo 0
    End Sub

    Private Sub loadChildrenControls(childNodes As System.Xml.XmlNodeList)
        Dim node As Xml.XmlNode
        Dim control As IVbaControl
        '----------------------------------------------------------------------------------------------------------------------------

        pControls = New Collection
        For Each node In childNodes
            control = createControlByXmlNode(node, Me)
            If Not control Is Nothing Then
                Call pControls.Add(control)
                Call pVbNetObject.AddControl(control.getVbNetObject)
            End If
        Next

    End Sub

    Public Sub SetParent(value As IVbaContainer) Implements IVbaControl.SetParent
        pContainer = value
    End Sub

    Public Function getVbNetObject() As Object Implements IVbaContainer.getVbNetObject, IVbaControl.getVbNetObject
        Return pVbNetObject
    End Function

End Class
