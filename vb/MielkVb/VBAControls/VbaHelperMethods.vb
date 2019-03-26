Module VbaHelperMethods

    Public Function createControlByTag(tag As String, parent As IVbaContainer) As IVbaControl
        Select Case tag.ToLower
            Case "div" : Return New VbaDiv(parent)
            Case "button" : Return New VbaButton(parent)
            Case Else : Return Nothing
        End Select
    End Function

    Public Function createControlByXmlNode(node As System.Xml.XmlNode, parent As IVbaContainer) As IVbaControl
        Dim elementTag As String
        Dim control As IVbaControl
        '----------------------------------------------------------------------------------------------------------
        elementTag = node.Name
        control = createControlByTag(elementTag, parent)
        If Not control Is Nothing Then
            Call control.LoadFromXml(node)
        End If
        Return control
    End Function

End Module
