Imports System.Xml

Public Class VBAWindow
    Implements IVbaContainer

    Private pControls As Collection
    Private pVbWindow As Window
    '----------------------------------------------------------------------------------------------------------------------------------------
    Private pId As String
    Private pCssClass As String
    Private pCaption As String
    '----------------------------------------------------------------------------------------------------------------------------------------

    Public Sub SetVbWindow(value As Window)
        pVbWindow = value
    End Sub

    Public Sub insertControls(htmlNode As XmlNode)
        Dim node As Object
        Dim control As IVbaControl
        '----------------------------------------------------------------------------------------------------------

        Call loadAttributes(htmlNode)

        pControls = New Collection
        For Each node In htmlNode.ChildNodes
            control = createControlByXmlNode(node, Me)
            If Not control Is Nothing Then
                Call pControls.Add(control)
                Call pVbWindow.AddControl(control.getVbNetObject)
            End If
        Next node

    End Sub

    Private Sub loadAttributes(xmlNode As Xml.XmlNode)
        On Error Resume Next
        pId = xmlNode.Attributes.ItemOf("id").Value
        pCssClass = xmlNode.Attributes.ItemOf("class").Value
        pCaption = xmlNode.Attributes.ItemOf("caption").Value
        On Error GoTo 0

        With pVbWindow
            Call .SetId(pId)
            Call .SetTitle(pCaption)
            Call .SetStyleClasses(pCssClass, True)

        End With

    End Sub

    Public Function GetVbNetObject() Implements IVbaContainer.getVbNetObject
        Return pVbWindow
    End Function

    'Public Function RaiseEvent_Load()
    '    'Debug.Print "load"
    'End Function

    'Public Function RaiseEvent_Disposed()
    '    'Debug.Print "disposed"
    'End Function

    'Public Function RaiseEvent_MouseEnter()
    '    'Debug.Print "mouse enter"
    'End Function

    'Public Function RaiseEvent_MouseLeave()
    '    'Debug.Print "mouse leave"
    'End Function

    'Public Function RaiseEvent_GotFocus()
    '    Stop
    'End Function

    'Public Function RaiseEvent_LostFocus()
    '    'Call setBorderBox(True)
    '    'Debug.Print "got focus"
    'End Function

    'Public Function RaiseEvent_SizeChanged(ByVal width As Single, ByVal height As Single, ByVal left As Single, ByVal top As Single)
    '    'Debug.Print "size changed | width: " & width & " | height: " & height & " | left: " & left & " | top: " & top
    'End Function

    'Public Function RaiseEvent_Scroll(ByVal left As Single, ByVal top As Single)
    '    'Debug.Print "scroll"
    'End Function

End Class
