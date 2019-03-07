Imports System.Xml

Public Class VBAWindow
    Implements IVbaContainer

    Private pControls As Collection
    Private pVbWindow As Window

    Public Sub SetVbWindow(value As Window)
        pVbWindow = value
    End Sub

    Public Sub insertControls(htmlNode As XmlNode)
        Dim node As Object
        Dim elementTag As String
        Dim control As IVbaControl
        '----------------------------------------------------------------------------------------------------------

        pControls = New Collection
        For Each node In htmlNode.ChildNodes
            'elementTag = node.BaseName
            elementTag = node.Name
            control = createControlByTag(elementTag, Me)
            If Not control Is Nothing Then
                Call control.LoadFromXml(node)
                Call pControls.Add(control)
                Call pVbWindow.AddControl(control.getVbNetObject)
            End If
        Next node

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
