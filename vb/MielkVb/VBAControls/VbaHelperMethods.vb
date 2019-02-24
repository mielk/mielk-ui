Module VbaHelperMethods

    Public Function createControlByTag(tag As String, parent As IVbaContainer) As IVbaControl
        Select Case tag.ToLower
            Case "div" : Return New VbaDiv(parent)
            Case Else : Return Nothing
        End Select
    End Function

End Module
