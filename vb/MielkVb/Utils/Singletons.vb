Module Singletons

    Public Function GetStylesManager() As StylesManager
        Static instance As StylesManager
        '----------------------------------------------------------------------
        If instance Is Nothing Then
            instance = New StylesManager()
        End If
        Return instance
    End Function

    Public Function GetTranslation(tag As String) As String
        Select Case tag
            Case "Ok" : Return "OK"
            Case "Cancel" : Return "Cancel"
            Case Else : Return "[Unknown]"
        End Select
    End Function

End Module