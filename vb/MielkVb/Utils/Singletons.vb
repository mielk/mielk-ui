Module Singletons

    Public Function GetStylesManager() As StylesManager
        Static instance As StylesManager
        '----------------------------------------------------------------------
        If instance Is Nothing Then
            instance = New StylesManager()
        End If
        Return instance
    End Function


End Module