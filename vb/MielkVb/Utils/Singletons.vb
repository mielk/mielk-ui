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

    Public Function RQ() As ResizeQueue
        Static instance As ResizeQueue
        '----------------------------------------------------------------------
        If instance Is Nothing Then
            instance = New ResizeQueue()
        End If
        Return instance
    End Function

    Public Function VQ() As ViewQueue
        Static instance As ViewQueue
        '----------------------------------------------------------------------
        If instance Is Nothing Then
            instance = New ViewQueue()
        End If
        Return instance
    End Function

End Module