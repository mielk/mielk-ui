Public Class VbaStylesHeader
    Private pElement As String
    Private pClass As String
    Private pId As String
    Private pEvent As String



    Public Property HeaderElement As String
        Get
            Return pElement
        End Get
        Set(ByVal value As String)
            pElement = value
        End Set
    End Property

    Public Property HeaderClass As String
        Get
            Return pClass
        End Get
        Set(ByVal value As String)
            pClass = value
        End Set
    End Property

    Public Property HeaderId As String
        Get
            Return pId
        End Get
        Set(ByVal value As String)
            pId = value
        End Set
    End Property

    Public Property HeaderEvent As String
        Get
            Return pEvent
        End Get
        Set(ByVal value As String)
            pEvent = value
        End Set
    End Property

End Class