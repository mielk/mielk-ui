Public Class MielkUI

    Private pForms As Forms
    Private pStylesManager As StylesManager

#Region "Constructor"

    Sub New()
        pForms = New Forms()
        pStylesManager = New StylesManager()
    End Sub

#End Region


#Region "Public properties"

    Public ReadOnly Property Forms()
        Get
            Return pForms
        End Get
    End Property

    Public ReadOnly Property StylesManager()
        Get
            Return pStylesManager
        End Get
    End Property

#End Region


End Class