﻿Imports System.Drawing

Public Class StylePropertyValue
    Private pIsNull As Boolean
    Private pIsInherited As Boolean
    Private pIsAuto As Boolean
    Private pCssValue As String
    Private pRealValue As Object

#Region "Constructors"

    Public Shared Function Null() As StylePropertyValue
        Dim instance As StylePropertyValue
        instance = New StylePropertyValue
        instance.IsNull = True
        Return instance
    End Function

    Public Shared Function Inherited() As StylePropertyValue
        Dim instance As StylePropertyValue
        instance = New StylePropertyValue
        instance.IsInherited = True
        Return instance
    End Function

    Public Shared Function Auto() As StylePropertyValue
        Dim instance As StylePropertyValue
        instance = New StylePropertyValue
        instance.IsAuto = True
        Return instance
    End Function

    Public Sub New()

    End Sub

    Public Sub New(value As String)
        pCssValue = value
    End Sub

#End Region


#Region "Public properties"

    Public Property IsNull() As Boolean
        Get
            Return pIsNull
        End Get
        Set(value As Boolean)
            pIsNull = value
        End Set
    End Property

    Public Property IsInherited() As Boolean
        Get
            Return pIsInherited
        End Get
        Set(value As Boolean)
            pIsInherited = value
        End Set
    End Property

    Public Property IsAuto() As Boolean
        Get
            Return pIsAuto
        End Get
        Set(value As Boolean)
            pIsAuto = value
        End Set
    End Property

    Public Property CssValue() As Object
        Get
            Return pCssValue
        End Get
        Set(v As Object)
            pCssValue = v
        End Set
    End Property

    Public Property Value() As Object
        Get
            Return pRealValue
        End Get
        Set(v As Object)
            pRealValue = v
        End Set
    End Property

    'Public Property Color() As Color
    '    Get
    '        Return pColor
    '    End Get
    '    Set(v As Color)
    '        pColor = v
    '    End Set
    'End Property

#End Region

End Class