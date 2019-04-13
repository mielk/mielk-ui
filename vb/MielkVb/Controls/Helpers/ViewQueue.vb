Imports System.Collections

Public Class ViewQueue

    Private pDictionary As Dictionary(Of String, QueueItem)

    Public Sub New()
        pDictionary = New Dictionary(Of String, QueueItem)
    End Sub

    Public Sub Add(control As IControl, prop As StylePropertyEnum)
        Dim item As QueueItem
        '--------------------------------------------------------------------------------------------------
        item = New QueueItem(control, prop)
        Call addItemToQueue(item)
    End Sub

    Private Sub addItemToQueue(item As QueueItem)
        Dim key As String
        '--------------------------------------------------------------------------------------------------
        key = item.Key
        If Not pDictionary.ContainsKey(key) Then
            Call pDictionary.Add(key, item)
        End If
    End Sub

    Public Function HasKey(key As String) As Boolean
        Return pDictionary.ContainsKey(key)
    End Function

    Public Sub removeKey(key As String)
        Call pDictionary.Remove(key)
    End Sub

    Public Sub Run()
        Dim item As QueueItem
        '--------------------------------------------------------------------------------------------------
        Do While (pDictionary.Values.Count > 0)
            item = pDictionary.Values(0)
            Call item.Update()
        Loop
    End Sub





    Private Class QueueItem
        Private pControl As IControl
        Private pProperty As StylePropertyEnum

        Public Sub New(control As IControl, prop As StylePropertyEnum)
            pControl = control
            pProperty = prop
        End Sub

        Public Property Control As IControl
            Get
                Return pControl
            End Get
            Set(ByVal value As IControl)
                pControl = value
            End Set
        End Property

        Public Property Prop As StylePropertyEnum
            Get
                Return pProperty
            End Get
            Set(ByVal value As StylePropertyEnum)
                pProperty = value
            End Set
        End Property

        Public ReadOnly Property Key As String
            Get
                Return pControl.GetHashCode & "|" & pProperty
            End Get
        End Property

        Public Sub Update()
            Call pControl.UpdateView(pProperty)
        End Sub

    End Class

End Class
