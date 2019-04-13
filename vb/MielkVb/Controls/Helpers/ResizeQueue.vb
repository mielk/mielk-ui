Imports System.Collections

Public Class ResizeQueue

    Private pDictionary As Dictionary(Of String, QueueItem)

    Public Sub New()
        pDictionary = New Dictionary(Of String, QueueItem)
    End Sub

    Public Sub AddResizeTask(control As IControl, prop As StylePropertyEnum)
        Dim item As QueueItem
        '--------------------------------------------------------------------------------------------------
        item = New QueueItem(control, prop)
        Call addItemToQueue(item)
    End Sub

    Public Sub AddLayoutTask(control As IControl)
        Dim item As QueueItem
        '--------------------------------------------------------------------------------------------------
        item = New QueueItem(control, True)
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
        Private pContainer As IContainer
        Private pProperty As StylePropertyEnum
        Private pIsLayoutEvent As Boolean

        Public Sub New(control As IControl, prop As StylePropertyEnum)
            pControl = control
            pProperty = prop
        End Sub

        Public Sub New(control As IControl, isLayoutEvent As Boolean)
            pControl = control
            pContainer = control
            pIsLayoutEvent = isLayoutEvent
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
                If pIsLayoutEvent Then
                    Return pContainer.GetHashCode & "|Layout"
                Else
                    Return pControl.GetHashCode & "|" & pProperty
                End If
            End Get
        End Property

        Public Sub Update()
            If pIsLayoutEvent Then
                Call pContainer.ArrangeControls()
            ElseIf pProperty = StylePropertyEnum.StyleProperty_Width Then
                Call pControl.UpdateWidth()
            ElseIf pProperty = StylePropertyEnum.StyleProperty_Height Then
                Call pControl.UpdateHeight()
            End If
        End Sub

    End Class

End Class