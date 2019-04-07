Imports System.Collections

Public Class ResizeQueue

    'Private pHeadItem As QueueItem
    'Private pTailItem As QueueItem
    Private pDictionary As Dictionary(Of String, QueueItem)

    'Private pChecked As Dictionary(Of String, QueueItem)
    'Private pNotChecked As Array

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
            'If pTailItem Is Nothing Then
            '    pHeadItem = item
            '    pTailItem = item
            'Else
            '    pTailItem.NextItem = item
            '    item.PreviousItem = pTailItem
            '    pTailItem = item
            'End If
        End If

    End Sub

    Public Function HasKey(key As String) As Boolean
        Return pDictionary.ContainsKey(key)
    End Function

    Public Sub removeKey(key As String)
        Call pDictionary.Remove(key)
    End Sub













    Public Sub Run()
        'Call sortItems()
        'Stop
    End Sub

    'Private Sub sortItems()
    '    Dim item As QueueItem
    '    '--------------------------------------------------------------------------------------------------

    '    pChecked = New Dictionary(Of String, QueueItem)
    '    pNotChecked = pDictionary.Values.ToArray

    '    For Each item In pNotChecked
    '        Call processItemSorting(item)
    '    Next

    'End Sub

    'Private Sub processItemSorting(item As QueueItem)
    '    Dim key As String : key = item.Key
    '    Dim prop As StylePropertyEnum : prop = item.Prop
    '    Dim ctrl As IControl : ctrl = item.Control
    '    Dim sizeType As ControlSizeTypeEnum
    '    Dim child As IControl
    '    '--------------------------------------------------------------------------------------------------

    '    If Not pChecked.ContainsKey(key) Then
    '        sizeType = ctrl.GetSizeType(prop)
    '        If sizeType = ControlSizeTypeEnum.ControlSizeType_Const Then

    '        ElseIf sizeType = ControlSizeTypeEnum.ControlSizeType_Auto Then

    '        ElseIf sizeType = ControlSizeTypeEnum.ControlSizeType_ParentRelative Then

    '        End If
    '        Call pChecked.Add(key, item)
    '    End If

    'End Sub



    Private Class QueueItem
        Private pControl As IControl
        Private pProperty As StylePropertyEnum
        Private pIsLayoutEvent As Boolean
        Private pPrevious As QueueItem
        Private pNext As QueueItem

        Public Sub New(control As IControl, prop As StylePropertyEnum)
            pControl = control
            pProperty = prop
        End Sub

        Public Sub New(control As IControl, isLayoutEvent As Boolean)
            pControl = control
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

        Public Property PreviousItem As QueueItem
            Get
                Return pPrevious
            End Get
            Set(ByVal value As QueueItem)
                pPrevious = value
            End Set
        End Property

        Public Property NextItem As QueueItem
            Get
                Return pNext
            End Get
            Set(ByVal value As QueueItem)
                pNext = value
            End Set
        End Property

        Public ReadOnly Property Key As String
            Get
                If pIsLayoutEvent Then
                    Return pControl.GetHashCode & "|Layout"
                Else
                    Return pControl.GetHashCode & "|" & pProperty
                End If
            End Get
        End Property

    End Class

End Class