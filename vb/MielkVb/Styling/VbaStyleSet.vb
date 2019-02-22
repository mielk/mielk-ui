Public Class VbaStyleSet
    Private pElementTag As String
    Private pClassName As String
    Private pId As String
    Private pNodes As Dictionary(Of CssStyleNodeEnum, VbaStyle)



    '[INITIALIZER]
    Public Sub New()
        Call initializeContainers()
        Call initializeDefaultValues()
        Call initializeObjects()
    End Sub

    Private Sub initializeContainers()
        pNodes = New Dictionary(Of CssStyleNodeEnum, VbaStyle)
    End Sub

    Private Sub initializeDefaultValues()
    End Sub

    Private Sub initializeObjects()
    End Sub



    '[SETTERS]
    Public Sub setElement(value As String)
        pElementTag = value
    End Sub

    Public Sub setClass(value As String)
        pClassName = value
    End Sub

    Public Sub setId(value As String)
        pId = value
    End Sub



    '[GETTERS]
    Public Function getKey() As String
        getKey = pElementTag & "|" & pClassName & "|" & pId
    End Function

    Public Function getElementTag() As String
        getElementTag = pElementTag
    End Function

    Public Function getClassName() As String
        getClassName = pClassName
    End Function

    Public Function getId() As String
        getId = pId
    End Function

    Public Function GetStyleType() As CssStyleTypeEnum
        If pId.Length > 0 Then
            Return CssStyleTypeEnum.CssStyleTypeEnum_Id
        ElseIf pClassName.Length > 0 Then
            Return IIf(pElementTag.Length > 0, CssStyleTypeEnum.CssStyleTypeEnum_ElementClass, CssStyleTypeEnum.CssStyleTypeEnum_Class)
        ElseIf pElementTag.Length > 0 Then
            Return IIf(pClassName.Length > 0, CssStyleTypeEnum.CssStyleTypeEnum_ElementClass, CssStyleTypeEnum.CssStyleTypeEnum_Element)
        Else
            Return CssStyleTypeEnum.CssStyleTypeEnum_Id
        End If
    End Function




    '[PSEUDO CLASSES]
    Public Function getNode(nodeType As CssStyleNodeEnum) As VbaStyle
        If pNodes.ContainsKey(nodeType) Then
            Return pNodes(nodeType)
        Else
            Return Nothing
        End If
    End Function

    Public Function getNodes() As Dictionary(Of CssStyleNodeEnum, VbaStyle)
        getNodes = pNodes
    End Function


    '[LOADING FROM FILE]
    Public Function hasNode(nodeType As CssStyleNodeEnum) As Boolean
        hasNode = pNodes.ContainsKey(nodeType)
    End Function

    Public Sub addNode(nodeType As CssStyleNodeEnum, node As VbaStyle)
        Call pNodes.Add(nodeType, node)
    End Sub

    Private Function DictEntry_Key() As Integer
        Throw New NotImplementedException
    End Function


End Class
