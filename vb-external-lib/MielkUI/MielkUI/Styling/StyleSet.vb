Public Class StyleSet

#Region "Private variables"
    Private pElementTag As String
    Private pClassName As String
    Private pId As String
    '------------------------------------------------------------------------------------
    Private pStyleNormal As Style
    Private pStyleHover As Style
    Private pStyleClicked As Style
    Private pStyleChecked As Style
    Private pStyleDisabled As Style
#End Region


#Region "Constructor"

    Public Sub New(ByVal elementTag As String, ByVal className As String, ByVal id As String)
        pElementTag = elementTag
        pClassName = className
        pId = id
    End Sub

    Public Shared Function CreateInstanceFromVbaObject(vbaObject As Object)
        Dim instance As StyleSet
        '------------------------------------------------------------------------------------------------------------------

        With vbaObject
            instance = New StyleSet(.getElementTag, .getClassName, .getId)
        End With
        Call instance.loadStylesFromVbaObject(vbaObject)

        Return instance

    End Function

    Private Sub loadStylesFromVbaObject(vbaObject As Object)
        Dim vbaNode As Object
        '------------------------------------------------------------------------------------------------------------------

        'Load nodes
        For Each nodeType As StyleNodeTypeEnum In System.Enum.GetValues(GetType(StyleNodeTypeEnum))
            vbaNode = vbaObject.GetNode(nodeType)
            If Not vbaNode Is Nothing Then
                Call AddStyle(Style.FromVbaObject(vbaNode))
            End If
        Next

    End Sub

#End Region



#Region "Public methods"

    Public Sub AddStyle(style As Style)
        Select Case style.Type
            Case StyleNodeTypeEnum.StyleNodeType_Normal : pStyleNormal = style
            Case StyleNodeTypeEnum.StyleNodeType_Hover : pStyleHover = style
            Case StyleNodeTypeEnum.StyleNodeType_Clicked : pStyleClicked = style
            Case StyleNodeTypeEnum.StyleNodeType_Checked : pStyleChecked = style
            Case StyleNodeTypeEnum.StyleNodeType_Disabled : pStyleDisabled = style
        End Select
    End Sub

#End Region


#Region "Public properties"

    Public Function getKey() As String
        Return pElementTag & "|" & pClassName & "|" & pId
    End Function

    Public Function getId() As String
        Return pId
    End Function

    Public Function getElementTag() As String
        Return pElementTag
    End Function

    Public Function getClassName() As String
        Return pClassName
    End Function

    Public Function getStyleType() As StyleTypeEnum
        If Len(pId) > 0 Then
            Return StyleTypeEnum.StyleType_Id
        ElseIf Len(pClassName) > 0 Then
            Return IIf(Len(pElementTag) > 0, StyleTypeEnum.StyleType_ElementClass, StyleTypeEnum.StyleType_Class)
        ElseIf Len(pElementTag) > 0 Then
            Return IIf(Len(pClassName) > 0, StyleTypeEnum.StyleType_ElementClass, StyleTypeEnum.StyleType_Element)
        Else
            Return StyleTypeEnum.StyleType_Element
        End If
    End Function

    Public Function getProperty(propertyType As StylePropertyEnum, nodeType As StyleNodeTypeEnum) As StylePropertyValue
        Dim style As Style
        '------------------------------------------------------------------------------------------------------------------

        style = getStyle(nodeType)
        If Not style Is Nothing Then
            Return style.getProperty(propertyType)
        Else
            Return Nothing
        End If

    End Function

    Public Function getStyle(nodeType As StyleNodeTypeEnum) As Style
        Select Case nodeType
            Case StyleNodeTypeEnum.StyleNodeType_Normal : Return pStyleNormal
            Case StyleNodeTypeEnum.StyleNodeType_Hover : Return pStyleHover
            Case StyleNodeTypeEnum.StyleNodeType_Clicked : Return pStyleClicked
            Case StyleNodeTypeEnum.StyleNodeType_Checked : Return pStyleChecked
            Case StyleNodeTypeEnum.StyleNodeType_Disabled : Return pStyleDisabled
            Case Else : Return Nothing
        End Select
    End Function

#End Region


End Class