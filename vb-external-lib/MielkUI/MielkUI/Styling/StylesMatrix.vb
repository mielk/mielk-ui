Imports System.Drawing

Public Class StylesMatrix

#Region "Private variables"
    Private pParent As IControl
    Private pElementStyleSet As StyleSet
    Private pIdStyleSet As StyleSet
    Private pClassStyleSets As List(Of StyleSet)
    Private pInlineStyle As StyleSet
    Private pPropertiesArraysMap(0 To CountStyleNodeTypesEnums(), 0 To CountStylePropertiesEnums()) As Object
    Private pLog As String
#End Region


#Region "Constructor"

    Public Sub New(parent As IControl, Optional runLoadingStyles As Boolean = True)
        pParent = parent
        pInlineStyle = New StyleSet(vbNullString, vbNullString, vbNullString)
        If runLoadingStyles Then LoadStyles()
    End Sub

    Public Sub LoadStyles()
        Call addElementStyleSet()
        Call addIdStyleSet()
        Call updatePropertiesArrayMap()
    End Sub

#End Region




#Region "Debugging"

    Public Function hasElementStyleSet() As Boolean
        Return Not pElementStyleSet Is Nothing
    End Function

    Public Function getElementStyleSet() As StyleSet
        Return pElementStyleSet
    End Function

    Public Function getLog() As String
        Return pLog
    End Function

#End Region




#Region "Managing styles"

    Private Sub addElementStyleSet()
        Dim tagName As String
        '------------------------------------------------------------
        tagName = pParent.GetElementTag
        pElementStyleSet = pParent.GetStylesManager().GetStyleByElementTag(tagName)
    End Sub

    Private Sub addIdStyleSet()
        Dim tagName As String
        '------------------------------------------------------------
        tagName = pParent.GetElementTag
        pIdStyleSet = pParent.GetStylesManager().GetStyleByElementTag(tagName)
    End Sub

    Private Sub addStyleSetToArray(styleSet As StyleSet)
        If Not styleSet Is Nothing Then
            If pClassStyleSets Is Nothing Then
                pClassStyleSets.Add(styleSet)
            End If
        End If
    End Sub



    Private Function createArrayForPropertyValuesStorage() As Object()
        Dim propertiesCounter As Integer
        '------------------------------------------------------------
        propertiesCounter = [Enum].GetValues(GetType(StylePropertyEnum)).Cast(Of Integer).Max
        Dim arr(0 To propertiesCounter) As Object
        Return arr
    End Function

    Private Sub updatePropertiesArrayMap()
        Dim styleSet As StyleSet
        '------------------------------------------------------------

        Call Array.Clear(pPropertiesArraysMap, 0, pPropertiesArraysMap.Length)

        If Not pElementStyleSet Is Nothing Then
            Call applySingleStyleToPropertiesArray(pElementStyleSet)
        End If

        'For Each styleSet In pClassStyleSets
        '    Call applySingleStyleToPropertiesArray(styleSet)
        'Next

        'If Not pIdStyleSet Is Nothing Then
        '    Call applySingleStyleToPropertiesArray(pIdStyleSet)
        'End If

        'If Not pInlineStyle Is Nothing Then
        '    Call applySingleStyleToPropertiesArray(pInlineStyle)
        'End If

    End Sub



    Private Sub applySingleStyleToPropertiesArray(styleSet As StyleSet)
        Dim style As Style
        '------------------------------------------------------------

        If Not styleSet Is Nothing Then

            style = styleSet.getStyle(StyleNodeTypeEnum.StyleNodeType_Normal)
            If Not style Is Nothing Then
                pPropertiesArraysMap(StyleNodeTypeEnum.StyleNodeType_Normal, 0) = True
                Call setPropertyValue(StyleNodeTypeEnum.StyleNodeType_Normal, StylePropertyEnum.StyleProperty_BackgroundColor, style.BackgroundColor)
            End If

            style = styleSet.getStyle(StyleNodeTypeEnum.StyleNodeType_Hover)
            If Not style Is Nothing Then
                pPropertiesArraysMap(StyleNodeTypeEnum.StyleNodeType_Hover, 0) = True
                Call setPropertyValue(StyleNodeTypeEnum.StyleNodeType_Hover, StylePropertyEnum.StyleProperty_BackgroundColor, style.BackgroundColor)
            End If

            '(...) pozostałe eventy tak samo

        End If

    End Sub

    Private Sub setPropertyValue(nodeType As StyleNodeTypeEnum, propertyType As StylePropertyEnum, value As StylePropertyValue)
        If Not value Is Nothing Then
            If Not value.IsNull Then
                pPropertiesArraysMap(nodeType, propertyType) = value.Value
            End If
        End If
    End Sub


#End Region








#Region "Public methods"

    Public Sub AddStyleClass(className As String)
        'GetStylesManager().GetStyleByElementAndClass(TAG_NAME, name)
    End Sub

    Public Sub RemoveStyleClass(className As String)

    End Sub

    Public Sub AddInlineStyle(propertyType As Long, value As Object)

    End Sub

#End Region


#Region "Access to styles definitions"

    Public Function IsNodeStyleDefined(nodeType As StyleNodeTypeEnum) As Boolean
        Return pPropertiesArraysMap(nodeType, 0)
    End Function

    Public Function GetPropertyValue(propertyType As StylePropertyEnum, nodeType As StyleNodeTypeEnum) As Object
        Return pPropertiesArraysMap(nodeType, propertyType)
    End Function

    Public Function GetBackgroundColor(nodeType As StyleNodeTypeEnum) As Color
        Return pPropertiesArraysMap(nodeType, StylePropertyEnum.StyleProperty_BackgroundColor)
    End Function

#End Region


End Class