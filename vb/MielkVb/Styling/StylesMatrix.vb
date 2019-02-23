Imports System.Drawing

Public Class StylesMatrix

#Region "Private variables"
    Private pParent As IControl
    Private pElementStyleSet As StyleSet
    Private pIdStyleSet As StyleSet
    Private pClassStyleSets As List(Of StyleSet)
    Private pInlineStyle As StyleSet
    Private pPropertiesArraysMap(0 To CountStyleNodeTypesEnums(), 0 To CountStylePropertiesEnums()) As Object
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

    Private Sub updatePropertiesArrayMap()

        Call Array.Clear(pPropertiesArraysMap, 0, pPropertiesArraysMap.Length)
        If Not pElementStyleSet Is Nothing Then Call applySingleStyleToPropertiesArray(pElementStyleSet)

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
        If Not styleSet Is Nothing Then
            Call processSingleStyleNodeType(styleSet, StyleNodeTypeEnum.StyleNodeType_Normal)
            Call processSingleStyleNodeType(styleSet, StyleNodeTypeEnum.StyleNodeType_Hover)
            Call processSingleStyleNodeType(styleSet, StyleNodeTypeEnum.StyleNodeType_Clicked)
            Call processSingleStyleNodeType(styleSet, StyleNodeTypeEnum.StyleNodeType_Checked)
            Call processSingleStyleNodeType(styleSet, StyleNodeTypeEnum.StyleNodeType_Disabled)
        End If
    End Sub

    Private Sub processSingleStyleNodeType(styleSet As StyleSet, nodeType As StyleNodeTypeEnum)
        Dim style As Style
        '------------------------------------------------------------
        style = styleSet.getStyle(nodeType)
        If Not style Is Nothing Then
            pPropertiesArraysMap(nodeType, 0) = True
            Call setPropertyValue(nodeType, StylePropertyEnum.StyleProperty_Float, style.Float)
            Call setPropertyValue(nodeType, StylePropertyEnum.StyleProperty_Position, style.Position)
            Call setPropertyValue(nodeType, StylePropertyEnum.StyleProperty_BorderBox, style.BorderBox)
            Call setPropertyValue(nodeType, StylePropertyEnum.StyleProperty_Top, style.Top)
            Call setPropertyValue(nodeType, StylePropertyEnum.StyleProperty_Left, style.Left)
            Call setPropertyValue(nodeType, StylePropertyEnum.StyleProperty_Bottom, style.Bottom)
            Call setPropertyValue(nodeType, StylePropertyEnum.StyleProperty_Right, style.Right)
            Call setPropertyValue(nodeType, StylePropertyEnum.StyleProperty_Width, style.Width)
            Call setPropertyValue(nodeType, StylePropertyEnum.StyleProperty_MinWidth, style.MinWidth)
            Call setPropertyValue(nodeType, StylePropertyEnum.StyleProperty_MaxWidth, style.MaxWidth)
            Call setPropertyValue(nodeType, StylePropertyEnum.StyleProperty_Height, style.Height)
            Call setPropertyValue(nodeType, StylePropertyEnum.StyleProperty_MinHeight, style.MinHeight)
            Call setPropertyValue(nodeType, StylePropertyEnum.StyleProperty_MaxHeight, style.MaxHeight)
            Call setPropertyValue(nodeType, StylePropertyEnum.StyleProperty_Padding, style.Paddings)
            Call setPropertyValue(nodeType, StylePropertyEnum.StyleProperty_PaddingTop, style.PaddingTop)
            Call setPropertyValue(nodeType, StylePropertyEnum.StyleProperty_PaddingLeft, style.PaddingLeft)
            Call setPropertyValue(nodeType, StylePropertyEnum.StyleProperty_PaddingBottom, style.PaddingBottom)
            Call setPropertyValue(nodeType, StylePropertyEnum.StyleProperty_PaddingRight, style.PaddingRight)
            Call setPropertyValue(nodeType, StylePropertyEnum.StyleProperty_Margin, style.Margins)
            Call setPropertyValue(nodeType, StylePropertyEnum.StyleProperty_MarginTop, style.MarginTop)
            Call setPropertyValue(nodeType, StylePropertyEnum.StyleProperty_MarginLeft, style.MarginLeft)
            Call setPropertyValue(nodeType, StylePropertyEnum.StyleProperty_MarginBottom, style.MarginBottom)
            Call setPropertyValue(nodeType, StylePropertyEnum.StyleProperty_MarginRight, style.MarginRight)
            Call setPropertyValue(nodeType, StylePropertyEnum.StyleProperty_BackgroundColor, style.BackgroundColor)
            Call setPropertyValue(nodeType, StylePropertyEnum.StyleProperty_BorderThickness, style.BorderThickness)
            Call setPropertyValue(nodeType, StylePropertyEnum.StyleProperty_BorderColor, style.BorderColor)
            Call setPropertyValue(nodeType, StylePropertyEnum.StyleProperty_FontSize, style.FontSize)
            Call setPropertyValue(nodeType, StylePropertyEnum.StyleProperty_FontFamily, style.FontFamily)
            Call setPropertyValue(nodeType, StylePropertyEnum.StyleProperty_FontColor, style.FontColor)
            Call setPropertyValue(nodeType, StylePropertyEnum.StyleProperty_FontBold, style.FontBold)
            Call setPropertyValue(nodeType, StylePropertyEnum.StyleProperty_HorizontalAlignment, style.HorizontalAlignment)
            Call setPropertyValue(nodeType, StylePropertyEnum.StyleProperty_VerticalAlignment, style.VerticalAlignment)
            Call setPropertyValue(nodeType, StylePropertyEnum.StyleProperty_ImageFilePath, style.ImageFilePath)
            Call setPropertyValue(nodeType, StylePropertyEnum.StyleProperty_ImageName, style.ImageName)
            Call setPropertyValue(nodeType, StylePropertyEnum.StyleProperty_ImageWidth, style.ImageWidth)
            Call setPropertyValue(nodeType, StylePropertyEnum.StyleProperty_ImageHeight, style.ImageHeight)
            Call setPropertyValue(nodeType, StylePropertyEnum.StyleProperty_ImageSize, style.ImageSize)
        Else
            pPropertiesArraysMap(nodeType, 0) = False
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

    Public Function GetPropertiesArray() As Object
        Return pPropertiesArraysMap
    End Function

#End Region


End Class