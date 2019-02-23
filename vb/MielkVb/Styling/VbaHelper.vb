Module VbaHelper

    Public Function convertStringToNodeType(text As String) As CssStyleNodeEnum
        If Not text Is Nothing Then
            Select Case text.ToLower
                Case vbNullString : Return CssStyleNodeEnum.CssStyleNode_Normal
                Case "hover" : Return CssStyleNodeEnum.CssStyleNode_Hover
                Case "clicked" : Return CssStyleNodeEnum.CssStyleNode_Clicked
                Case "checked" : Return CssStyleNodeEnum.CssStyleNode_Checked
                Case "disabled" : Return CssStyleNodeEnum.CssStyleNode_Disabled
            End Select
        End If
        Return CssStyleNodeEnum.CssStyleNode_Normal
    End Function

    Public Function convertStringToStyleProperty(str As String) As StylePropertyEnum
        Select Case str.ToLower.Trim
            Case "width" : Return StylePropertyEnum.StyleProperty_Width
            Case "min-width" : Return StylePropertyEnum.StyleProperty_MinWidth
            Case "max-width" : Return StylePropertyEnum.StyleProperty_MaxWidth
            Case "height" : Return StylePropertyEnum.StyleProperty_Height
            Case "min-height" : Return StylePropertyEnum.StyleProperty_MinHeight
            Case "max-height" : Return StylePropertyEnum.StyleProperty_MaxHeight
            Case "background-color" : Return StylePropertyEnum.StyleProperty_BackgroundColor
            Case "border-color" : Return StylePropertyEnum.StyleProperty_BorderColor
            Case "border-thickness" : Return StylePropertyEnum.StyleProperty_BorderThickness
            Case "font-family" : Return StylePropertyEnum.StyleProperty_FontFamily
            Case "font-size" : Return StylePropertyEnum.StyleProperty_FontSize
            Case "font-bold" : Return StylePropertyEnum.StyleProperty_FontBold
            Case "font-color" : Return StylePropertyEnum.StyleProperty_FontColor
            Case "margin-top" : Return StylePropertyEnum.StyleProperty_MarginTop
            Case "margin-right" : Return StylePropertyEnum.StyleProperty_MarginRight
            Case "margin-bottom" : Return StylePropertyEnum.StyleProperty_MarginBottom
            Case "margin-left" : Return StylePropertyEnum.StyleProperty_MarginLeft
            Case "padding-top" : Return StylePropertyEnum.StyleProperty_PaddingTop
            Case "padding-right" : Return StylePropertyEnum.StyleProperty_PaddingRight
            Case "padding-bottom" : Return StylePropertyEnum.StyleProperty_PaddingBottom
            Case "padding-left" : Return StylePropertyEnum.StyleProperty_PaddingLeft
            Case "float" : Return StylePropertyEnum.StyleProperty_Float
            Case "position" : Return StylePropertyEnum.StyleProperty_Position
            Case "border-box" : Return StylePropertyEnum.StyleProperty_BorderBox
            Case "top" : Return StylePropertyEnum.StyleProperty_Top
            Case "right" : Return StylePropertyEnum.StyleProperty_Right
            Case "bottom" : Return StylePropertyEnum.StyleProperty_Bottom
            Case "left" : Return StylePropertyEnum.StyleProperty_Left
            Case "image-file-path" : Return StylePropertyEnum.StyleProperty_ImageFilePath
            Case "image-name" : Return StylePropertyEnum.StyleProperty_ImageName
            Case "image-height" : Return StylePropertyEnum.StyleProperty_ImageHeight
            Case "image-width" : Return StylePropertyEnum.StyleProperty_ImageWidth
            Case "image-size" : Return StylePropertyEnum.StyleProperty_ImageSize
            Case "horizontal-alignment" : Return StylePropertyEnum.StyleProperty_HorizontalAlignment
            Case "vertical-alignment" : Return StylePropertyEnum.StyleProperty_VerticalAlignment
            Case "padding" : Return StylePropertyEnum.StyleProperty_Padding
            Case "margin" : Return StylePropertyEnum.StyleProperty_Margin
        End Select
    End Function

End Module
