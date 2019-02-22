Public Class VbaStyle

    Private Const CLASS_NAME As String = "Style"
    '----------------------------------------------------------------------------------------------------------
    Private pParent As VbaStyleSet
    Private pType As CssStyleNodeEnum
    '----------------------------------------------------------------------------------------------------------

    '[Properties] ---------------------------------------------------------------------------------------------
    '[Size]
    Private pWidth As Object                   ' Single / Auto / Inherit
    Private pMinWidth As Object                ' Single / Auto / Inherit
    Private pMaxWidth As Object                ' Single / Auto / Inherit
    Private pHeight As Object                  ' Single / Auto / Inherit
    Private pMinHeight As Object               ' Single / Auto / Inherit
    Private pMaxHeight As Object               ' Single / Auto / Inherit
    '[Layout]
    Private pPosition As Object                ' CssPositionEnum
    Private pFloat As Object                   ' CssFloatEnum
    Private pTop As Object                     ' Single
    Private pBottom As Object                  ' Single
    Private pLeft As Object                    ' Single
    Private pRight As Object                   ' Single
    '[Margin]
    Private pMarginTop As Object               ' Single
    Private pMarginRight As Object             ' Single
    Private pMarginBottom As Object            ' Single
    Private pMarginLeft As Object              ' Single
    Private pMargins As Object
    '[Padding]
    Private pPaddingTop As Object              ' Single
    Private pPaddingRight As Object            ' Single
    Private pPaddingBottom As Object           ' Single
    Private pPaddingLeft As Object             ' Single
    Private pPaddings As Object
    '[Font]
    Private pFontFamily As Object              ' String
    Private pFontSize As Object                ' Single
    Private pFontColor As Object               ' Long
    Private pFontBold As Object                ' Boolean
    Private pHorizontalAlignment As Object     ' MSForms.fmTextAlign
    Private pVerticalAlignment As Object       ' MSForms.fmTextAlign
    '[Background]
    Private pBackgroundColor As Object         ' Long
    '[Border]
    Private pBorderThickness As Object         ' Long
    Private pBorderColor As Object             ' Long
    '[Image]
    Private pImageName As Object               ' String
    Private pImageFilePath As Object           ' String
    Private pImageWidth As Object
    Private pImageHeight As Object
    Private pImageSize As Object
    '----------------------------------------------------------------------------------------------------------


    Public Sub New(ss As VbaStyleSet, node As CssStyleNodeEnum)
        pParent = ss
        pType = node
    End Sub



    '[SETTERS]
    Public Function setParent(value As VbaStyleSet) As VbaStyle
        setParent = Me
        pParent = value
    End Function

    Public Function setType(value As CssStyleNodeEnum) As VbaStyle
        setType = Me
        pType = value
    End Function

    Public Sub setValue(prop As VbaStylePropertyEnum, propertyValue As Object)
        Select Case prop
            Case StylePropertyEnum.StyleProperty_Width : Call setWidth(propertyValue)
            Case StylePropertyEnum.StyleProperty_MinWidth : Call setMinWidth(propertyValue)
            Case StylePropertyEnum.StyleProperty_MaxWidth : Call setMaxWidth(propertyValue)
            Case StylePropertyEnum.StyleProperty_Height : Call setHeight(propertyValue)
            Case StylePropertyEnum.StyleProperty_MinHeight : Call setMinHeight(propertyValue)
            Case StylePropertyEnum.StyleProperty_MaxHeight : Call setMaxHeight(propertyValue)
            Case StylePropertyEnum.StyleProperty_BackgroundColor : Call setBackgroundColor(propertyValue)
            Case StylePropertyEnum.StyleProperty_BorderColor : Call setBorderColor(propertyValue)
            Case StylePropertyEnum.StyleProperty_BorderThickness : Call setBorderThickness(propertyValue)
            Case StylePropertyEnum.StyleProperty_FontFamily : Call setFontFamily(propertyValue)
            Case StylePropertyEnum.StyleProperty_FontSize : Call setFontSize(propertyValue)
            Case StylePropertyEnum.StyleProperty_FontBold : Call setFontBold(propertyValue)
            Case StylePropertyEnum.StyleProperty_FontColor : Call setFontColor(propertyValue)
            Case StylePropertyEnum.StyleProperty_MarginTop : Call setMarginTop(propertyValue)
            Case StylePropertyEnum.StyleProperty_MarginRight : Call setMarginRight(propertyValue)
            Case StylePropertyEnum.StyleProperty_MarginBottom : Call setMarginBottom(propertyValue)
            Case StylePropertyEnum.StyleProperty_MarginLeft : Call setMarginLeft(propertyValue)
            Case StylePropertyEnum.StyleProperty_PaddingTop : Call setPaddingTop(propertyValue)
            Case StylePropertyEnum.StyleProperty_PaddingRight : Call setPaddingRight(propertyValue)
            Case StylePropertyEnum.StyleProperty_PaddingBottom : Call setPaddingBottom(propertyValue)
            Case StylePropertyEnum.StyleProperty_PaddingLeft : Call setPaddingLeft(propertyValue)
            Case StylePropertyEnum.StyleProperty_Float : Call setFloat(propertyValue)
            Case StylePropertyEnum.StyleProperty_Position : Call setPosition(propertyValue)
            Case StylePropertyEnum.StyleProperty_Top : Call setTop(propertyValue)
            Case StylePropertyEnum.StyleProperty_Right : Call setRight(propertyValue)
            Case StylePropertyEnum.StyleProperty_Bottom : Call setBottom(propertyValue)
            Case StylePropertyEnum.StyleProperty_Left : Call setLeft(propertyValue)
            Case StylePropertyEnum.StyleProperty_ImageFilePath : Call setImageFilePath(propertyValue)
            Case StylePropertyEnum.StyleProperty_ImageName : Call setImageName(propertyValue)
            Case StylePropertyEnum.StyleProperty_ImageHeight : Call setImageHeight(propertyValue)
            Case StylePropertyEnum.StyleProperty_ImageWidth : Call setImageWidth(propertyValue)
            Case StylePropertyEnum.StyleProperty_ImageSize : Call setImageSize(propertyValue)
            Case StylePropertyEnum.StyleProperty_HorizontalAlignment : Call setHorizontalAlignment(propertyValue)
            Case StylePropertyEnum.StyleProperty_VerticalAlignment : Call setVerticalAlignment(propertyValue)
            Case StylePropertyEnum.StyleProperty_Padding : Call setPaddings(propertyValue)
            Case StylePropertyEnum.StyleProperty_Margin : Call setMargins(propertyValue)
        End Select
    End Sub

    Public Function setWidth(value As Object) As VbaStyle
        setWidth = Me
        pWidth = value
    End Function
    Public Function setMinWidth(value As Object) As VbaStyle
        setMinWidth = Me
        pMinWidth = value
    End Function
    Public Function setMaxWidth(value As Object) As VbaStyle
        setMaxWidth = Me
        pMaxWidth = value
    End Function
    Public Function setHeight(value As Object) As VbaStyle
        setHeight = Me
        pHeight = value
    End Function
    Public Function setMinHeight(value As Object) As VbaStyle
        setMinHeight = Me
        pMinHeight = value
    End Function
    Public Function setMaxHeight(value As Object) As VbaStyle
        setMaxHeight = Me
        pMaxHeight = value
    End Function
    Public Function setBackgroundColor(value As Object) As VbaStyle
        setBackgroundColor = Me
        pBackgroundColor = value
    End Function
    Public Function setBorderColor(value As Object) As VbaStyle
        setBorderColor = Me
        pBorderColor = value
    End Function
    Public Function setBorderThickness(value As Object) As VbaStyle
        setBorderThickness = Me
        pBorderThickness = value
    End Function
    Public Function setFontFamily(value As Object) As VbaStyle
        setFontFamily = Me
        pFontFamily = value
    End Function
    Public Function setFontSize(value As Object) As VbaStyle
        setFontSize = Me
        pFontSize = CSng(value)
    End Function
    Public Function setFontBold(value As Object) As VbaStyle
        setFontBold = Me
        pFontBold = CBool(value)
    End Function
    Public Function setFontColor(value As Object) As VbaStyle
        setFontColor = Me
        pFontColor = value
    End Function
    Public Function setHorizontalAlignment(value As Object) As VbaStyle
        setHorizontalAlignment = Me
        pHorizontalAlignment = value
    End Function
    Public Function setVerticalAlignment(value As Object) As VbaStyle
        setVerticalAlignment = Me
        pVerticalAlignment = value
    End Function
    Public Function setMarginTop(value As Object) As VbaStyle
        setMarginTop = Me
        pMarginTop = value
    End Function
    Public Function setMarginRight(value As Object) As VbaStyle
        setMarginRight = Me
        pMarginRight = value
    End Function
    Public Function setMarginBottom(value As Object) As VbaStyle
        setMarginBottom = Me
        pMarginTop = value
    End Function
    Public Function setMarginLeft(value As Object) As VbaStyle
        setMarginLeft = Me
        pMarginLeft = value
    End Function
    Public Function setPaddingTop(value As Object) As VbaStyle
        setPaddingTop = Me
        pPaddingTop = value
    End Function
    Public Function setPaddingRight(value As Object) As VbaStyle
        setPaddingRight = Me
        pPaddingRight = value
    End Function
    Public Function setPaddingBottom(value As Object) As VbaStyle
        setPaddingBottom = Me
        pPaddingBottom = value
    End Function
    Public Function setPaddingLeft(value As Object) As VbaStyle
        setPaddingLeft = Me
        pPaddingLeft = value
    End Function
    Public Function setFloat(value As Object) As VbaStyle
        setFloat = Me
        pFloat = value
    End Function
    Public Function setPosition(value As Object) As VbaStyle
        setPosition = Me
        pPosition = CssPositionEnum.CssPosition_Absolute
    End Function
    Public Function setTop(value As Object) As VbaStyle
        setTop = Me
        pTop = value
    End Function
    Public Function setRight(value As Object) As VbaStyle
        setRight = Me
        pRight = value
    End Function
    Public Function setBottom(value As Object) As VbaStyle
        setBottom = Me
        pBottom = value
    End Function
    Public Function setLeft(value As Object) As VbaStyle
        setLeft = Me
        pLeft = value
    End Function
    Public Function setImageFilePath(value As Object) As VbaStyle
        setImageFilePath = Me
        pImageFilePath = value
    End Function
    Public Function setImageName(value As Object) As VbaStyle
        setImageName = Me
        pImageName = value
    End Function
    Public Function setImageWidth(value As Object) As VbaStyle
        setImageWidth = Me
        pImageWidth = value
    End Function
    Public Function setImageHeight(value As Object) As VbaStyle
        setImageHeight = Me
        pImageHeight = value
    End Function
    Public Function setImageSize(value As Object) As VbaStyle
        setImageSize = Me
        pImageSize = value
    End Function
    Public Function setPaddings(value As Object) As VbaStyle
        setPaddings = Me
        pPaddings = value
    End Function
    Public Function setMargins(value As Object) As VbaStyle
        setMargins = Me
        pMargins = value
    End Function


    '[GETTERS]
    Public Function getNodeType() As CssStyleNodeEnum
        getNodeType = pType
    End Function
    Public Function getWidth() As Object
        getWidth = pWidth
    End Function
    Public Function getMinWidth() As Object
        getMinWidth = pMinWidth
    End Function
    Public Function getMaxWidth() As Object
        getMaxWidth = pMaxWidth
    End Function
    Public Function getHeight() As Object
        getHeight = pHeight
    End Function
    Public Function getMinHeight() As Object
        getMinHeight = pMinHeight
    End Function
    Public Function getMaxHeight() As Object
        getMaxHeight = pMaxHeight
    End Function
    Public Function getBackgroundColor() As Object
        getBackgroundColor = pBackgroundColor
    End Function
    Public Function getBorderColor() As Object
        getBorderColor = pBorderColor
    End Function
    Public Function getBorderThickness() As Object
        getBorderThickness = pBorderThickness
    End Function
    Public Function getFontFamily() As Object
        getFontFamily = pFontFamily
    End Function
    Public Function getFontSize() As Object
        getFontSize = pFontSize
    End Function
    Public Function getFontBold() As Object
        getFontBold = pFontBold
    End Function
    Public Function getFontColor() As Object
        getFontColor = pFontColor
    End Function
    Public Function getHorizontalAlignment() As Object
        getHorizontalAlignment = pHorizontalAlignment
    End Function
    Public Function getVerticalAlignment() As Object
        getVerticalAlignment = pVerticalAlignment
    End Function
    Public Function getMarginTop() As Object
        getMarginTop = pMarginTop
    End Function
    Public Function getMarginRight() As Object
        getMarginRight = pMarginRight
    End Function
    Public Function getMarginBottom() As Object
        getMarginBottom = pMarginBottom
    End Function
    Public Function getMarginLeft() As Object
        getMarginLeft = pMarginLeft
    End Function
    Public Function getPaddingTop() As Object
        getPaddingTop = pPaddingTop
    End Function
    Public Function getPaddingRight() As Object
        getPaddingRight = pPaddingRight
    End Function
    Public Function getPaddingBottom() As Object
        getPaddingBottom = pPaddingBottom
    End Function
    Public Function getPaddingLeft() As Object
        getPaddingLeft = pPaddingLeft
    End Function
    Public Function getFloat() As Object
        getFloat = pFloat
    End Function
    Public Function getPosition() As Object
        getPosition = pPosition
    End Function
    Public Function getTop() As Object
        getTop = pTop
    End Function
    Public Function getRight() As Object
        getRight = pRight
    End Function
    Public Function getBottom() As Object
        getBottom = pBottom
    End Function
    Public Function getLeft() As Object
        getLeft = pLeft
    End Function
    Public Function getImageFilePath() As Object
        getImageFilePath = pImageFilePath
    End Function
    Public Function getImageName() As Object
        getImageName = pImageName
    End Function
    Public Function getImageWidth() As Object
        getImageWidth = pImageWidth
    End Function
    Public Function getImageHeight() As Object
        getImageHeight = pImageHeight
    End Function
    Public Function getImageSize() As Object
        getImageSize = pImageSize
    End Function
    Public Function getMargins() As Object
        getMargins = pMargins
    End Function
    Public Function getPaddings() As Object
        getPaddings = pPaddings
    End Function

    Public Function getPropertyAsString(propType As VbaStylePropertyEnum) As Object
        Dim value As Object
        '----------------------------------------------------------------------------------------------------------

        Select Case propType
            '[Position]
            Case VbaStylePropertyEnum.StyleProperty_Float : Return getFloat()
            Case VbaStylePropertyEnum.StyleProperty_Position : Return getPosition()
            Case VbaStylePropertyEnum.StyleProperty_Top : Return getTop()
            Case VbaStylePropertyEnum.StyleProperty_Left : Return getLeft()
            Case VbaStylePropertyEnum.StyleProperty_Bottom : Return getBottom()
            Case VbaStylePropertyEnum.StyleProperty_Right : Return getRight()
                '[Size]
            Case VbaStylePropertyEnum.StyleProperty_Width : Return getWidth()
            Case VbaStylePropertyEnum.StyleProperty_MinWidth : Return getMinWidth()
            Case VbaStylePropertyEnum.StyleProperty_MaxWidth : Return getMaxHeight()
            Case VbaStylePropertyEnum.StyleProperty_Height : Return getHeight()
            Case VbaStylePropertyEnum.StyleProperty_MinHeight : Return getMinHeight()
            Case VbaStylePropertyEnum.StyleProperty_MaxHeight : Return getMaxHeight()
                '[Margins & paddings]
            Case VbaStylePropertyEnum.StyleProperty_Padding : Return getPaddings()
            Case VbaStylePropertyEnum.StyleProperty_PaddingTop : Return getPaddingTop()
            Case VbaStylePropertyEnum.StyleProperty_PaddingLeft : Return getPaddingLeft()
            Case VbaStylePropertyEnum.StyleProperty_PaddingBottom : Return getPaddingBottom()
            Case VbaStylePropertyEnum.StyleProperty_PaddingRight : Return getPaddingRight()
            Case VbaStylePropertyEnum.StyleProperty_Margin : Return getMargins()
            Case VbaStylePropertyEnum.StyleProperty_MarginTop : Return getMarginTop()
            Case VbaStylePropertyEnum.StyleProperty_MarginLeft : Return getMarginLeft()
            Case VbaStylePropertyEnum.StyleProperty_MarginBottom : Return getMarginBottom()
            Case VbaStylePropertyEnum.StyleProperty_MarginRight : Return getMarginRight()
                '[Background & borders]
            Case VbaStylePropertyEnum.StyleProperty_BackgroundColor : Return getBackgroundColor()
            Case VbaStylePropertyEnum.StyleProperty_BorderThickness : Return getBorderThickness()
            Case VbaStylePropertyEnum.StyleProperty_BorderColor : Return getBorderColor()
                '[Font]
            Case VbaStylePropertyEnum.StyleProperty_FontSize : Return getFontSize()
            Case VbaStylePropertyEnum.StyleProperty_FontFamily : Return getFontFamily()
            Case VbaStylePropertyEnum.StyleProperty_FontColor : Return getFontColor()
            Case VbaStylePropertyEnum.StyleProperty_FontBold : Return getFontBold()
                '[Alignment]
            Case VbaStylePropertyEnum.StyleProperty_HorizontalAlignment : Return getHorizontalAlignment()
            Case VbaStylePropertyEnum.StyleProperty_VerticalAlignment : Return getVerticalAlignment()
                '[Images]
            Case VbaStylePropertyEnum.StyleProperty_ImageFilePath : Return getImageFilePath()
            Case VbaStylePropertyEnum.StyleProperty_ImageName : Return getImageName()
            Case VbaStylePropertyEnum.StyleProperty_ImageWidth : Return getImageWidth()
            Case VbaStylePropertyEnum.StyleProperty_ImageHeight : Return getImageHeight()
            Case VbaStylePropertyEnum.StyleProperty_ImageSize : Return getImageSize()
            Case Else : Return vbNullString
        End Select


    End Function

    Public Function getDto() As Dictionary(Of VbaStylePropertyEnum, Object)
        getDto = New Dictionary(Of VbaStylePropertyEnum, Object)
        With getDto
            Call .Add(VbaStylePropertyEnum.StyleProperty_Width, getWidth)
            Call .Add(VbaStylePropertyEnum.StyleProperty_MinWidth, getMinWidth)
            Call .Add(VbaStylePropertyEnum.StyleProperty_MaxWidth, getMaxWidth)
            Call .Add(VbaStylePropertyEnum.StyleProperty_Height, getHeight)
            Call .Add(VbaStylePropertyEnum.StyleProperty_MinHeight, getMinHeight)
            Call .Add(VbaStylePropertyEnum.StyleProperty_MaxHeight, getMaxHeight)
            Call .Add(VbaStylePropertyEnum.StyleProperty_BackgroundColor, getBackgroundColor)
            Call .Add(VbaStylePropertyEnum.StyleProperty_BorderColor, getBorderColor)
            Call .Add(VbaStylePropertyEnum.StyleProperty_BorderThickness, getBorderThickness)
            Call .Add(VbaStylePropertyEnum.StyleProperty_FontFamily, getFontFamily)
            Call .Add(VbaStylePropertyEnum.StyleProperty_FontSize, getFontSize)
            Call .Add(VbaStylePropertyEnum.StyleProperty_FontBold, getFontBold)
            Call .Add(VbaStylePropertyEnum.StyleProperty_FontColor, getFontColor)
            Call .Add(VbaStylePropertyEnum.StyleProperty_MarginTop, getMarginTop)
            Call .Add(VbaStylePropertyEnum.StyleProperty_MarginRight, getMarginRight)
            Call .Add(VbaStylePropertyEnum.StyleProperty_MarginBottom, getMarginBottom)
            Call .Add(VbaStylePropertyEnum.StyleProperty_MarginLeft, getMarginLeft)
            Call .Add(VbaStylePropertyEnum.StyleProperty_PaddingTop, getPaddingTop)
            Call .Add(VbaStylePropertyEnum.StyleProperty_PaddingRight, getPaddingRight)
            Call .Add(VbaStylePropertyEnum.StyleProperty_PaddingBottom, getPaddingBottom)
            Call .Add(VbaStylePropertyEnum.StyleProperty_PaddingLeft, getPaddingLeft)
            Call .Add(VbaStylePropertyEnum.StyleProperty_Float, getFloat)
            Call .Add(VbaStylePropertyEnum.StyleProperty_Position, getPosition)
            Call .Add(VbaStylePropertyEnum.StyleProperty_Top, getTop)
            Call .Add(VbaStylePropertyEnum.StyleProperty_Right, getRight)
            Call .Add(VbaStylePropertyEnum.StyleProperty_Bottom, getBottom)
            Call .Add(VbaStylePropertyEnum.StyleProperty_Left, getLeft)
            Call .Add(VbaStylePropertyEnum.StyleProperty_HorizontalAlignment, getHorizontalAlignment)
            Call .Add(VbaStylePropertyEnum.StyleProperty_VerticalAlignment, getVerticalAlignment)
            Call .Add(VbaStylePropertyEnum.StyleProperty_ImageFilePath, getImageFilePath)
            Call .Add(VbaStylePropertyEnum.StyleProperty_ImageName, getImageName)
            Call .Add(VbaStylePropertyEnum.StyleProperty_ImageHeight, getImageHeight)
            Call .Add(VbaStylePropertyEnum.StyleProperty_ImageWidth, getImageWidth)
            Call .Add(VbaStylePropertyEnum.StyleProperty_ImageSize, getImageSize)
            Call .Add(VbaStylePropertyEnum.StyleProperty_Margin, getMargins)
            Call .Add(VbaStylePropertyEnum.StyleProperty_Padding, getPaddings)
        End With
    End Function





    '[ADDING PROPERTIES]
    Public Sub addProperty(propertyEnum As VbaStylePropertyEnum, propertyValue As String)
        Call setValue(propertyEnum, propertyValue)
    End Sub

    Public Sub addProperties(dict As Dictionary(Of VbaStylePropertyEnum, String))
        Dim varKey As Object
        Dim propertyEnum As VbaStylePropertyEnum
        Dim propertyValue As String
        '----------------------------------------------------------------------------------------------------------

        For Each varKey In dict.Keys
            propertyEnum = varKey
            propertyValue = dict.Item(varKey)
            Call setValue(propertyEnum, propertyValue)
        Next varKey

    End Sub


    'Public Function merge(entry As VbaStyle) As VbaStyle
    '    Dim dto As Scripting.Dictionary
    '    Dim varKey As Variant
    '    Dim value As Variant
    '    '----------------------------------------------------------------------------------------------------------
    '    Set dto = entry.getDto
    '    For Each varKey In dto.Keys
    '        value = dto.item(varKey)
    '        If Not MielkCore.isUndefined(value) Then
    '            Call setValue(VBA.CLng(varKey), value)
    '        End If
    '    Next varKey
    '    Set merge = Me
    'End Function
    '
    'Public Function clone() As VbaStyle
    '    Dim dto As Scripting.Dictionary
    '    Dim varKey As Variant
    '    Dim value As Variant
    '    '----------------------------------------------------------------------------------------------------------
    '    Set dto = getDto
    '    Set clone = New Style
    '    For Each varKey In dto.Keys
    '        value = dto.item(varKey)
    '        If Not MielkCore.isUndefined(value) Then
    '            Call clone.setValue(VBA.CLng(varKey), value)
    '        End If
    '    Next varKey
    'End Function
    '

End Class
