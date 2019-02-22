Imports System.Drawing

Public Class Style

#Region "Private variables"

    Private pType As StyleNodeTypeEnum

    '[Position]
    Private pFloat As StylePropertyValue
    Private pPosition As StylePropertyValue
    Private pTop As StylePropertyValue
    Private pLeft As StylePropertyValue
    Private pBottom As StylePropertyValue
    Private pRight As StylePropertyValue

    '[Size]
    Private pWidth As StylePropertyValue
    Private pMinWidth As StylePropertyValue
    Private pMaxWidth As StylePropertyValue
    Private pHeight As StylePropertyValue
    Private pMinHeight As StylePropertyValue
    Private pMaxHeight As StylePropertyValue

    '[Margins and paddings]
    Private pMargins As StylePropertyValue
    Private pMarginTop As StylePropertyValue
    Private pMarginLeft As StylePropertyValue
    Private pMarginBottom As StylePropertyValue
    Private pMarginRight As StylePropertyValue
    Private pPaddings As StylePropertyValue
    Private pPaddingTop As StylePropertyValue
    Private pPaddingLeft As StylePropertyValue
    Private pPaddingBottom As StylePropertyValue
    Private pPaddingRight As StylePropertyValue

    '[Background & Border]
    Private pBackgroundColor As StylePropertyValue
    Private pBorderColor As StylePropertyValue
    Private pBorderThickness As StylePropertyValue

    '[Alignments]
    Private pHorizontalAlignment As StylePropertyValue
    Private pVerticalAlignment As StylePropertyValue

    '[Font]
    Private pFontSize As StylePropertyValue
    Private pFontFamily As StylePropertyValue
    Private pFontBold As StylePropertyValue
    Private pFontColor As StylePropertyValue

    '[Image]
    Private pImageFilePath As StylePropertyValue
    Private pImageName As StylePropertyValue
    Private pImageWidth As StylePropertyValue
    Private pImageHeight As StylePropertyValue
    Private pImageSize As StylePropertyValue


#End Region


#Region "Constructor"

    Public Sub New(type As StyleNodeTypeEnum)
        pType = type
    End Sub

    Public Shared Function FromVbaObject(vbaObject As Object)
        Dim instance As Style
        '----------------------------------------------------------------------------------------------------------------
        If Not vbaObject Is Nothing Then
            instance = New Style(vbaObject.getNodeType)
            Call instance.loadPropertiesFromVbaObject(vbaObject)
            Return instance
        End If
        Return Nothing
    End Function

#End Region


#Region "Loading properties from VBA"

    Private Sub loadPropertiesFromVbaObject(vbaObject As Object)
        Float = convertValue(vbaObject.GetPropertyAsString(StylePropertyEnum.StyleProperty_Float))
        Position = convertValue(vbaObject.GetPropertyAsString(StylePropertyEnum.StyleProperty_Position))
        Top = convertValue(vbaObject.GetPropertyAsString(StylePropertyEnum.StyleProperty_Top))
        Left = convertValue(vbaObject.GetPropertyAsString(StylePropertyEnum.StyleProperty_Left))
        Bottom = convertValue(vbaObject.GetPropertyAsString(StylePropertyEnum.StyleProperty_Bottom))
        Right = convertValue(vbaObject.GetPropertyAsString(StylePropertyEnum.StyleProperty_Right))
        Width = convertValue(vbaObject.GetPropertyAsString(StylePropertyEnum.StyleProperty_Width))
        MinWidth = convertValue(vbaObject.GetPropertyAsString(StylePropertyEnum.StyleProperty_MinWidth))
        MaxWidth = convertValue(vbaObject.GetPropertyAsString(StylePropertyEnum.StyleProperty_MaxWidth))
        Height = convertValue(vbaObject.GetPropertyAsString(StylePropertyEnum.StyleProperty_Height))
        MinHeight = convertValue(vbaObject.GetPropertyAsString(StylePropertyEnum.StyleProperty_MinHeight))
        MaxHeight = convertValue(vbaObject.GetPropertyAsString(StylePropertyEnum.StyleProperty_MaxHeight))
        Paddings = convertValue(vbaObject.GetPropertyAsString(StylePropertyEnum.StyleProperty_Padding))
        PaddingTop = convertValue(vbaObject.GetPropertyAsString(StylePropertyEnum.StyleProperty_PaddingTop))
        PaddingLeft = convertValue(vbaObject.GetPropertyAsString(StylePropertyEnum.StyleProperty_PaddingLeft))
        PaddingBottom = convertValue(vbaObject.GetPropertyAsString(StylePropertyEnum.StyleProperty_PaddingBottom))
        PaddingRight = convertValue(vbaObject.GetPropertyAsString(StylePropertyEnum.StyleProperty_PaddingRight))
        Margins = convertValue(vbaObject.GetPropertyAsString(StylePropertyEnum.StyleProperty_Margin))
        MarginTop = convertValue(vbaObject.GetPropertyAsString(StylePropertyEnum.StyleProperty_MarginTop))
        MarginLeft = convertValue(vbaObject.GetPropertyAsString(StylePropertyEnum.StyleProperty_MarginLeft))
        MarginBottom = convertValue(vbaObject.GetPropertyAsString(StylePropertyEnum.StyleProperty_MarginBottom))
        MarginRight = convertValue(vbaObject.GetPropertyAsString(StylePropertyEnum.StyleProperty_MarginRight))
        BackgroundColor = convertValue(vbaObject.GetPropertyAsString(StylePropertyEnum.StyleProperty_BackgroundColor))
        BorderThickness = convertValue(vbaObject.GetPropertyAsString(StylePropertyEnum.StyleProperty_BorderThickness))
        BorderColor = convertValue(vbaObject.GetPropertyAsString(StylePropertyEnum.StyleProperty_BorderColor))
        FontSize = convertValue(vbaObject.GetPropertyAsString(StylePropertyEnum.StyleProperty_FontSize))
        FontBold = convertValue(vbaObject.GetPropertyAsString(StylePropertyEnum.StyleProperty_FontBold))
        FontColor = convertValue(vbaObject.GetPropertyAsString(StylePropertyEnum.StyleProperty_FontColor))
        FontFamily = convertValue(vbaObject.GetPropertyAsString(StylePropertyEnum.StyleProperty_FontFamily))
        HorizontalAlignment = convertValue(vbaObject.GetPropertyAsString(StylePropertyEnum.StyleProperty_HorizontalAlignment))
        VerticalAlignment = convertValue(vbaObject.GetPropertyAsString(StylePropertyEnum.StyleProperty_VerticalAlignment))
        ImageFilePath = convertValue(vbaObject.GetPropertyAsString(StylePropertyEnum.StyleProperty_ImageFilePath))
        ImageWidth = convertValue(vbaObject.GetPropertyAsString(StylePropertyEnum.StyleProperty_ImageWidth))
        ImageHeight = convertValue(vbaObject.GetPropertyAsString(StylePropertyEnum.StyleProperty_ImageHeight))
        ImageName = convertValue(vbaObject.GetPropertyAsString(StylePropertyEnum.StyleProperty_ImageName))
        ImageSize = convertValue(vbaObject.GetPropertyAsString(StylePropertyEnum.StyleProperty_ImageSize))
    End Sub

    Private Function convertValue(ByVal strValue As String) As StylePropertyValue
        If strValue Is Nothing Then
            Return StylePropertyValue.Null()
        ElseIf strValue.Length = 0 Then
            Return StylePropertyValue.Null()
        ElseIf strValue = "auto" Then
            Return StylePropertyValue.Auto()
        ElseIf strValue = "inherit" Then
            Return StylePropertyValue.Inherited()
        Else
            Return New StylePropertyValue(strValue)
        End If
    End Function

#End Region


#Region "Public properties"


    Public Property Type() As StyleNodeTypeEnum
        Get
            Return pType
        End Get
        Set(value As StyleNodeTypeEnum)
            pType = value
        End Set
    End Property



    '[Position]
    Public Property Float() As StylePropertyValue
        Get
            Return pFloat
        End Get
        Set(value As StylePropertyValue)
            pFloat = value
        End Set
    End Property

    Public Property Position() As StylePropertyValue
        Get
            Return pPosition
        End Get
        Set(value As StylePropertyValue)
            pPosition = value
        End Set
    End Property

    Public Property Top() As StylePropertyValue
        Get
            Return pTop
        End Get
        Set(value As StylePropertyValue)
            Call applyForSizeOrPosition(value)
            pTop = value
        End Set
    End Property

    Public Property Left() As StylePropertyValue
        Get
            Return pLeft
        End Get
        Set(value As StylePropertyValue)
            Call applyForSizeOrPosition(value)
            pLeft = value
        End Set
    End Property

    Public Property Bottom() As StylePropertyValue
        Get
            Return pBottom
        End Get
        Set(value As StylePropertyValue)
            pBottom = value
        End Set
    End Property

    Public Property Right() As StylePropertyValue
        Get
            Return pRight
        End Get
        Set(value As StylePropertyValue)
            pRight = value
        End Set
    End Property



    '[Size]
    Public Property Width() As StylePropertyValue
        Get
            Return pWidth
        End Get
        Set(value As StylePropertyValue)
            Call applyForSizeOrPosition(value)
            pWidth = value
        End Set
    End Property

    Public Property MinWidth() As StylePropertyValue
        Get
            Return pMinWidth
        End Get
        Set(value As StylePropertyValue)
            Call applyForSizeOrPosition(value)
            pMinWidth = value
        End Set
    End Property

    Public Property MaxWidth() As StylePropertyValue
        Get
            Return pMaxWidth
        End Get
        Set(value As StylePropertyValue)
            Call applyForSizeOrPosition(value)
            pMaxWidth = value
        End Set
    End Property

    Public Property Height() As StylePropertyValue
        Get
            Return pHeight
        End Get
        Set(value As StylePropertyValue)
            Call applyForSizeOrPosition(value)
            pHeight = value
        End Set
    End Property

    Public Property MinHeight() As StylePropertyValue
        Get
            Return pMinHeight
        End Get
        Set(value As StylePropertyValue)
            Call applyForSizeOrPosition(value)
            pMinHeight = value
        End Set
    End Property

    Public Property MaxHeight() As StylePropertyValue
        Get
            Return pMaxHeight
        End Get
        Set(value As StylePropertyValue)
            Call applyForSizeOrPosition(value)
            pMaxHeight = value
        End Set
    End Property



    '[Margins and paddings]
    Public Property Margins() As StylePropertyValue
        Get
            Return pMargins
        End Get
        Set(value As StylePropertyValue)
            pMargins = value
        End Set
    End Property

    Public Property MarginTop() As StylePropertyValue
        Get
            Return pMarginTop
        End Get
        Set(value As StylePropertyValue)
            pMarginTop = value
        End Set
    End Property

    Public Property MarginLeft() As StylePropertyValue
        Get
            Return pMarginLeft
        End Get
        Set(value As StylePropertyValue)
            pMarginLeft = value
        End Set
    End Property

    Public Property MarginBottom() As StylePropertyValue
        Get
            Return pMarginBottom
        End Get
        Set(value As StylePropertyValue)
            pMarginBottom = value
        End Set
    End Property

    Public Property MarginRight() As StylePropertyValue
        Get
            Return pMarginRight
        End Get
        Set(value As StylePropertyValue)
            pMarginRight = value
        End Set
    End Property

    Public Property Paddings() As StylePropertyValue
        Get
            Return pPaddings
        End Get
        Set(value As StylePropertyValue)
            pPaddings = value
        End Set
    End Property

    Public Property PaddingTop() As StylePropertyValue
        Get
            Return pPaddingTop
        End Get
        Set(value As StylePropertyValue)
            pPaddingTop = value
        End Set
    End Property

    Public Property PaddingLeft() As StylePropertyValue
        Get
            Return pPaddingLeft
        End Get
        Set(value As StylePropertyValue)
            pPaddingLeft = value
        End Set
    End Property

    Public Property PaddingBottom() As StylePropertyValue
        Get
            Return pPaddingBottom
        End Get
        Set(value As StylePropertyValue)
            pPaddingBottom = value
        End Set
    End Property

    Public Property PaddingRight() As StylePropertyValue
        Get
            Return pPaddingRight
        End Get
        Set(value As StylePropertyValue)
            pPaddingRight = value
        End Set
    End Property



    '[Background & Border]
    Public Property BackgroundColor() As StylePropertyValue
        Get
            Return pBackgroundColor
        End Get
        Set(value As StylePropertyValue)
            Call applyForColor(value)
            pBackgroundColor = value
        End Set
    End Property

    Public Property BorderColor() As StylePropertyValue
        Get
            Return pBorderColor
        End Get
        Set(value As StylePropertyValue)
            Call applyForColor(value)
            pBorderColor = value
        End Set
    End Property

    Public Property BorderThickness() As StylePropertyValue
        Get
            Return pBorderThickness
        End Get
        Set(value As StylePropertyValue)
            pBorderThickness = value
        End Set
    End Property



    '[Alignments]
    Public Property HorizontalAlignment() As StylePropertyValue
        Get
            Return pHorizontalAlignment
        End Get
        Set(value As StylePropertyValue)
            pHorizontalAlignment = value
        End Set
    End Property

    Public Property VerticalAlignment() As StylePropertyValue
        Get
            Return pVerticalAlignment
        End Get
        Set(value As StylePropertyValue)
            pVerticalAlignment = value
        End Set
    End Property



    '[Font]
    Public Property FontSize() As StylePropertyValue
        Get
            Return pFontSize
        End Get
        Set(value As StylePropertyValue)
            pFontSize = value
        End Set
    End Property

    Public Property FontFamily() As StylePropertyValue
        Get
            Return pFontFamily
        End Get
        Set(value As StylePropertyValue)
            pFontFamily = value
        End Set
    End Property

    Public Property FontBold() As StylePropertyValue
        Get
            Return pFontBold
        End Get
        Set(value As StylePropertyValue)
            pFontBold = value
        End Set
    End Property

    Public Property FontColor() As StylePropertyValue
        Get
            Return pFontColor
        End Get
        Set(value As StylePropertyValue)
            Call applyForColor(value)
            pFontColor = value
        End Set
    End Property


    '[Images]
    Public Property ImageFilePath() As StylePropertyValue
        Get
            Return pImageFilePath
        End Get
        Set(value As StylePropertyValue)
            pImageFilePath = value
        End Set
    End Property

    Public Property ImageWidth() As StylePropertyValue
        Get
            Return pImageWidth
        End Get
        Set(value As StylePropertyValue)
            pImageWidth = value
        End Set
    End Property

    Public Property ImageHeight() As StylePropertyValue
        Get
            Return pImageHeight
        End Get
        Set(value As StylePropertyValue)
            pImageHeight = value
        End Set
    End Property

    Public Property ImageName() As StylePropertyValue
        Get
            Return pImageName
        End Get
        Set(value As StylePropertyValue)
            pImageName = value
        End Set
    End Property

    Public Property ImageSize() As StylePropertyValue
        Get
            Return pImageSize
        End Get
        Set(value As StylePropertyValue)
            pImageSize = value
        End Set
    End Property




#End Region


#Region "Converting property values"

    Private Sub applyForSizeOrPosition(value As StylePropertyValue)
        If Not value Is Nothing Then
            If IsNumeric(value.CssValue) Then
                value.Value = CSng(value.CssValue)
            End If
        End If
    End Sub

    Private Sub applyForColor(value As StylePropertyValue)
        If Not value Is Nothing Then
            value.Value = colorFromString(value.CssValue)
        End If
    End Sub

#End Region

#Region "Access"

    Public Function getProperty(ByVal propertyType As StylePropertyEnum) As StylePropertyValue
        Select Case propertyType
            ' [Position]
            Case StylePropertyEnum.StyleProperty_Float : Return pFloat
            Case StylePropertyEnum.StyleProperty_Position : Return pPosition
            Case StylePropertyEnum.StyleProperty_Top : Return pTop
            Case StylePropertyEnum.StyleProperty_Left : Return pLeft
            Case StylePropertyEnum.StyleProperty_Bottom : Return pBottom
            Case StylePropertyEnum.StyleProperty_Right : Return pRight
                ' [Size]
            Case StylePropertyEnum.StyleProperty_Width : Return pWidth
            Case StylePropertyEnum.StyleProperty_MinWidth : Return pMinWidth
            Case StylePropertyEnum.StyleProperty_MaxWidth : Return pMaxWidth
            Case StylePropertyEnum.StyleProperty_Height : Return pHeight
            Case StylePropertyEnum.StyleProperty_MinHeight : Return pMinHeight
            Case StylePropertyEnum.StyleProperty_MaxHeight : Return pMaxHeight
                ' [Margins & paddings]
            Case StylePropertyEnum.StyleProperty_Padding : Return pPaddings
            Case StylePropertyEnum.StyleProperty_PaddingTop : Return pPaddingTop
            Case StylePropertyEnum.StyleProperty_PaddingLeft : Return pPaddingLeft
            Case StylePropertyEnum.StyleProperty_PaddingBottom : Return pPaddingBottom
            Case StylePropertyEnum.StyleProperty_PaddingRight : Return pPaddingRight
            Case StylePropertyEnum.StyleProperty_Margin : Return pMargins
            Case StylePropertyEnum.StyleProperty_MarginTop : Return pMarginTop
            Case StylePropertyEnum.StyleProperty_MarginLeft : Return pMarginLeft
            Case StylePropertyEnum.StyleProperty_MarginBottom : Return pMarginBottom
            Case StylePropertyEnum.StyleProperty_MarginRight : Return pMarginRight
                '[Background & borders]
            Case StylePropertyEnum.StyleProperty_BackgroundColor : Return pBackgroundColor
            Case StylePropertyEnum.StyleProperty_BorderThickness : Return pBorderThickness
            Case StylePropertyEnum.StyleProperty_BorderColor : Return pBorderColor
                '[Font]
            Case StylePropertyEnum.StyleProperty_FontSize : Return pFontSize
            Case StylePropertyEnum.StyleProperty_FontFamily : Return pFontFamily
            Case StylePropertyEnum.StyleProperty_FontColor : Return pFontColor
            Case StylePropertyEnum.StyleProperty_FontBold : Return pFontBold
                '[Alignment]
            Case StylePropertyEnum.StyleProperty_HorizontalAlignment : Return pHorizontalAlignment
            Case StylePropertyEnum.StyleProperty_VerticalAlignment : Return pVerticalAlignment
                '[Images]
            Case StylePropertyEnum.StyleProperty_ImageFilePath : Return pImageFilePath
            Case StylePropertyEnum.StyleProperty_ImageName : Return pImageName
            Case StylePropertyEnum.StyleProperty_ImageWidth : Return pImageWidth
            Case StylePropertyEnum.StyleProperty_ImageHeight : Return pImageHeight
            Case StylePropertyEnum.StyleProperty_ImageSize : Return pImageSize
            Case Else : Return Nothing
        End Select

    End Function

#End Region


End Class
