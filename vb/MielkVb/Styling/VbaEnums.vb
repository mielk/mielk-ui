Public Enum CssStyleTypeEnum
    CssStyleTypeEnum_Element = 1
    CssStyleTypeEnum_ElementClass = 2
    CssStyleTypeEnum_Class = 3
    CssStyleTypeEnum_Id = 4
End Enum

Public Enum CssStyleNodeEnum
    CssStyleNode_Normal = 0
    CssStyleNode_Hover = 1
    CssStyleNode_Clicked = 2
    CssStyleNode_Checked = 3
    CssStyleNode_Disabled = 4
End Enum

Public Enum CssEdgeEnum
    CssEdge_Top = 1
    CssEdge_Right = 2
    CssEdge_Bottom = 3
    CssEdge_Left = 4
End Enum

Public Enum CssPositionEnum
    CssPosition_Relative = 1
    CssPosition_Absolute = 2
End Enum

Public Enum CssFloatEnum
    CssFloat_Left = 1
    CssFloat_Right = 2
End Enum

Public Enum VbaStylePropertyEnum
    StyleProperty_Unknown = -1
    '[Position]
    StyleProperty_Float = 1
    StyleProperty_Position = 2
    StyleProperty_BorderBox = 3
    StyleProperty_Top = 4
    StyleProperty_Left = 5
    StyleProperty_Bottom = 6
    StyleProperty_Right = 7
    '[Size]
    StyleProperty_Width = 8
    StyleProperty_MinWidth = 9
    StyleProperty_MaxWidth = 10
    StyleProperty_Height = 11
    StyleProperty_MinHeight = 12
    StyleProperty_MaxHeight = 13
    '[Margins & paddings]
    StyleProperty_Padding = 14
    StyleProperty_PaddingTop = 15
    StyleProperty_PaddingLeft = 16
    StyleProperty_PaddingBottom = 17
    StyleProperty_PaddingRight = 18
    StyleProperty_Margin = 19
    StyleProperty_MarginTop = 20
    StyleProperty_MarginLeft = 21
    StyleProperty_MarginBottom = 22
    StyleProperty_MarginRight = 23
    '[Background & borders]
    StyleProperty_BackgroundColor = 24
    StyleProperty_BorderThickness = 25
    StyleProperty_BorderColor = 26
    '[Font]
    StyleProperty_FontSize = 27
    StyleProperty_FontFamily = 28
    StyleProperty_FontColor = 29
    StyleProperty_FontBold = 30
    '[Alignment]
    StyleProperty_HorizontalAlignment = 31
    StyleProperty_VerticalAlignment = 32
    '[Images]
    StyleProperty_ImageFilePath = 33
    StyleProperty_ImageName = 34
    StyleProperty_ImageWidth = 35
    StyleProperty_ImageHeight = 36
    StyleProperty_ImageSize = 37
End Enum