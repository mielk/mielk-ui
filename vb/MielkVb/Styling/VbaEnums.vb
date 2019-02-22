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
    StyleProperty_Top = 3
    StyleProperty_Left = 4
    StyleProperty_Bottom = 5
    StyleProperty_Right = 6
    '[Size]
    StyleProperty_Width = 7
    StyleProperty_MinWidth = 8
    StyleProperty_MaxWidth = 9
    StyleProperty_Height = 10
    StyleProperty_MinHeight = 11
    StyleProperty_MaxHeight = 12
    '[Margins & paddings]
    StyleProperty_Padding = 13
    StyleProperty_PaddingTop = 14
    StyleProperty_PaddingLeft = 15
    StyleProperty_PaddingBottom = 16
    StyleProperty_PaddingRight = 17
    StyleProperty_Margin = 18
    StyleProperty_MarginTop = 19
    StyleProperty_MarginLeft = 20
    StyleProperty_MarginBottom = 21
    StyleProperty_MarginRight = 22
    '[Background & borders]
    StyleProperty_BackgroundColor = 23
    StyleProperty_BorderThickness = 24
    StyleProperty_BorderColor = 25
    '[Font]
    StyleProperty_FontSize = 26
    StyleProperty_FontFamily = 27
    StyleProperty_FontColor = 28
    StyleProperty_FontBold = 29
    '[Alignment]
    StyleProperty_HorizontalAlignment = 30
    StyleProperty_VerticalAlignment = 31
    '[Images]
    StyleProperty_ImageFilePath = 32
    StyleProperty_ImageName = 33
    StyleProperty_ImageWidth = 34
    StyleProperty_ImageHeight = 35
    StyleProperty_ImageSize = 36
End Enum