Module CalculatingSize

    Public Function CalculateControlHeight(item As IControl, properties As Object) As Single
        Dim sngHeight As Single
        Dim height As Object
        Dim minHeight As Object
        Dim maxHeight As Object
        Dim parentHeight As Single
        Dim bottomPosition As Object
        Dim topPosition As Object
        '------------------------------------------------------------------------------------------------------------------------------------------------------------

        height = properties(StylePropertyEnum.StyleProperty_Height)
        minHeight = properties(StylePropertyEnum.StyleProperty_MinHeight)
        maxHeight = properties(StylePropertyEnum.StyleProperty_MaxHeight)


        If Not height Is Nothing Then
            If IsNumeric(height) Then
                sngHeight = CSng(height)
            ElseIf height = AUTO Then
                Return 0
            ElseIf Right(height, 1) = "%" Then
                Stop
            End If
        Else
            If properties(StylePropertyEnum.StyleProperty_Position) = CssPositionEnum.CssPosition_Absolute Then
                parentHeight = item.GetParent.GetHeight
                bottomPosition = properties(StylePropertyEnum.StyleProperty_Bottom)
                topPosition = properties(StylePropertyEnum.StyleProperty_Top)
                If Not bottomPosition Is Nothing And Not topPosition Is Nothing Then
                    sngHeight = parentHeight - topPosition - bottomPosition
                End If
            End If
        End If

        If Not minHeight Is Nothing Then
            If sngHeight < minHeight Then sngHeight = minHeight
        End If

        If Not maxHeight Is Nothing Then
            If sngHeight > maxHeight Then sngHeight = maxHeight
        End If

        Return sngHeight

    End Function

    Public Function CalculateControlWidth(item As IControl, properties As Object) As Single
        Dim width As Object
        Dim sngWidth As Single
        Dim minWidth As Object
        Dim maxWidth As Object
        Dim parentWidth As Single
        Dim leftPosition As Object
        Dim rightPosition As Object
        '------------------------------------------------------------------------------------------------------------------------------------------------------------

        width = properties(StylePropertyEnum.StyleProperty_Width)
        minWidth = properties(StylePropertyEnum.StyleProperty_MinWidth)
        maxWidth = properties(StylePropertyEnum.StyleProperty_MaxWidth)

        If Not width Is Nothing Then
            If IsNumeric(width) Then
                sngWidth = CSng(width)
            ElseIf width = AUTO Then
                Return 0
            ElseIf Right(width, 1) = "%" Then
                Stop
            End If
        Else
            If properties(StylePropertyEnum.StyleProperty_Position) = CssPositionEnum.CssPosition_Absolute Then
                parentWidth = item.GetParent.GetWidth
                leftPosition = properties(StylePropertyEnum.StyleProperty_Left)
                rightPosition = properties(StylePropertyEnum.StyleProperty_Right)
                If Not leftPosition Is Nothing And Not rightPosition Is Nothing Then
                    Return parentWidth - leftPosition - rightPosition
                End If
            End If
        End If

            If Not minWidth Is Nothing Then
                If sngWidth < minWidth Then sngWidth = minWidth
            End If

            If Not maxWidth Is Nothing Then
                If sngWidth > maxWidth Then sngWidth = maxWidth
            End If

            Return sngWidth

    End Function

    Public Function CalculateControlLeft(item As IControl, properties As Object) As Single
        Dim left As Object
        Dim right As Object
        Dim parentLeftPadding As Single
        Dim parentWidth As Single
        ''------------------------------------------------------------------------------------------------------------------------------------------------------------

        If properties(StylePropertyEnum.StyleProperty_Position) = CssPositionEnum.CssPosition_Absolute Then
            left = properties(StylePropertyEnum.StyleProperty_Left)
            If Not left Is Nothing Then
                Return CSng(left)
            Else
                right = properties(StylePropertyEnum.StyleProperty_Right)
                If Not right Is Nothing Then
                    parentWidth = item.GetParent.GetWidth
                    Return parentWidth - right - item.GetWidth
                Else
                    Return 0
                End If
            End If
        Else
            Return 1
        End If

    End Function

    Public Function CalculateControlTop(item As IControl, properties As Object) As Single
        Dim top As Object
        Dim bottom As Object
        Dim parentTopPadding As Single
        Dim parentHeight As Single
        ''------------------------------------------------------------------------------------------------------------------------------------------------------------

        If properties(StylePropertyEnum.StyleProperty_Position) = CssPositionEnum.CssPosition_Absolute Then
            top = properties(StylePropertyEnum.StyleProperty_Top)
            If Not top Is Nothing Then
                Return CSng(top)
            Else
                bottom = properties(StylePropertyEnum.StyleProperty_Bottom)
                If Not bottom Is Nothing Then
                    parentHeight = item.GetParent.GetHeight
                    Return parentHeight - bottom - item.GetHeight
                Else
                    Return 0
                End If
            End If
        Else
            Return 1
        End If

    End Function

    Public Function isAutoValue(value As Object) As Boolean
        If Not value Is Nothing Then
            If Not IsNumeric(value) Then
                If value = AUTO Then
                    Return True
                End If
            End If
        End If
        Return False
    End Function



    Public Function createFont(properties As Object) As Font
        Const DEFAULT_FONT_SIZE As Single = 10
        '------------------------------------------------------------------------------------------------------------------------------------------------------------
        Dim bold As System.Drawing.FontStyle
        Dim family As String
        Dim size As Single
        '------------------------------------------------------------------------------------------------------------------------------------------------------------
        bold = IIf(properties(StylePropertyEnum.StyleProperty_FontBold), FontStyle.Bold, FontStyle.Regular)
        family = properties(StylePropertyEnum.StyleProperty_FontFamily)
        size = If(properties(StylePropertyEnum.StyleProperty_FontSize), DEFAULT_FONT_SIZE)
        Return New Font(family, size, bold)
    End Function

End Module