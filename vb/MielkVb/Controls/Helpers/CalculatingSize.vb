Module CalculatingSize

    Public Function CalculateControlHeight(parent As IContainer, properties As Object) As Single
        Dim sngHeight As Single
        Dim height As Object
        Dim minHeight As Object
        Dim maxHeight As Object
        Dim parentHeight As Single
        Dim bottom As Object
        Dim top As Object
        '------------------------------------------------------------------------------------------------------------------------------------------------------------

        height = properties(StylePropertyEnum.StyleProperty_Height)
        minHeight = properties(StylePropertyEnum.StyleProperty_MinHeight)
        maxHeight = properties(StylePropertyEnum.StyleProperty_MaxHeight)

        If Not height Is Nothing Then
            sngHeight = CSng(height)
        Else
            If properties(StylePropertyEnum.StyleProperty_Position) = CssPositionEnum.CssPosition_Absolute Then
                parentHeight = parent.GetHeight
                bottom = properties(StylePropertyEnum.StyleProperty_Bottom)
                top = properties(StylePropertyEnum.StyleProperty_Top)
                If Not bottom Is Nothing And Not top Is Nothing Then
                    sngHeight = parentHeight - top - bottom
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

    Public Function CalculateControlWidth(parent As IContainer, properties As Object) As Single
        Dim width As Object
        Dim sngWidth As Single
        Dim minWidth As Object
        Dim maxWidth As Object
        Dim parentWidth As Single
        Dim left As Object
        Dim right As Object
        '------------------------------------------------------------------------------------------------------------------------------------------------------------

        width = properties(StylePropertyEnum.StyleProperty_Width)
        minWidth = properties(StylePropertyEnum.StyleProperty_MinWidth)
        maxWidth = properties(StylePropertyEnum.StyleProperty_MaxWidth)

        If Not width Is Nothing Then
            sngWidth = CSng(width)
        Else
            If properties(StylePropertyEnum.StyleProperty_Position) = CssPositionEnum.CssPosition_Absolute Then
                parentWidth = parent.GetWidth
                left = properties(StylePropertyEnum.StyleProperty_Left)
                right = properties(StylePropertyEnum.StyleProperty_Right)
                If Not left Is Nothing And Not right Is Nothing Then
                    Return parentWidth - left - right
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

    Public Function CalculateControlLeft(item As IControl, parent As IContainer, properties As Object) As Single
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
                    parentWidth = parent.GetWidth
                    Return parentWidth - right - item.GetWidth
                Else
                    Return 0
                End If
            End If
        Else
            Return 1
        End If

    End Function

    Public Function CalculateControlTop(item As IControl, parent As IContainer, properties As Object) As Single
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
                    parentHeight = parent.GetHeight
                    Return parentHeight - bottom - item.GetHeight
                Else
                    Return 0
                End If
            End If
        Else
            Return 1
        End If

    End Function

End Module
