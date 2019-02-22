Imports System.Drawing

Module Colors

    Public Function colorFromNumeric(color As Long, Optional opacity As Single = 1) As Color
        Dim blue As Long = color \ 65536
        Dim green As Long = (color - blue * 65536) \ 256
        Dim red As Long = color - blue * 65536 - green * 256
        '------------------------------------------------------------------------------------------------------------------------------------------------------------
        Return System.Drawing.Color.FromArgb(Math.Max(Math.Min(opacity * 255, 255), 0), red, green, blue)
    End Function


    Public Function colorFromString(str As String) As Color
        Dim colorParts() As String
        Dim blue As Long
        Dim green As Long
        Dim red As Long
        Dim opacity As Single
        '------------------------------------------------------------------------------------------------------------------------------------------------------------

        If str Is Nothing Then

        ElseIf str.StartsWith("rgb(") Then
            colorParts = Split(Mid(str, 5, str.Length - 5), ",")
            If UBound(colorParts) - LBound(colorParts) = 2 Then
                red = colorParts(LBound(colorParts))
                green = colorParts(LBound(colorParts) + 1)
                blue = colorParts(LBound(colorParts) + 2)
            End If
        ElseIf str.StartsWith("rgba(") Then
            colorParts = Split(Mid(str, 6, str.Length - 6), ",")
            If UBound(colorParts) - LBound(colorParts) = 3 Then
                red = colorParts(LBound(colorParts))
                green = colorParts(LBound(colorParts) + 1)
                blue = colorParts(LBound(colorParts) + 2)
                opacity = colorParts(LBound(colorParts) + 3)
            End If
            Return System.Drawing.Color.FromArgb(opacity, red, green, blue)
        ElseIf str.StartsWith("#") Then
            Return ColorTranslator.FromHtml(str)
        End If

        Return System.Drawing.Color.FromArgb(red, green, blue)

    End Function

End Module
