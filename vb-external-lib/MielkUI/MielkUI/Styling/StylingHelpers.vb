Imports System.Drawing

Module StylingHelpers

    Public Function CountStyleNodeTypesEnums() As Long
        Static value As Long
        If value = 0 Then
            value = [Enum].GetValues(GetType(StyleNodeTypeEnum)).Cast(Of Integer).Max()
        End If
        Return value
    End Function

    Public Function CountStylePropertiesEnums() As Long
        Static value As Long
        If value = 0 Then
            value = [Enum].GetValues(GetType(StylePropertyEnum)).Cast(Of Integer).Max()
        End If
        Return value
    End Function



    Public Function getAutoValue(prop As StylePropertyEnum) As Object
        Static dict As Dictionary(Of StylePropertyEnum, Object)
        '---------------------------------------------------------------------------------------------------------

        If dict Is Nothing Then
            dict = New Dictionary(Of StylePropertyEnum, Object)
            With dict
                Call .Add(StylePropertyEnum.StyleProperty_BackgroundColor, Color.Transparent)
            End With
        End If

        With dict
            If .ContainsKey(prop) Then
                Return .Item(prop)
            Else
                Return Nothing
            End If
        End With

        'StyleProperty_Float = 1
        'StyleProperty_Position = 2
        'StyleProperty_Top = 3
        'StyleProperty_Left = 4
        'StyleProperty_Bottom = 5
        'StyleProperty_Right = 6
        ''[Size]
        'StyleProperty_Width = 7
        'StyleProperty_MinWidth = 8
        'StyleProperty_MaxWidth = 9
        'StyleProperty_Height = 10
        'StyleProperty_MinHeight = 11
        'StyleProperty_MaxHeight = 12
        ''[Margins & paddings]
        'StyleProperty_Padding = 13
        'StyleProperty_PaddingTop = 14
        'StyleProperty_PaddingLeft = 15
        'StyleProperty_PaddingBottom = 16
        'StyleProperty_PaddingRight = 17
        'StyleProperty_Margin = 18
        'StyleProperty_MarginTop = 19
        'StyleProperty_MarginLeft = 20
        'StyleProperty_MarginBottom = 21
        'StyleProperty_MarginRight = 22
        ''[Background & borders]
        'StyleProperty_BackgroundColor = 23
        'StyleProperty_BorderThickness = 24
        'StyleProperty_BorderColor = 25
        ''[Font]
        'StyleProperty_FontSize = 26
        'StyleProperty_FontFamily = 27
        'StyleProperty_FontColor = 28
        'StyleProperty_FontBold = 29
        ''[Alignment]
        'StyleProperty_HorizontalAlignment = 30
        'StyleProperty_VerticalAlignment = 31
        ''[Images]
        'StyleProperty_ImageFilePath = 32
        'StyleProperty_ImageName = 33
        'StyleProperty_ImageWidth = 34
        'StyleProperty_ImageHeight = 35
        'StyleProperty_ImageSize = 36
    End Function

End Module
