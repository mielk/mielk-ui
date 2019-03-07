Public Class StylesManager
    Private pIdStyleSets As Dictionary(Of String, StyleSet)
    Private pClassStyleSets As Dictionary(Of String, StyleSet)
    Private pElementStyleSets As Dictionary(Of String, StyleSet)
    Private pElementClassStyleSets As Dictionary(Of String, Dictionary(Of String, StyleSet))
    '-------------------------------------------------------------------
    Private pIsLoaded As Boolean
    '-------------------------------------------------------------------


    Public Sub New()
        pIdStyleSets = New Dictionary(Of String, StyleSet)
        pClassStyleSets = New Dictionary(Of String, StyleSet)
        pElementStyleSets = New Dictionary(Of String, StyleSet)
        pElementClassStyleSets = New Dictionary(Of String, Dictionary(Of String, StyleSet))
    End Sub


#Region "Loading styles"

    Public Sub LoadStyleSet(vbaStyleSet As Object)
        Dim ss As StyleSet
        '-------------------------------------------------------------------------------------------------------------------

        ss = StyleSet.CreateInstanceFromVbaObject(vbaStyleSet)
        With ss
            Select Case .getStyleType
                Case StyleTypeEnum.StyleType_Id : Call pIdStyleSets.Add(.getId, ss)
                Case StyleTypeEnum.StyleType_Class : Call pClassStyleSets.Add(.getClassName, ss)
                Case StyleTypeEnum.StyleType_Element : Call pElementStyleSets.Add(.getElementTag, ss)
                Case StyleTypeEnum.StyleType_ElementClass : Call addElementClassStyleSet(ss)
            End Select
        End With
        pIsLoaded = True

    End Sub

    Private Sub addElementClassStyleSet(styleSet As StyleSet)
        Dim dict As Dictionary(Of String, StyleSet)
        Dim elementTag As String = styleSet.getElementTag
        Dim className As String = styleSet.getClassName
        '-------------------------------------------------------------------------------------------------------------------

        If pElementClassStyleSets.ContainsKey(elementTag) Then
            dict = pElementClassStyleSets.Item(elementTag)
        Else
            dict = New Dictionary(Of String, StyleSet)
            Call pElementClassStyleSets.Add(elementTag, dict)
        End If

        Call dict.Add(className, styleSet)

    End Sub

    Public Function IsLoaded() As Boolean
        Return pIsLoaded
    End Function

#End Region



#Region "Access methods"

    Public Function GetStyleByElementAndClass(elementTag As String, className As String) As StyleSet
        Dim dict As Dictionary(Of String, StyleSet)
        '-------------------------------------------------------------------
        If pElementClassStyleSets.ContainsKey(elementTag) Then
            dict = pElementClassStyleSets.Item(elementTag)
            If dict.ContainsKey(className) Then
                Return dict.Item(className)
            Else
                Return Nothing
            End If
        Else
            Return Nothing
        End If
    End Function

    Public Function GetStyleByElementTag(elementTag As String) As StyleSet
        If pElementStyleSets.ContainsKey(elementTag) Then
            Return pElementStyleSets.Item(elementTag)
        Else
            Return Nothing
        End If
    End Function

    Public Function GetStyleByClassName(className As String) As StyleSet
        If pClassStyleSets.ContainsKey(className) Then
            Return pClassStyleSets.Item(className)
        Else
            Return Nothing
        End If
    End Function

    Public Function GetStyleById(id As String) As StyleSet
        If pIdStyleSets.ContainsKey(id) Then
            Return pIdStyleSets.Item(id)
        Else
            Return Nothing
        End If
    End Function

    Public Function countStyles() As Long
        Return pElementStyleSets.Count
    End Function

    Public Function getAllElementNames() As String
        Dim v As Object
        Dim x As String = vbNullString
        For Each v In pElementStyleSets.Keys
            x = x & v
        Next
        Return x
    End Function

#End Region

End Class
