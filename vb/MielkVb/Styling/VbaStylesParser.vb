Imports System.Text.RegularExpressions

Public Class VbaStylesParser

    Public Function parse(ByVal text As String) As Collection
        Dim items As Object
        Dim item As String
        Dim i As Long
        Dim header As String
        Dim body As String
        '------------------------------------------------------------------------------------------------------
        Dim dict As New Dictionary(Of String, VbaStyleSet)
        Dim headerDef As VbaStylesHeader
        Dim key As String
        Dim objStyleSet As VbaStyleSet
        '------------------------------------------------------------------------------------------------------
        Dim objStyle As VbaStyle
        Dim nodeType As CssStyleNodeEnum
        Dim propertiesDict As Dictionary(Of VbaStylePropertyEnum, String)
        '------------------------------------------------------------------------------------------------------
        Dim openBracketPosition As Integer
        '------------------------------------------------------------------------------------------------------

        items = text.Split("}")

        For i = LBound(items) To UBound(items)
            item = items(i)
            If item.Trim().Length > 0 Then
                openBracketPosition = item.IndexOf("{")
                header = item.Substring(0, openBracketPosition).Trim
                body = item.Substring(openBracketPosition + 1).Trim

                headerDef = parseCssEntryHeader(header)

                objStyleSet = New VbaStyleSet
                Call objStyleSet.setElement(headerDef.HeaderElement)
                Call objStyleSet.setClass(headerDef.HeaderClass)
                Call objStyleSet.setId(headerDef.HeaderId)

                If dict.ContainsKey(objStyleSet.getKey) Then
                    objStyleSet = dict.Item(objStyleSet.getKey)
                Else
                    Call dict.Add(objStyleSet.getKey, objStyleSet)
                End If

                nodeType = convertStringToNodeType(headerDef.HeaderEvent)
                If Not objStyleSet.hasNode(nodeType) Then
                    objStyle = New VbaStyle(objStyleSet, nodeType)
                    propertiesDict = parseBody(body)
                    Call objStyle.addProperties(propertiesDict)
                    Call objStyleSet.addNode(nodeType, objStyle)
                End If

            End If
        Next i

        Dim col As New Collection
        Dim varKey As String

        For Each varKey In dict.Keys
            Call col.Add(dict.Item(varKey))
        Next varKey

        Return col

    End Function

    Private Function parseCssEntryHeader(ByVal text As String) As VbaStylesHeader
        Const REGEX_PATTERN As String = "(.*?)[.:#{]"
        '------------------------------------------------------------------------------------------------------
        Dim regex As Regex = New Regex(REGEX_PATTERN)
        Dim matches As Object
        Dim typeSymbol As String
        '------------------------------------------------------------------------------------------------------
        Dim header As New VbaStylesHeader
        Dim openBracketPosition As Integer
        Dim colonPosition As Integer
        '------------------------------------------------------------------------------------------------------

        openBracketPosition = text.IndexOf("{")
        If openBracketPosition >= 0 Then
            text = text.Substring(openBracketPosition + 1)
        End If

        'Check for event.
        colonPosition = text.IndexOf(":")
        If colonPosition >= 0 Then
            header.HeaderEvent = text.Substring(colonPosition + 1)
            text = text.Substring(0, colonPosition)
        End If

        typeSymbol = text.Substring(0, 1)
        If typeSymbol = "#" Then
            header.HeaderId = text.Substring(1)
        ElseIf typeSymbol = "." Then
            header.HeaderClass = text.Substring(1)
        Else
            matches = regex.Matches(text)
            If matches.Count Then
                header.HeaderElement = CStr(matches(0).Value).Replace(".", vbNullString)
                header.HeaderClass = text.Replace(header.HeaderElement & ".", vbNullString)
            Else
                header.HeaderElement = text.Trim
            End If
        End If

        Return header

    End Function

    Private Function parseBody(body As String) As Dictionary(Of VbaStylePropertyEnum, String)
        Dim lines As Object
        Dim strLine As String
        Dim propertyName As String
        Dim propertyEnum As VbaStylePropertyEnum
        Dim propertyValue As String
        Dim colonPosition As Integer
        Dim dict As New Dictionary(Of VbaStylePropertyEnum, String)
        '------------------------------------------------------------------------------------------------------

        lines = body.Split(";")
        For Each strLine In lines
            If strLine.Length > 0 Then
                colonPosition = strLine.IndexOf(":")
                propertyName = strLine.Substring(0, colonPosition)
                propertyEnum = convertStringToStyleProperty(propertyName)
                If Not propertyEnum = VbaStylePropertyEnum.StyleProperty_Unknown Then
                    propertyValue = strLine.Substring(colonPosition + 1).Trim
                    With dict
                        If .ContainsKey(propertyEnum) Then .Remove(propertyEnum)
                        Call .Add(propertyEnum, propertyValue)
                    End With
                End If
            End If
        Next strLine

        Return dict

    End Function

End Class