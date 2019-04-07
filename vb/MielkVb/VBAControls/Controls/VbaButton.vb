Imports System.Text.RegularExpressions

Public Class VbaButton
    Implements IVbaControl

    Private pVbNetObject As Button
    Private pContainer As IVbaContainer
    Private pId As String
    Private pCssClass As String
    Private pCaption As String


    Public Sub New(parent As IVbaContainer)
        pContainer = parent
    End Sub


    Public Sub LoadFromXml(xmlNode As Xml.XmlNode) Implements IVbaControl.LoadFromXml
        Call loadAttributes(xmlNode)
        pVbNetObject = New Button(pContainer.getVbNetObject, pId)
        With pVbNetObject
            Call .SetCaption(pCaption)
            Call .SetListener(Me)
            Call .SetStyleClasses(pCssClass, True)
        End With
    End Sub

    Private Sub loadAttributes(xmlNode As Xml.XmlNode)
        On Error Resume Next
        pId = xmlNode.Attributes.ItemOf("id").Value
        pCssClass = xmlNode.Attributes.ItemOf("class").Value
        pCaption = getCaption(xmlNode.InnerText)
        On Error GoTo 0
    End Sub

    Private Function getCaption(innerText As String) As String
        Const REGEX_PATTERN_CHECK As String = "^\[.*\]$"
        Const REGEX_PATTERN_MATCH As String = "(?<=\[).+?(?=\])"
        '------------------------------------------------------------------------------------------------------
        Dim regexCheck As Regex = New Regex(REGEX_PATTERN_CHECK)
        Dim regexMatch As Regex = New Regex(REGEX_PATTERN_MATCH)
        Dim match As System.Text.RegularExpressions.Match
        Dim tag As String
        '------------------------------------------------------------------------------------------------------

        If regexCheck.Match(innerText).Success Then
            match = regexMatch.Match(innerText)
            tag = match.Value
            Return GetTranslation(tag)
        Else
            Return innerText
        End If

    End Function

    Public Sub SetParent(value As IVbaContainer) Implements IVbaControl.SetParent
        pContainer = value
    End Sub

    Public Function getVbNetObject() As Object Implements IVbaControl.getVbNetObject
        Return pVbNetObject
    End Function

    Public Sub TriggerEvent(eventName As String)
        If eventName = "click" Then
            'Call MsgBox(pCaption & " clicked")
        End If
    End Sub

End Class
