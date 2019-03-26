Imports System.Windows.Forms
Imports System.Runtime.InteropServices
Imports System.Drawing

Public Class Body
    Inherits System.Windows.Forms.Panel
    Implements IContainer


#Region "Private variables"

    '[Main properties]
    Private pWindow As Window
    Private pHandle As IntPtr
    Private pId As String

    '[Size and position]
    Private pCurrentProperties(0 To CountStyleNodeTypesEnums()) As Object
    Private pStylesMatrix As StylesMatrix

    '[Controls]
    Private pControls As Collection
    Private pControlsDict As Dictionary(Of String, IControl)

#End Region


#Region "Constructor"

    Public Sub New(window As Window)
        pWindow = Window
        pHandle = Me.Handle
        pControls = New Collection
        pControlsDict = New Dictionary(Of String, IControl)
    End Sub

#End Region


#Region "Setters"

    Public Sub SetWindow(value As Window)
        pWindow = value
    End Sub

#End Region


#Region "Getters"

    Public Function GetParent() As IContainer Implements IContainer.GetParent
        Return pWindow
    End Function

    Public Function GetWindow() As Window Implements IContainer.GetWindow
        Return pWindow
    End Function

    Public Function GetWidth() As Single Implements IContainer.GetWidth
        Return Me.Width
    End Function

    Public Function GetInnerWidth() As Single Implements IContainer.GetInnerWidth
        Return Me.Width - getPaddingsWidth()
    End Function

    Public Function GetHeight() As Single Implements IContainer.GetHeight
        Return Me.Height
    End Function

    Public Function GetInnerHeight() As Single Implements IContainer.GetInnerHeight
        Return Me.Height - getPaddingsHeight()
    End Function

    Public Function GetPaddingLeft() As Single Implements IContainer.GetPaddingLeft
        Return pCurrentProperties(StylePropertyEnum.StyleProperty_PaddingLeft)
    End Function

    Public Function GetPaddingRight() As Single Implements IContainer.GetPaddingRight
        Return pCurrentProperties(StylePropertyEnum.StyleProperty_PaddingRight)
    End Function

    Public Function GetPaddingTop() As Single Implements IContainer.GetPaddingTop
        Return pCurrentProperties(StylePropertyEnum.StyleProperty_PaddingTop)
    End Function

    Public Function GetPaddingBottom() As Single Implements IContainer.GetPaddingBottom
        Return pCurrentProperties(StylePropertyEnum.StyleProperty_PaddingBottom)
    End Function

    Private Function getPaddingsWidth() As Single
        Return pCurrentProperties(StylePropertyEnum.StyleProperty_PaddingLeft) + pCurrentProperties(StylePropertyEnum.StyleProperty_PaddingRight)
    End Function

    Private Function getPaddingsHeight() As Single
        Return pCurrentProperties(StylePropertyEnum.StyleProperty_PaddingTop) + pCurrentProperties(StylePropertyEnum.StyleProperty_PaddingBottom)
    End Function

#End Region


#Region "Rendering"

    Public Sub updateSizeAndPosition()
        Call updateSize()
        Call updatePosition()
        Call updateBackground()
        Call rearrangeControls()
    End Sub

    Private Sub updateSize()
        With pWindow
            Me.Width = .GetBodyWidth
            Me.Height = .GetBodyHeight
        End With
    End Sub

    Private Sub updatePosition()
        Me.Top = pWindow.GetPaddingTop
        Me.Left = pWindow.GetPaddingLeft
    End Sub

    Private Sub updateBackground()
        'Me.BackColor = colorFromString("#CCCCCC")
    End Sub

    Public Sub RearrangeControls() Implements IContainer.RearrangeControls
        Dim ctrl As IControl
        '------------------------------------------------------------
        For Each ctrl In pControls
            Call ctrl.UpdateSizeAndPosition()
        Next
    End Sub

#End Region


#Region "Styling"

    Public Function GetStylesManager() As StylesManager Implements IContainer.GetStylesManager
        Static instance As StylesManager
        '------------------------------------------------------------
        If instance Is Nothing Then instance = pWindow.GetStylesManager
        Return instance
    End Function

#End Region


#Region "Managing children"

    Public Sub AddControl(ctrl As IControl)
        Call ctrl.SetParent(pWindow)
        Call pControls.Add(ctrl)
        Call pControlsDict.Add(ctrl.GetId, ctrl)
        Call Me.Controls.Add(ctrl)
    End Sub

    'Public Sub RemoveControl(ctrl As Control)

    'End Sub

#End Region


End Class