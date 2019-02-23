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

    Private Sub rearrangeControls()

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

    'Public Sub AddControl(ctrl As IControl)
    '    Call ctrl.SetParent(Me)
    '    Call pControls.Add(ctrl)
    '    Call pControlsDict.Add(ctrl.GetId, ctrl)
    '    Call Me.Controls.Add(ctrl)
    'End Sub

    'Public Sub RemoveControl(ctrl As Control)

    'End Sub

#End Region


End Class