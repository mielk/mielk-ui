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

    '[Size and position]
    Private pCurrentProperties(0 To CountStyleNodeTypesEnums()) As Object
    Private pStylesMatrix As StylesMatrix
    'Private pTop As Single
    'Private pBottom As Single
    'Private pLeft As Single
    'Private pRight As Single
    'Private pWidth As Single
    'Private pHeight As Single

    '[Controls]
    Private pControls As Collection
    Private pControlsDict As Dictionary(Of String, IControl)

    '[Debugging]
    Private pLog As String

#End Region


#Region "Constructor"

    Public Sub New(window As Window)
        pWindow = Window
        pHandle = Me.Handle
        pControls = New Collection
        pControlsDict = New Dictionary(Of String, IControl)
        pLog = pLog & "Created | "
    End Sub

#End Region


#Region "Setters"

    ''[Position]
    'Public Sub SetTop(value As Single)
    '    pTop = value
    '    Me.Top = pTop
    '    Call updatePosition()
    'End Sub

    'Public Sub SetLeft(value As Single)
    '    pLeft = value
    '    Me.Left = pLeft
    '    Call updatePosition()
    'End Sub

    'Public Sub SetBottom(value As Single)
    '    pBottom = value
    '    pTop = Single.NaN
    '    Call updatePosition()
    'End Sub

    'Public Sub SetRight(value As Single)
    '    pRight = value
    '    pLeft = Single.NaN
    '    Call updatePosition()
    'End Sub


    ''[Size]
    'Public Sub SetWidth(value As VariantType?)
    '    pWidth = value
    '    Call updateSizeAndPosition()
    'End Sub

    'Public Sub SetAutoWidth(value As Boolean)
    '    pAutoWidth = value
    '    If pAutoWidth Then
    '        pWidth = Single.NaN
    '    Else
    '        pWidth = Me.Width
    '    End If
    '    Call updateSizeAndPosition()
    'End Sub

    'Public Sub SetHeight(value As Single)
    '    pHeight = value
    '    pAutoHeight = False
    '    Call updateSizeAndPosition()
    'End Sub

    'Public Sub SetAutoHeight(value As Boolean)
    '    pAutoHeight = value
    '    If pAutoHeight Then
    '        pHeight = Single.NaN
    '    Else
    '        pHeight = Me.Height
    '    End If
    '    Call updateSizeAndPosition()
    'End Sub



    ''[Margins]
    'Public Sub SetMargins(ByVal value As Single)
    '    pMarginTop = value
    '    pMarginRight = value
    '    pMarginBottom = value
    '    pMarginLeft = value
    '    Call updateSizeAndPosition()
    'End Sub

    'Public Sub SetMarginTop(value As Single)
    '    pMarginTop = value
    '    Call updateSizeAndPosition()
    'End Sub

    'Public Sub SetMarginRight(value As Single)
    '    pMarginRight = value
    '    Call updateSizeAndPosition()
    'End Sub

    'Public Sub SetMarginBottom(value As Single)
    '    pMarginBottom = value
    '    Call updateSizeAndPosition()
    'End Sub

    'Public Sub SetMarginLeft(value As Single)
    '    pMarginLeft = value
    '    Call updateSizeAndPosition()
    'End Sub


    ''[Paddings]
    'Public Sub SetPaddings(ByVal value As Single)
    '    pPaddingTop = value
    '    pPaddingRight = value
    '    pPaddingBottom = value
    '    pPaddingLeft = value
    '    Call updateSizeAndPosition()
    'End Sub

    'Public Sub SetPaddingTop(value As Single)
    '    pPaddingTop = value
    '    Call updateSizeAndPosition()
    'End Sub

    'Public Sub SetPaddingRight(value As Single)
    '    pPaddingRight = value
    '    Call updateSizeAndPosition()
    'End Sub

    'Public Sub SetPaddingBottom(value As Single)
    '    pPaddingBottom = value
    '    Call updateSizeAndPosition()
    'End Sub

    'Public Sub SetPaddingLeft(value As Single)
    '    pPaddingLeft = value
    '    Call updateSizeAndPosition()
    'End Sub



    ''[Background]
    'Public Sub SetBackgroundColor(color As Long, Optional opacity As Single = 1)
    '    pBackgroundColor = colorFromNumeric(Color, opacity)
    '    Me.BackColor = pBackgroundColor
    'End Sub

    'Public Sub SetBorder(color As Long, thickness As Long)
    '    pBorderColor = colorFromNumeric(color)
    '    pBorderThickness = thickness
    'End Sub

    'Public Sub SetBorderThickness(thickness As Long)
    '    pBorderThickness = thickness
    '    Call updateSizeAndPosition()
    '    Call updateBorders()
    'End Sub

    'Public Sub SetBorderColor(color As Long)
    '    pBorderColor = colorFromNumeric(color)
    '    Call updateBorders()
    'End Sub

#End Region



#Region "Rendering"

    Public Sub Redraw()
        Call updateSizeAndPosition()
    End Sub

    Private Sub updateSizeAndPosition()
        updateSize()
        updatePosition()
    End Sub

    Private Sub updateSize()
        'Me.Height = pHeight
        'Me.Width = pWidth
    End Sub

    Private Sub updatePosition()
        'Me.Top = pTop
        'Me.Left = pLeft
    End Sub

    'Private Sub redrawControls()

    'End Sub

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
        Call ctrl.SetParent(Me)
        Call pControls.Add(ctrl)
        Call pControlsDict.Add(ctrl.GetId, ctrl)
        Call Me.Controls.Add(ctrl)
    End Sub

    Public Sub RemoveControl(ctrl As Control)

    End Sub

#End Region


    Public Function getLog() As String
        Return pLog
    End Function

End Class
