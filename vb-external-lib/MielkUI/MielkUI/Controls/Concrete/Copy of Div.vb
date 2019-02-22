Imports System.Windows.Forms
Imports System.Runtime.InteropServices
Imports System.Drawing

Public Class Div
    Inherits System.Windows.Forms.Panel
    Implements IContainer, IControl



#Region "Private variables"

    Private Const TAG_NAME As String = "div"

    '[Main properties]
    Private pParent As IContainer
    Private pListener As Object
    Private pHandle As IntPtr
    Private pId As String

    '[Size and position]
    Private pBorderBox As Boolean
    Private pTop As Single
    Private pBottom As Single
    Private pLeft As Single
    Private pRight As Single
    Private pWidth As Single
    Private pAutoWidth As Single
    Private pHeight As Single
    Private pAutoHeight As Single

    '[Margins & paddings]
    Private pMarginTop As Single
    Private pMarginLeft As Single
    Private pMarginRight As Single
    Private pMarginBottom As Single
    Private pPaddingTop As Single
    Private pPaddingLeft As Single
    Private pPaddingRight As Single
    Private pPaddingBottom As Single

    '[Background & Border]
    Private pBackgroundColor As Color
    Private pBorderThickness As Double
    Private pBorderColor As Color

    '[Controls]
    Private pControls As Collection
    Private pControlsDict As Dictionary(Of String, IControl)

    '[Style]
    Private pStylesMatrix As StylesMatrix

    '[State]
    Private pIsRendered As Boolean

    Private pLog As String

#End Region


#Region "Constructor"

    Public Sub New(parent As IContainer, id As String)
        pId = id
        pParent = parent
        pHandle = Me.Handle
        pControls = New Collection
        pControlsDict = New Dictionary(Of String, IControl)
        pStylesMatrix = New StylesMatrix(Me)
        pLog = pLog & "Constructor|"
    End Sub

#End Region


#Region "Setters"

    '[Connection with VBA]
    Public Sub SetListener(value As Object)
        pListener = value
    End Sub

    Public Sub SetBorderBox(value As Boolean)
        pBorderBox = value
        Call updateSizeAndPosition()
    End Sub



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


#Region "Getters"

    Public Function GetId() As String Implements IControl.GetId
        Return pId
    End Function

    Public Function GetElementTag() As String Implements IControl.GetElementTag
        Return TAG_NAME
    End Function

    Public Function GetBorderBox() As Boolean
        Return pBorderBox
    End Function

    Public Function GetTop() As Single
        Return Me.Top
    End Function

    Public Function GetBottom() As Single
        Return Me.Bottom
    End Function

    Public Function GetLeft() As Single
        Return Me.Left
    End Function

    Public Function GetRight() As Single
        Return Me.Right
    End Function

    Public Function GetWidth() As Single
        Return Me.Width
    End Function

    Public Function GetHeight() As Single
        Return Me.Height
    End Function



    ''[Background & borders]
    'Public Function GetBackgroundColor() As Color
    '    Return pBackgroundColor
    'End Function

    'Public Function GetBorderThickness() As Long
    '    Return pBorderThickness
    'End Function

    'Public Function GetBorderColor() As Color
    '    Return pBorderColor
    'End Function



    ''[Margin and paddings]
    'Public Function GetMarginsHeight() As Single
    '    Return pMarginBottom + pMarginTop
    'End Function

    'Public Function GetMarginsWidth() As Single
    '    Return pMarginLeft + pMarginRight
    'End Function

    'Public Function GetMarginTop() As Single
    '    Return pMarginTop
    'End Function

    'Public Function GetMarginRighe() As Single
    '    Return pMarginRight
    'End Function

    'Public Function GetMarginBottom() As Single
    '    Return pMarginBottom
    'End Function

    'Public Function GetMarginLeft() As Single
    '    Return pMarginLeft
    'End Function

    'Public Function GetPaddingsHeight() As Single
    '    Return pPaddingBottom + pPaddingTop
    'End Function

    'Public Function GetPaddingsWidth() As Single
    '    Return pPaddingLeft + pPaddingRight
    'End Function

    'Public Function GetPaddingTop() As Single
    '    Return pPaddingTop
    'End Function

    'Public Function GetPaddingRighe() As Single
    '    Return pPaddingRight
    'End Function

    'Public Function GetPaddingBottom() As Single
    '    Return pPaddingBottom
    'End Function

    'Public Function GetPaddingLeft() As Single
    '    Return pPaddingLeft
    'End Function

#End Region


#Region "Rendering"

    Public Sub Redraw()
        Call updateSizeAndPosition()
        Call updateBackground()
        'Call updateBorders()
    End Sub

    Private Sub updateSizeAndPosition()
        updateSize()
        updatePosition()
    End Sub

    Private Sub updateSize()
        Me.Height = 20 'pHeight ' + IIf(pBorderBox, 0, Math.Max(pPaddingTop, 0) + Math.Max(pPaddingBottom, 0) + GetTitleBarHeigth())
        Me.Width = 30 'pWidth '+ IIf(pBorderBox, 0, Math.Max(pPaddingLeft, 0) + Math.Max(pPaddingRight, 0))
    End Sub

    Private Sub updatePosition()
        Me.Top = 6 'pTop
        Me.Left = 24 'pLeft
        'If pWindowPosition = FormStartPosition.Manual Then
        '    Call locateWindowByCoordinates()
        'ElseIf pWindowPosition = FormStartPosition.CenterScreen Then
        '    Call locateWindowInTheMiddleOfTheScreen()
        'ElseIf pWindowPosition = FormStartPosition.CenterParent Then
        '    Call locateWindowInTheMiddleOfTheScreen()
        'End If
    End Sub

    Private Sub updateBackground()
        Me.BackColor = pBackgroundColor
    End Sub

    Private Sub updateBorders()

    End Sub

    'Private Sub redrawControls()

    'End Sub

#End Region


#Region "Events"

    'Private Sub Window_Load(sender As Object, e As EventArgs) Handles MyBase.Load
    '    If Not pListener Is Nothing Then
    '        On Error Resume Next
    '        Call pListener.RaiseEvent_Load()
    '    End If
    'End Sub

    'Private Sub Window_Disposed(sender As Object, e As EventArgs) Handles Me.Disposed
    '    If Not pListener Is Nothing Then
    '        On Error Resume Next
    '        Call pListener.RaiseEvent_Disposed()
    '    End If
    'End Sub

    Private Sub Div_MouseEnter(sender As Object, e As EventArgs) Handles Me.MouseEnter
        If Not pStylesMatrix Is Nothing Then
            Dim color As Color
            color = pStylesMatrix.getBackgroundColor(StyleNodeTypeEnum.StyleNodeType_Hover)
            pBackgroundColor = color
            Call updateBackground()
        End If
    End Sub

    Private Sub Div_MouseLeave(sender As Object, e As EventArgs) Handles Me.MouseLeave
        If Not pStylesMatrix Is Nothing Then
            Dim color As Color
            color = pStylesMatrix.getBackgroundColor(StyleNodeTypeEnum.StyleNodeType_Normal)
            pBackgroundColor = color
            Call updateBackground()
        End If
    End Sub

    'Private Sub Window_MouseEnter(sender As Object, e As EventArgs) Handles Me.MouseEnter
    '    If Not pListener Is Nothing Then
    '        On Error Resume Next
    '        Call pListener.RaiseEvent_MouseEnter()
    '    End If
    'End Sub

    'Private Sub Window_MouseLeave(sender As Object, e As EventArgs) Handles Me.MouseLeave
    '    If Not pListener Is Nothing Then
    '        On Error Resume Next
    '        Call pListener.RaiseEvent_MouseLeave()
    '    End If
    'End Sub

    'Private Sub Window_GotFocus(sender As Object, e As EventArgs) Handles Me.GotFocus
    '    If Not pListener Is Nothing Then
    '        On Error Resume Next
    '        Call pListener.RaiseEvent_GotFocus()
    '    End If
    'End Sub

    'Private Sub Window_LostFocus(sender As Object, e As EventArgs) Handles Me.LostFocus
    '    If Not pListener Is Nothing Then
    '        On Error Resume Next
    '        Call pListener.RaiseEvent_LostFocus()
    '    End If
    'End Sub

    'Private Sub Window_ResizeEnd(sender As Object, e As EventArgs) Handles Me.ResizeEnd
    '    If Not pListener Is Nothing Then
    '        On Error Resume Next
    '        Call pListener.RaiseEvent_SizeChanged(Me.Width, Me.Height, Me.Left, Me.Top)
    '    End If
    'End Sub

    'Private Sub Window_Scroll(sender As Object, e As ScrollEventArgs) Handles Me.Scroll
    '    If Not pListener Is Nothing Then
    '        On Error Resume Next
    '        Call pListener.RaiseEvent_Scroll()
    '    End If
    'End Sub

#End Region


#Region "Styling"

    Public Function GetStylesManager() As StylesManager Implements IContainer.GetStylesManager, IControl.GetStylesManager
        Dim instance As StylesManager
        '------------------------------------------------------------
        instance = pParent.GetStylesManager
        If instance Is Nothing Then instance = New StylesManager
        Return instance
    End Function

    Public Sub AddStyleClass(name As String) Implements IControl.AddStyleClass
        Call pStylesMatrix.AddStyleClass(name)
    End Sub

    Public Sub RemoveStyleClass(name As String) Implements IControl.RemoveStyleClass
        Call pStylesMatrix.removeStyleClass(name)
    End Sub

    Public Sub SetStyleProperty(propertyType As Long, value As Object) Implements IControl.SetStyleProperty
        Call pStylesMatrix.AddInlineStyle(propertyType, value)
    End Sub

#End Region


#Region "Managing children"

    Public Sub AddControl(ctrl As Control)

    End Sub

    Public Sub RemoveControl(ctrl As Control)

    End Sub

#End Region


    Public Function getLog() As String
        Return pLog & vbCrLf & _
                "BackgroundColor: " & pBackgroundColor.ToString() & vbCrLf & _
                "Top: " & pTop.ToString() & vbCrLf & _
                "Left: " & pLeft.ToString() & vbCrLf & _
                "Width: " & pWidth.ToString() & vbCrLf & _
                "Height: " & pHeight.ToString()
    End Function

End Class
