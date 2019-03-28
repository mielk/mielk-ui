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
    Private pWindow As Window
    Private pListener As Object
    Private pHandle As IntPtr
    Private pId As String

    '[Controls]
    Private pControls As Collection
    Private pControlsDict As Dictionary(Of String, IControl)

    '[Style]
    Private pStylesMatrix As StylesMatrix
    Private pCurrentProperties(0 To CountStylePropertiesEnums()) As Object
    Private pIsAutoWidth As Boolean
    Private pIsAutoHeight As Boolean

    '[State]
    Private pIsRendered As Boolean
    Private pState As StyleNodeTypeEnum

    '[Borders]
    Private pBorderThickness As Single
    Private pBorderColor As Color

#End Region


#Region "Constructor"

    Public Sub New(parent As IContainer, id As String)
        Me.BorderStyle = Windows.Forms.BorderStyle.None
        pId = id
        pParent = parent
        pWindow = pParent.GetWindow
        pHandle = Me.Handle
        pControls = New Collection
        pControlsDict = New Dictionary(Of String, IControl)
        pStylesMatrix = New StylesMatrix(Me)
    End Sub

#End Region


#Region "Custom borders"
    Public Overridable Sub MyPanel_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        e.Graphics.DrawRectangle(New Pen(pBorderColor, 2 * pBorderThickness), Me.ClientRectangle)
    End Sub
#End Region


#Region "Setters"

    '[Connection with VBA]
    Public Sub SetListener(value As Object)
        pListener = value
    End Sub

    Public Sub SetParent(parent As IContainer) Implements IControl.SetParent
        pParent = parent
    End Sub

#End Region


#Region "Getters"

    Public Function GetParent() As IContainer Implements IContainer.GetParent, IControl.GetParent
        Return pParent
    End Function

    Public Function GetWindow() As Window Implements IContainer.GetWindow, IControl.GetWindow
        Return pWindow
    End Function

    Public Function GetId() As String Implements IControl.GetId
        Return pId
    End Function

    Public Function GetElementTag() As String Implements IControl.GetElementTag
        Return TAG_NAME
    End Function

    'Public Function GetBorderBox() As Boolean
    '    Return pBorderBox
    'End Function

    'Public Function GetTop() As Single
    '    Return Me.Top
    'End Function

    'Public Function GetBottom() As Single
    '    Return Me.Bottom
    'End Function

    'Public Function GetLeft() As Single
    '    Return Me.Left
    'End Function

    'Public Function GetRight() As Single
    '    Return Me.Right
    'End Function



    Public Function GetWidth() As Single Implements IContainer.GetWidth, IControl.GetWidth
        Return Me.Width
    End Function

    Public Function GetInnerWidth() As Single Implements IContainer.GetInnerWidth
        Return Me.Width - getPaddingsWidth()
    End Function

    Public Function GetHeight() As Single Implements IContainer.GetHeight, IControl.GetHeight
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

#End Region




#Region "Rendering"

    Public Sub UpdateView(Optional propagateDown As Boolean = False) Implements IControl.UpdateView, IContainer.UpdateView
        Dim newProperties As Object() = calculateProperties()
        '---------------------------------------------------------------------------------------------------------
        Dim bordersChanged As Boolean = False               'borders
        Dim insidePropertiesChanged As Boolean = False      'i.e. background, font color - properties that doesn't affect any other controls.
        Dim ctrl As IControl
        '---------------------------------------------------------------------------------------------------------

        Call compareSingleProperty(StylePropertyEnum.StyleProperty_BackgroundColor, newProperties(StylePropertyEnum.StyleProperty_BackgroundColor), insidePropertiesChanged)
        Call compareSingleProperty(StylePropertyEnum.StyleProperty_BorderThickness, newProperties(StylePropertyEnum.StyleProperty_BorderThickness), bordersChanged)
        Call compareSingleProperty(StylePropertyEnum.StyleProperty_BorderColor, newProperties(StylePropertyEnum.StyleProperty_BorderColor), bordersChanged)
        Call compareSingleProperty(StylePropertyEnum.StyleProperty_FontSize, newProperties(StylePropertyEnum.StyleProperty_FontSize), insidePropertiesChanged)
        Call compareSingleProperty(StylePropertyEnum.StyleProperty_FontFamily, newProperties(StylePropertyEnum.StyleProperty_FontFamily), insidePropertiesChanged)
        Call compareSingleProperty(StylePropertyEnum.StyleProperty_FontColor, newProperties(StylePropertyEnum.StyleProperty_FontColor), insidePropertiesChanged)
        Call compareSingleProperty(StylePropertyEnum.StyleProperty_FontBold, newProperties(StylePropertyEnum.StyleProperty_FontBold), insidePropertiesChanged)
        Call compareSingleProperty(StylePropertyEnum.StyleProperty_HorizontalAlignment, newProperties(StylePropertyEnum.StyleProperty_HorizontalAlignment), insidePropertiesChanged)
        Call compareSingleProperty(StylePropertyEnum.StyleProperty_VerticalAlignment, newProperties(StylePropertyEnum.StyleProperty_VerticalAlignment), insidePropertiesChanged)
        Call compareSingleProperty(StylePropertyEnum.StyleProperty_ImageFilePath, newProperties(StylePropertyEnum.StyleProperty_ImageFilePath), insidePropertiesChanged)
        Call compareSingleProperty(StylePropertyEnum.StyleProperty_ImageName, newProperties(StylePropertyEnum.StyleProperty_ImageName), insidePropertiesChanged)
        Call compareSingleProperty(StylePropertyEnum.StyleProperty_ImageWidth, newProperties(StylePropertyEnum.StyleProperty_ImageWidth), insidePropertiesChanged)
        Call compareSingleProperty(StylePropertyEnum.StyleProperty_ImageHeight, newProperties(StylePropertyEnum.StyleProperty_ImageHeight), insidePropertiesChanged)
        Call compareSingleProperty(StylePropertyEnum.StyleProperty_ImageSize, newProperties(StylePropertyEnum.StyleProperty_ImageSize), insidePropertiesChanged)

        If insidePropertiesChanged Then Call updateInsideProperties()
        If bordersChanged Then Call updateBorders()

        If propagateDown Then
            For Each ctrl In pControls
                Call ctrl.UpdateView(True)
            Next
        End If

    End Sub

    Private Sub compareSingleProperty(propertyType As StylePropertyEnum, propertyValue As Object, ByRef changeMarker As Boolean)
        If Not propertyValue Is Nothing Then
            If propertyValue <> pCurrentProperties(propertyType) Then
                pCurrentProperties(propertyType) = propertyValue
                changeMarker = True
            End If
        End If
    End Sub


    Public Sub UpdateLayout(Optional propagateDown As Boolean = False) Implements IControl.UpdateLayout, IContainer.UpdateLayout
        Dim newProperties As Object() = calculateProperties()
        '---------------------------------------------------------------------------------------------------------
        Dim sizeChanged As Boolean = False                  'size of this element and therefore size and layout of children elements
        Dim positionChanged As Boolean = False              'position on the screen
        Dim insideLayoutChanged As Boolean = False          'properties that can affect layout of children elements
        Dim outsideLayoutChanged As Boolean = False         'properties that can affect layout of siblings elements
        '---------------------------------------------------------------------------------------------------------

        Call compareSingleProperty(StylePropertyEnum.StyleProperty_Float, newProperties(StylePropertyEnum.StyleProperty_Float), positionChanged)
        Call compareSingleProperty(StylePropertyEnum.StyleProperty_Position, newProperties(StylePropertyEnum.StyleProperty_Position), positionChanged)
        Call compareSingleProperty(StylePropertyEnum.StyleProperty_BorderBox, newProperties(StylePropertyEnum.StyleProperty_BorderBox), sizeChanged)
        Call compareSingleProperty(StylePropertyEnum.StyleProperty_Top, newProperties(StylePropertyEnum.StyleProperty_Top), positionChanged)
        Call compareSingleProperty(StylePropertyEnum.StyleProperty_Left, newProperties(StylePropertyEnum.StyleProperty_Left), positionChanged)
        Call compareSingleProperty(StylePropertyEnum.StyleProperty_Bottom, newProperties(StylePropertyEnum.StyleProperty_Bottom), positionChanged)
        Call compareSingleProperty(StylePropertyEnum.StyleProperty_Right, newProperties(StylePropertyEnum.StyleProperty_Right), positionChanged)
        Call compareSingleProperty(StylePropertyEnum.StyleProperty_Width, newProperties(StylePropertyEnum.StyleProperty_Width), sizeChanged)
        Call compareSingleProperty(StylePropertyEnum.StyleProperty_MinWidth, newProperties(StylePropertyEnum.StyleProperty_MinWidth), sizeChanged)
        Call compareSingleProperty(StylePropertyEnum.StyleProperty_MaxWidth, newProperties(StylePropertyEnum.StyleProperty_MaxWidth), sizeChanged)
        Call compareSingleProperty(StylePropertyEnum.StyleProperty_Height, newProperties(StylePropertyEnum.StyleProperty_Height), sizeChanged)
        Call compareSingleProperty(StylePropertyEnum.StyleProperty_MinHeight, newProperties(StylePropertyEnum.StyleProperty_MinHeight), sizeChanged)
        Call compareSingleProperty(StylePropertyEnum.StyleProperty_MaxHeight, newProperties(StylePropertyEnum.StyleProperty_MaxHeight), sizeChanged)
        Call compareSingleProperty(StylePropertyEnum.StyleProperty_Padding, newProperties(StylePropertyEnum.StyleProperty_Padding), insideLayoutChanged)
        Call compareSingleProperty(StylePropertyEnum.StyleProperty_PaddingTop, newProperties(StylePropertyEnum.StyleProperty_PaddingTop), insideLayoutChanged)
        Call compareSingleProperty(StylePropertyEnum.StyleProperty_PaddingLeft, newProperties(StylePropertyEnum.StyleProperty_PaddingLeft), insideLayoutChanged)
        Call compareSingleProperty(StylePropertyEnum.StyleProperty_PaddingBottom, newProperties(StylePropertyEnum.StyleProperty_PaddingBottom), insideLayoutChanged)
        Call compareSingleProperty(StylePropertyEnum.StyleProperty_PaddingRight, newProperties(StylePropertyEnum.StyleProperty_PaddingRight), insideLayoutChanged)
        Call compareSingleProperty(StylePropertyEnum.StyleProperty_Margin, newProperties(StylePropertyEnum.StyleProperty_Margin), outsideLayoutChanged)
        Call compareSingleProperty(StylePropertyEnum.StyleProperty_MarginTop, newProperties(StylePropertyEnum.StyleProperty_MarginTop), outsideLayoutChanged)
        Call compareSingleProperty(StylePropertyEnum.StyleProperty_MarginLeft, newProperties(StylePropertyEnum.StyleProperty_MarginLeft), outsideLayoutChanged)
        Call compareSingleProperty(StylePropertyEnum.StyleProperty_MarginBottom, newProperties(StylePropertyEnum.StyleProperty_MarginBottom), outsideLayoutChanged)
        Call compareSingleProperty(StylePropertyEnum.StyleProperty_MarginRight, newProperties(StylePropertyEnum.StyleProperty_MarginRight), outsideLayoutChanged)

        If sizeChanged Then Call updateSize(insideLayoutChanged)
        If positionChanged Then Call updatePosition(outsideLayoutChanged)
        'If outsideLayoutChanged Or insideLayoutChanged Then Call pParent.RearrangeControls()
        'If insideLayoutChanged And Not sizeChanged Then
        '    Call ResizeControls()
        '    Call RearrangeControls()
        'End If

        'Call updateSize(anyChanges)
        'Call updatePosition()
        'If anyChanges Then Call RearrangeControls()
    End Sub

    Private Sub updateSize(Optional ByRef anyChanges As Boolean = False) 'Implements IControl.UpdateSize
        Dim height As Single
        Dim width As Single
        '------------------------------------------------------------------------------------------------------------------------------------------------------------

        Call checkIfAnyDimensionIsAutoScaled()

        height = CalculateControlHeight(Me, pCurrentProperties)
        width = CalculateControlWidth(Me, pCurrentProperties)

        If height <> Me.Height Then
            anyChanges = True
            Me.Height = height
        End If

        If width <> Me.Width Then
            anyChanges = True
            Me.Width = width
        End If

    End Sub

    Private Sub checkIfAnyDimensionIsAutoScaled()
        pIsAutoWidth = isAutoValue(pCurrentProperties(StylePropertyEnum.StyleProperty_Width))
        pIsAutoHeight = isAutoValue(pCurrentProperties(StylePropertyEnum.StyleProperty_Height))
    End Sub



    Private Sub updatePosition(Optional ByRef anyChanges As Boolean = False) 'Implements IControl.UpdatePosition
        Dim left As Single : left = CalculateControlLeft(Me, pCurrentProperties)
        Dim top As Single : top = CalculateControlTop(Me, pCurrentProperties)
        '------------------------------------------------------------------------------------------------------------------------------------------------------------

        If left <> Me.Left Then
            anyChanges = True
            Me.Left = left
        End If

        If top <> Me.Top Then
            anyChanges = True
            Me.Top = top
        End If

    End Sub

    Private Sub updateInsideProperties()
        Me.BackColor = pCurrentProperties(StylePropertyEnum.StyleProperty_BackgroundColor)
    End Sub

    Private Sub updateBorders()
        pBorderThickness = If(pCurrentProperties(StylePropertyEnum.StyleProperty_BorderThickness), 0)
        pBorderColor = If(pCurrentProperties(StylePropertyEnum.StyleProperty_BorderColor), System.Drawing.Color.Transparent)
    End Sub

    'Public Sub RearrangeControls() Implements IContainer.RearrangeControls
    '    Dim ctrl As IControl
    '    For Each ctrl In pControls
    '        Call ctrl.UpdateLayout()
    '    Next
    'End Sub

    'Public Sub ResizeControls() Implements IContainer.ResizeControls
    '    Dim ctrl As IControl
    '    For Each ctrl In pControls
    '        Call ctrl.UpdateSize()
    '    Next
    'End Sub

    Public Sub AdjustAfterChildrenSizeChange() Implements IContainer.AdjustAfterChildrenSizeChange
        Dim width As Object
        Dim height As Object
        Dim anyChanges As Boolean
        '---------------------------------------------------------------------------------------------------------

        width = pCurrentProperties(StylePropertyEnum.StyleProperty_Width)
        height = pCurrentProperties(StylePropertyEnum.StyleProperty_Height)

        If Not IsNumeric(width) Then
            If width = AUTO Then
                width = CalculateAutoWidth()
                If width <> Me.Width Then
                    Me.Width = width
                    anyChanges = True
                End If
                'Stop
            End If
        End If

        If Not IsNumeric(height) Then
            If height = AUTO Then
                height = CalculateAutoHeight()
                If height <> Me.Height Then
                    Me.Height = height
                    anyChanges = True
                End If
            End If
        End If

        If anyChanges Then Call pParent.AdjustAfterChildrenSizeChange()

    End Sub

    Public Function CalculateAutoHeight() As Single Implements IContainer.CalculateAutoHeight
        Dim control As IControl
        Dim height As Single
        Dim minHeight As Object
        Dim maxHeight As Object
        '------------------------------------------------------------------------------------------------------------------------------------------------------------

        minHeight = pCurrentProperties(StylePropertyEnum.StyleProperty_MinHeight)
        maxHeight = pCurrentProperties(StylePropertyEnum.StyleProperty_MaxHeight)
        For Each control In pControls
            height = control.GetHeight
        Next

        If Not minHeight Is Nothing Then
            If height < minHeight Then height = minHeight
        End If

        If Not maxHeight Is Nothing Then
            If height > maxHeight Then height = maxHeight
        End If

        Return height

    End Function

    Public Function CalculateAutoWidth() As Single Implements IContainer.CalculateAutoWidth
        Return 0
    End Function

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

    'Private Sub Div_MouseEnter(sender As Object, e As EventArgs) Handles Me.MouseEnter
    '    If Not pStylesMatrix Is Nothing Then
    '        Dim color As Color
    '        color = pStylesMatrix.getBackgroundColor(StyleNodeTypeEnum.StyleNodeType_Hover)
    '        pBackgroundColor = color
    '        Call updateBackground()
    '    End If
    'End Sub

    'Private Sub Div_MouseLeave(sender As Object, e As EventArgs) Handles Me.MouseLeave
    '    If Not pStylesMatrix Is Nothing Then
    '        Dim color As Color
    '        color = pStylesMatrix.getBackgroundColor(StyleNodeTypeEnum.StyleNodeType_Normal)
    '        pBackgroundColor = color
    '        Call updateBackground()
    '    End If
    'End Sub

    Private Sub Div_MouseEnter(sender As Object, e As EventArgs) Handles Me.MouseEnter
        Call hover()
    End Sub

    Private Sub Div_MouseLeave(sender As Object, e As EventArgs) Handles Me.MouseLeave
        Call unhover()
    End Sub

    Private Sub Div_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        Call Invalidate()
    End Sub


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



#Region "Styles - Global"
    Public Function GetStylesManager() As StylesManager Implements IContainer.GetStylesManager, IControl.GetStylesManager
        Static instance As StylesManager
        '------------------------------------------------------------
        If instance Is Nothing Then instance = pParent.GetStylesManager
        Return instance
    End Function
#End Region

#Region "Styles - Local"

    Public Sub AddStyleClass(name As String) Implements IControl.AddStyleClass
        Call pStylesMatrix.AddStyleClass(name)
    End Sub

    Public Sub RemoveStyleClass(name As String) Implements IControl.RemoveStyleClass
        Call pStylesMatrix.RemoveStyleClass(name)
    End Sub

    Public Sub SetStyleProperty(propertyType As Long, value As Object) Implements IControl.SetStyleProperty
        Call pStylesMatrix.AddInlineStyle(propertyType, value)
    End Sub

    Private Function calculateProperties() As Object
        Dim source As Object
        Dim result(0 To CountStylePropertiesEnums()) As Object
        Dim col As Long
        Dim i As Long
        '------------------------------------------------------------------------------------------------------------------------------------

        source = pStylesMatrix.GetPropertiesArray
        col = IIf(pState = StyleNodeTypeEnum.StyleNodeType_Hover, StyleNodeTypeEnum.StyleNodeType_Hover, StyleNodeTypeEnum.StyleNodeType_Normal)
        For i = LBound(source, 2) To UBound(source, 2)
            result(i) = source(col, i)
        Next
        Return result

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


#Region "Visual state changes"

    Private Sub hover()
        pState = StyleNodeTypeEnum.StyleNodeType_Hover
        Call UpdateView()
    End Sub

    Private Sub unhover()
        pState = StyleNodeTypeEnum.StyleNodeType_Normal
        Call UpdateView()
    End Sub

#End Region


End Class
