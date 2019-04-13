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
    Private pLevel As Integer

    '[Controls]
    Private pControls As Collection
    Private pControlsDict As Dictionary(Of String, IControl)

    '[Style]
    Private pStylesMatrix As StylesMatrix
    Private pCurrentProperties(0 To CountStylePropertiesEnums()) As Object
    Private pHeightValueType As ControlSizeTypeEnum
    Private pWidthValueType As ControlSizeTypeEnum
    Private pTopValueType As ControlSizeTypeEnum
    Private pLeftValueType As ControlSizeTypeEnum
    Private pHasAbsolutePosition As Boolean

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
        pLevel = pParent.GetLevel() + 1
    End Sub

#End Region


#Region "Getters"

    Public Function GetLevel() As Integer Implements IControl.GetLevel, IContainer.GetLevel
        Return pLevel
    End Function

    Public Function GetParent(Optional level As Integer = 0) As IContainer Implements IContainer.GetParent, IControl.GetParent
        If level = 0 Then
            Return pParent
        Else
            Dim i As Integer
            Dim parent As IContainer : parent = pParent
            For i = pLevel - 2 To level Step -1
                parent = parent.GetParent
            Next i
            Return parent
        End If
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


    Public Function GetSizeType(prop As StylePropertyEnum) As ControlSizeTypeEnum Implements IControl.GetSizeType, IContainer.GetSizeType
        If prop = StylePropertyEnum.StyleProperty_Width Then
            Return pWidthValueType
        ElseIf prop = StylePropertyEnum.StyleProperty_Height Then
            Return pHeightValueType
        Else
            Return ControlSizeTypeEnum.ControlSizeType_Const
        End If
    End Function

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

    Public Function GetMargin() As Coordinate Implements IControl.GetMargin
        Stop
    End Function

    Public Function IsAbsolutePositioned() As Boolean Implements IControl.IsAbsolutePositioned
        Return pTopValueType = ControlSizeTypeEnum.ControlSizeType_Const And pLeftValueType = ControlSizeTypeEnum.ControlSizeType_Const
    End Function

#End Region




#Region "Rendering"

    Public Function IsRendered() As Boolean Implements IControl.IsRendered
        Return pIsRendered
    End Function


    Public Sub addTasksToQueue()
        Call addSizeTasksToQueue()
        Call addViewTasksToQueue()
    End Sub

    Public Sub addSizeTasksToQueue()
        pCurrentProperties = calculateProperties()
        Call updateSizeValueTypes()
        Call updatePositionValueTypes()
        Call RQ.AddResizeTask(Me, StylePropertyEnum.StyleProperty_Width)
        Call RQ.AddResizeTask(Me, StylePropertyEnum.StyleProperty_Height)
        Call RQ.AddLayoutTask(Me)
    End Sub

    Public Sub UpdateHeight() Implements IControl.UpdateHeight, IContainer.UpdateHeight
        Dim heightUpdateKey As String : heightUpdateKey = Me.GetHashCode & "|" & StylePropertyEnum.StyleProperty_Height
        Dim widthUpdateKey As String : widthUpdateKey = Me.GetHashCode & "|" & StylePropertyEnum.StyleProperty_Width
        Dim ctrl As IControl
        Dim heightAlreadySet As Boolean
        '---------------------------------------------------------------------------------------------------------

        If RQ.HasKey(heightUpdateKey) Then
            Call RQ.removeKey(heightUpdateKey)
            If pHeightValueType = ControlSizeTypeEnum.ControlSizeType_Const Then
                Me.Height = CalculateControlHeight(Me, pCurrentProperties)
                heightAlreadySet = True
                For Each ctrl In pControls
                    Call ctrl.UpdateHeight()
                Next
            ElseIf pHeightValueType = ControlSizeTypeEnum.ControlSizeType_Auto Then
                For Each ctrl In pControls
                    Call ctrl.UpdateHeight()
                Next
            ElseIf pHeightValueType = ControlSizeTypeEnum.ControlSizeType_ParentRelative Then
                'not applies for Window (it has no parent)
            End If

            Call UpdateWidth()
            Call arrangeControls()
            If Not heightAlreadySet Then Me.Height = CalculateControlHeight(Me, pCurrentProperties)

        End If

    End Sub

    Public Sub UpdateWidth() Implements IControl.UpdateWidth, IContainer.UpdateWidth
        Dim widthUpdateKey As String : widthUpdateKey = Me.GetHashCode & "|" & StylePropertyEnum.StyleProperty_Width
        Dim heightUpdateKey As String : heightUpdateKey = Me.GetHashCode & "|" & StylePropertyEnum.StyleProperty_Height
        Dim ctrl As IControl
        Dim widthAlreadySet As Boolean
        '---------------------------------------------------------------------------------------------------------

        If RQ.HasKey(widthUpdateKey) Then
            Call RQ.removeKey(widthUpdateKey)
            If pWidthValueType = ControlSizeTypeEnum.ControlSizeType_Const Then
                Me.Width = CalculateControlWidth(Me, pCurrentProperties)
                widthAlreadySet = True
                For Each ctrl In pControls
                    Call ctrl.UpdateHeight()
                Next
            ElseIf pWidthValueType = ControlSizeTypeEnum.ControlSizeType_Auto Then
                For Each ctrl In pControls
                    Call ctrl.UpdateWidth()
                Next
            ElseIf pWidthValueType = ControlSizeTypeEnum.ControlSizeType_ParentRelative Then
                If RQ.HasKey(getParentUpdateKey(StylePropertyEnum.StyleProperty_Width)) Then
                    Call pParent.UpdateWidth()
                End If
                Me.Width = CalculateControlWidth(Me, pCurrentProperties)
            End If

            Call UpdateHeight()
            Call arrangeControls()
            If Not widthAlreadySet Then Me.Width = CalculateControlWidth(Me, pCurrentProperties)

        End If

    End Sub

    Public Sub ArrangeControls() Implements IContainer.ArrangeControls
        Dim key As String : key = Me.GetHashCode & "|Layout"
        '---------------------------------------------------------------------------------------------------------

        If RQ.HasKey(key) Then
            Call RQ.removeKey(key)
            '[Actual arranging layout]

        End If

    End Sub

    Private Sub updateSizeValueTypes(Optional ByRef heightChanged As Boolean = False, Optional ByRef widthChanged As Boolean = False)
        Dim height As ControlSizeTypeEnum
        Dim width As ControlSizeTypeEnum
        '---------------------------------------------------------------------------------------------------------

        height = GetControlSizeType(pCurrentProperties(StylePropertyEnum.StyleProperty_Height))
        width = GetControlSizeType(pCurrentProperties(StylePropertyEnum.StyleProperty_Width))

        If height <> pHeightValueType Then
            pHeightValueType = height
            heightChanged = True
        End If

        If width <> pWidthValueType Then
            pWidthValueType = width
            widthChanged = True
        End If

    End Sub

    Private Sub updatePositionValueTypes(Optional ByRef topChanged As Boolean = False, Optional ByRef leftChanged As Boolean = False)
        Dim topPositionType As ControlSizeTypeEnum
        Dim leftPositionType As ControlSizeTypeEnum
        '---------------------------------------------------------------------------------------------------------

        pHasAbsolutePosition = (pCurrentProperties(StylePropertyEnum.StyleProperty_Position) = CssPositionEnum.CssPosition_Absolute)
        topPositionType = GetControlSizeType(pCurrentProperties(StylePropertyEnum.StyleProperty_Top))
        leftPositionType = GetControlSizeType(pCurrentProperties(StylePropertyEnum.StyleProperty_Left))

        If topPositionType <> pTopValueType Then
            pTopValueType = topPositionType
            topChanged = True
        End If

        If leftPositionType <> pLeftValueType Then
            pLeftValueType = leftPositionType
            leftChanged = True
        End If

        If topPositionType = ControlSizeTypeEnum.ControlSizeType_Const Then Me.Top = CalculateControlTop(Me, pCurrentProperties)
        If leftPositionType = ControlSizeTypeEnum.ControlSizeType_Const Then Me.Left = CalculateControlLeft(Me, pCurrentProperties)

    End Sub

    Private Function getParentUpdateKey(prop As StylePropertyEnum) As String
        Return pParent.GetHashCode & "|" & prop
    End Function

    Public Function GetWidthSizeType() As ControlSizeTypeEnum Implements IControl.GetWidthSizeType
        Return pWidthValueType
    End Function

    Public Function GetHeightSizeType() As ControlSizeTypeEnum Implements IControl.GetHeightSizeType
        Return pHeightValueType
    End Function

    Public Sub LocateInContainer(ByRef top As Single, ByRef right As Single, ByRef bottom As Single, ByRef left As Single) Implements IControl.LocateInContainer
        If pHasAbsolutePosition Then
            If pTopValueType <> ControlSizeTypeEnum.ControlSizeType_Const Then Me.Top = CalculateControlTop(Me, pCurrentProperties)
            If pLeftValueType <> ControlSizeTypeEnum.ControlSizeType_Const Then Me.Left = CalculateControlLeft(Me, pCurrentProperties)
        Else
            Stop
        End If
    End Sub





    Private Sub addViewTasksToQueue()
        Call VQ.Add(Me, StylePropertyEnum.StyleProperty_BackgroundColor)
        Call VQ.Add(Me, StylePropertyEnum.StyleProperty_BorderThickness)
        Call VQ.Add(Me, StylePropertyEnum.StyleProperty_BorderColor)
        Call VQ.Add(Me, StylePropertyEnum.StyleProperty_FontSize)
        Call VQ.Add(Me, StylePropertyEnum.StyleProperty_FontBold)
        Call VQ.Add(Me, StylePropertyEnum.StyleProperty_FontColor)
        Call VQ.Add(Me, StylePropertyEnum.StyleProperty_FontFamily)
        Call VQ.Add(Me, StylePropertyEnum.StyleProperty_HorizontalAlignment)
        Call VQ.Add(Me, StylePropertyEnum.StyleProperty_VerticalAlignment)
    End Sub

    Public Sub UpdateView(prop As StylePropertyEnum) Implements IControl.UpdateView
        If prop = StylePropertyEnum.StyleProperty_BackgroundColor Then
            Call updateBackground()
        ElseIf prop = StylePropertyEnum.StyleProperty_BorderThickness Then
            Call updateBorders()
        ElseIf prop = StylePropertyEnum.StyleProperty_BorderColor Then
            Call updateBorders()
        ElseIf prop = StylePropertyEnum.StyleProperty_FontSize Then
            Call updateFont()
        ElseIf prop = StylePropertyEnum.StyleProperty_FontBold Then
            Call updateFont()
        ElseIf prop = StylePropertyEnum.StyleProperty_FontColor Then
            Call updateFont()
        ElseIf prop = StylePropertyEnum.StyleProperty_FontFamily Then
            Call updateFont()
        ElseIf prop = StylePropertyEnum.StyleProperty_HorizontalAlignment Then
            Call updateAlignment()
        ElseIf prop = StylePropertyEnum.StyleProperty_VerticalAlignment Then
            Call updateAlignment()
        End If
    End Sub

    Private Sub updateBackground()
        Dim key As String : key = Me.GetHashCode & "|" & StylePropertyEnum.StyleProperty_BackgroundColor
        '---------------------------------------------------------------------------------------------------------
        If VQ.HasKey(key) Then
            Call VQ.removeKey(key)
            Me.BackColor = pCurrentProperties(StylePropertyEnum.StyleProperty_BackgroundColor)
        End If
    End Sub

    Private Sub updateBorders()
        Dim keyPrefix As String : keyPrefix = Me.GetHashCode & "|"
        '---------------------------------------------------------------------------------------------------------

        Call VQ.removeKey(keyPrefix & StylePropertyEnum.StyleProperty_BorderColor)
        Call VQ.removeKey(keyPrefix & StylePropertyEnum.StyleProperty_BorderThickness)


        'children

    End Sub

    Private Sub updateFont()
        Dim keyPrefix As String : keyPrefix = Me.GetHashCode & "|"
        '---------------------------------------------------------------------------------------------------------

        Call VQ.removeKey(keyPrefix & StylePropertyEnum.StyleProperty_FontBold)
        Call VQ.removeKey(keyPrefix & StylePropertyEnum.StyleProperty_FontSize)
        Call VQ.removeKey(keyPrefix & StylePropertyEnum.StyleProperty_FontFamily)
        Call VQ.removeKey(keyPrefix & StylePropertyEnum.StyleProperty_FontColor)

        'children

    End Sub

    Private Sub updateAlignment()
        Dim keyPrefix As String : keyPrefix = Me.GetHashCode & "|"
        '---------------------------------------------------------------------------------------------------------

        Call VQ.removeKey(keyPrefix & StylePropertyEnum.StyleProperty_HorizontalAlignment)
        Call VQ.removeKey(keyPrefix & StylePropertyEnum.StyleProperty_VerticalAlignment)

        pBorderThickness = If(pCurrentProperties(StylePropertyEnum.StyleProperty_BorderThickness), 0)
        pBorderColor = If(pCurrentProperties(StylePropertyEnum.StyleProperty_BorderColor), System.Drawing.Color.Transparent)

    End Sub




    'Public Sub UpdateView(Optional propagateDown As Boolean = False) Implements IControl.UpdateView, IContainer.UpdateView
    '    Dim newProperties As Object() = calculateProperties()
    '    '---------------------------------------------------------------------------------------------------------
    '    Dim bordersChanged As Boolean = False               'borders
    '    Dim insidePropertiesChanged As Boolean = False      'i.e. background, font color - properties that doesn't affect any other controls.
    '    Dim ctrl As IControl
    '    '---------------------------------------------------------------------------------------------------------

    '    Call compareSingleProperty(StylePropertyEnum.StyleProperty_BackgroundColor, newProperties(StylePropertyEnum.StyleProperty_BackgroundColor), insidePropertiesChanged)
    '    Call compareSingleProperty(StylePropertyEnum.StyleProperty_BorderThickness, newProperties(StylePropertyEnum.StyleProperty_BorderThickness), bordersChanged)
    '    Call compareSingleProperty(StylePropertyEnum.StyleProperty_BorderColor, newProperties(StylePropertyEnum.StyleProperty_BorderColor), bordersChanged)
    '    Call compareSingleProperty(StylePropertyEnum.StyleProperty_FontSize, newProperties(StylePropertyEnum.StyleProperty_FontSize), insidePropertiesChanged)
    '    Call compareSingleProperty(StylePropertyEnum.StyleProperty_FontFamily, newProperties(StylePropertyEnum.StyleProperty_FontFamily), insidePropertiesChanged)
    '    Call compareSingleProperty(StylePropertyEnum.StyleProperty_FontColor, newProperties(StylePropertyEnum.StyleProperty_FontColor), insidePropertiesChanged)
    '    Call compareSingleProperty(StylePropertyEnum.StyleProperty_FontBold, newProperties(StylePropertyEnum.StyleProperty_FontBold), insidePropertiesChanged)
    '    Call compareSingleProperty(StylePropertyEnum.StyleProperty_HorizontalAlignment, newProperties(StylePropertyEnum.StyleProperty_HorizontalAlignment), insidePropertiesChanged)
    '    Call compareSingleProperty(StylePropertyEnum.StyleProperty_VerticalAlignment, newProperties(StylePropertyEnum.StyleProperty_VerticalAlignment), insidePropertiesChanged)
    '    Call compareSingleProperty(StylePropertyEnum.StyleProperty_ImageFilePath, newProperties(StylePropertyEnum.StyleProperty_ImageFilePath), insidePropertiesChanged)
    '    Call compareSingleProperty(StylePropertyEnum.StyleProperty_ImageName, newProperties(StylePropertyEnum.StyleProperty_ImageName), insidePropertiesChanged)
    '    Call compareSingleProperty(StylePropertyEnum.StyleProperty_ImageWidth, newProperties(StylePropertyEnum.StyleProperty_ImageWidth), insidePropertiesChanged)
    '    Call compareSingleProperty(StylePropertyEnum.StyleProperty_ImageHeight, newProperties(StylePropertyEnum.StyleProperty_ImageHeight), insidePropertiesChanged)
    '    Call compareSingleProperty(StylePropertyEnum.StyleProperty_ImageSize, newProperties(StylePropertyEnum.StyleProperty_ImageSize), insidePropertiesChanged)

    '    If insidePropertiesChanged Then Call updateInsideProperties()
    '    If bordersChanged Then Call updateBorders()

    '    If propagateDown Then
    '        For Each ctrl In pControls
    '            Call ctrl.UpdateView(True)
    '        Next
    '    End If

    'End Sub

    'Private Sub compareSingleProperty(propertyType As StylePropertyEnum, propertyValue As Object, ByRef changeMarker As Boolean)
    '    If Not propertyValue Is Nothing Then
    '        If propertyValue <> pCurrentProperties(propertyType) Then
    '            pCurrentProperties(propertyType) = propertyValue
    '            changeMarker = True
    '        End If
    '    End If
    'End Sub


    'Public Sub UpdateLayout(Optional propagateDown As Boolean = False) Implements IControl.UpdateLayout, IContainer.UpdateLayout
    '    Dim newProperties As Object() = calculateProperties()
    '    '---------------------------------------------------------------------------------------------------------
    '    Dim sizeChanged As Boolean = False                  'size of this element and therefore size and layout of children elements
    '    Dim positionChanged As Boolean = False              'position on the screen
    '    Dim insideLayoutChanged As Boolean = False          'properties that can affect layout of children elements
    '    Dim outsideLayoutChanged As Boolean = False         'properties that can affect layout of siblings elements
    '    '---------------------------------------------------------------------------------------------------------

    '    Call compareSingleProperty(StylePropertyEnum.StyleProperty_Float, newProperties(StylePropertyEnum.StyleProperty_Float), positionChanged)
    '    Call compareSingleProperty(StylePropertyEnum.StyleProperty_Position, newProperties(StylePropertyEnum.StyleProperty_Position), positionChanged)
    '    Call compareSingleProperty(StylePropertyEnum.StyleProperty_BorderBox, newProperties(StylePropertyEnum.StyleProperty_BorderBox), sizeChanged)
    '    Call compareSingleProperty(StylePropertyEnum.StyleProperty_Top, newProperties(StylePropertyEnum.StyleProperty_Top), positionChanged)
    '    Call compareSingleProperty(StylePropertyEnum.StyleProperty_Left, newProperties(StylePropertyEnum.StyleProperty_Left), positionChanged)
    '    Call compareSingleProperty(StylePropertyEnum.StyleProperty_Bottom, newProperties(StylePropertyEnum.StyleProperty_Bottom), positionChanged)
    '    Call compareSingleProperty(StylePropertyEnum.StyleProperty_Right, newProperties(StylePropertyEnum.StyleProperty_Right), positionChanged)
    '    Call compareSingleProperty(StylePropertyEnum.StyleProperty_Width, newProperties(StylePropertyEnum.StyleProperty_Width), sizeChanged)
    '    Call compareSingleProperty(StylePropertyEnum.StyleProperty_MinWidth, newProperties(StylePropertyEnum.StyleProperty_MinWidth), sizeChanged)
    '    Call compareSingleProperty(StylePropertyEnum.StyleProperty_MaxWidth, newProperties(StylePropertyEnum.StyleProperty_MaxWidth), sizeChanged)
    '    Call compareSingleProperty(StylePropertyEnum.StyleProperty_Height, newProperties(StylePropertyEnum.StyleProperty_Height), sizeChanged)
    '    Call compareSingleProperty(StylePropertyEnum.StyleProperty_MinHeight, newProperties(StylePropertyEnum.StyleProperty_MinHeight), sizeChanged)
    '    Call compareSingleProperty(StylePropertyEnum.StyleProperty_MaxHeight, newProperties(StylePropertyEnum.StyleProperty_MaxHeight), sizeChanged)
    '    Call compareSingleProperty(StylePropertyEnum.StyleProperty_Padding, newProperties(StylePropertyEnum.StyleProperty_Padding), insideLayoutChanged)
    '    Call compareSingleProperty(StylePropertyEnum.StyleProperty_PaddingTop, newProperties(StylePropertyEnum.StyleProperty_PaddingTop), insideLayoutChanged)
    '    Call compareSingleProperty(StylePropertyEnum.StyleProperty_PaddingLeft, newProperties(StylePropertyEnum.StyleProperty_PaddingLeft), insideLayoutChanged)
    '    Call compareSingleProperty(StylePropertyEnum.StyleProperty_PaddingBottom, newProperties(StylePropertyEnum.StyleProperty_PaddingBottom), insideLayoutChanged)
    '    Call compareSingleProperty(StylePropertyEnum.StyleProperty_PaddingRight, newProperties(StylePropertyEnum.StyleProperty_PaddingRight), insideLayoutChanged)
    '    Call compareSingleProperty(StylePropertyEnum.StyleProperty_Margin, newProperties(StylePropertyEnum.StyleProperty_Margin), outsideLayoutChanged)
    '    Call compareSingleProperty(StylePropertyEnum.StyleProperty_MarginTop, newProperties(StylePropertyEnum.StyleProperty_MarginTop), outsideLayoutChanged)
    '    Call compareSingleProperty(StylePropertyEnum.StyleProperty_MarginLeft, newProperties(StylePropertyEnum.StyleProperty_MarginLeft), outsideLayoutChanged)
    '    Call compareSingleProperty(StylePropertyEnum.StyleProperty_MarginBottom, newProperties(StylePropertyEnum.StyleProperty_MarginBottom), outsideLayoutChanged)
    '    Call compareSingleProperty(StylePropertyEnum.StyleProperty_MarginRight, newProperties(StylePropertyEnum.StyleProperty_MarginRight), outsideLayoutChanged)

    '    If sizeChanged Then Call updateSize(insideLayoutChanged)
    '    If positionChanged Then Call updatePosition(outsideLayoutChanged)
    '    'If outsideLayoutChanged Or insideLayoutChanged Then Call pParent.RearrangeControls()
    '    'If insideLayoutChanged And Not sizeChanged Then
    '    '    Call ResizeControls()
    '    '    Call RearrangeControls()
    '    'End If

    '    'Call updateSize(anyChanges)
    '    'Call updatePosition()
    '    'If anyChanges Then Call RearrangeControls()
    'End Sub

    'Private Sub updateSize(Optional ByRef anyChanges As Boolean = False) 'Implements IControl.UpdateSize
    '    'Dim height As Single
    '    'Dim width As Single
    '    ''------------------------------------------------------------------------------------------------------------------------------------------------------------

    '    'Call checkIfAnyDimensionIsAutoScaled()

    '    'height = CalculateControlHeight(Me, pCurrentProperties)
    '    'width = CalculateControlWidth(Me, pCurrentProperties)

    '    'If height <> Me.Height Then
    '    '    anyChanges = True
    '    '    Me.Height = height
    '    'End If

    '    'If width <> Me.Width Then
    '    '    anyChanges = True
    '    '    Me.Width = width
    '    'End If

    'End Sub


    'Private Sub updatePosition(Optional ByRef anyChanges As Boolean = False) 'Implements IControl.UpdatePosition
    '    Dim left As Single : left = CalculateControlLeft(Me, pCurrentProperties)
    '    Dim top As Single : top = CalculateControlTop(Me, pCurrentProperties)
    '    '------------------------------------------------------------------------------------------------------------------------------------------------------------

    '    If left <> Me.Left Then
    '        anyChanges = True
    '        Me.Left = left
    '    End If

    '    If top <> Me.Top Then
    '        anyChanges = True
    '        Me.Top = top
    '    End If

    'End Sub

    'Private Sub updateInsideProperties()
    '    Me.BackColor = pCurrentProperties(StylePropertyEnum.StyleProperty_BackgroundColor)
    'End Sub

    'Private Sub updateBorders()
    '    pBorderThickness = If(pCurrentProperties(StylePropertyEnum.StyleProperty_BorderThickness), 0)
    '    pBorderColor = If(pCurrentProperties(StylePropertyEnum.StyleProperty_BorderColor), System.Drawing.Color.Transparent)
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

    Public Sub SetStyleClasses(classes As String, Optional invalidate As Boolean = True) Implements IControl.SetStyleClasses
        With pStylesMatrix
            Call .RemoveAllStyleClasses(False)
            Call .AddStyleClasses(classes, True)
        End With
        If invalidate Then addTasksToQueue()
    End Sub

    Public Sub AddStyleClasses(name As String, Optional invalidate As Boolean = True) Implements IControl.AddStyleClasses
        Call pStylesMatrix.AddStyleClasses(name, True)
        If invalidate Then addTasksToQueue()
    End Sub

    Public Sub RemoveStyleClasses(name As String, Optional invalidate As Boolean = True) Implements IControl.RemoveStyleClasses
        Call pStylesMatrix.RemoveStyleClasses(name, True)
        If invalidate Then addTasksToQueue()
    End Sub

    Public Sub SetStyleProperty(propertyType As Long, value As Object, Optional invalidate As Boolean = True) Implements IControl.SetStyleProperty
        Call pStylesMatrix.AddInlineStyle(propertyType, value, True)
        If invalidate Then addTasksToQueue()
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
        'Call UpdateView()
    End Sub

    Private Sub unhover()
        pState = StyleNodeTypeEnum.StyleNodeType_Normal
        'Call UpdateView()
    End Sub

#End Region


End Class
