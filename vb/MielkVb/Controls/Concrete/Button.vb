Imports System.Windows.Forms
Imports System.Runtime.InteropServices
Imports System.Drawing

Public Class Button
    Inherits System.Windows.Forms.Button
    Implements IControl


#Region "Private variables"

    Private Const TAG_NAME As String = "button"

    '[Main properties]
    Private pParent As IContainer
    Private pWindow As Window
    Private pListener As Object
    Private pHandle As IntPtr
    Private pId As String
    Private pCaption As String
    Private pLevel As Integer

    '[Style]
    Private pStylesMatrix As StylesMatrix
    Private pCurrentProperties(0 To CountStylePropertiesEnums()) As Object
    Private pHeightValueType As ControlSizeTypeEnum
    Private pWidthValueType As ControlSizeTypeEnum
    Private pTopValueType As ControlSizeTypeEnum
    Private pLeftValueType As ControlSizeTypeEnum
    Private pImageList As ImageList

    '[State]
    Private pIsRendered As Boolean
    Private pState As StyleNodeTypeEnum

    '[Border]
    Private pShowBorder As Boolean

#End Region


#Region "Constructor"

    Public Sub New(parent As IContainer, id As String)
        pId = id
        pParent = parent
        pWindow = parent.GetWindow
        pHandle = Me.Handle
        pStylesMatrix = New StylesMatrix(Me)
        Call overrideDefaultStyle()
    End Sub

    Private Sub overrideDefaultStyle()
        Me.FlatStyle = Windows.Forms.FlatStyle.Flat
        Me.Cursor = Cursors.Hand
        'FlatAppearance.BorderColor = Color.FromArgb(51, 153, 255)
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

    Public Sub SetCaption(value As String)
        pCaption = value
        Me.Text = pCaption
    End Sub

#End Region


#Region "Getters"

    Public Function GetLevel() As Integer Implements IControl.GetLevel
        Return pLevel
    End Function

    Public Function GetParent(Optional level As Integer = 0) As IContainer Implements IControl.GetParent
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

    Public Function GetWindow() As Window Implements IControl.GetWindow
        Return pWindow
    End Function

    Public Function GetId() As String Implements IControl.GetId
        Return pId
    End Function

    Public Function GetElementTag() As String Implements IControl.GetElementTag
        Return TAG_NAME
    End Function

    Public Function GetWidth() As Single Implements IControl.GetWidth
        Return Me.Width
    End Function

    Public Function GetHeight() As Single Implements IControl.GetHeight
        Return Me.Height
    End Function

    Public Function GetSizeType(prop As StylePropertyEnum) As ControlSizeTypeEnum Implements IControl.GetSizeType
        If prop = StylePropertyEnum.StyleProperty_Width Then
            Return pWidthValueType
        ElseIf prop = StylePropertyEnum.StyleProperty_Height Then
            Return pHeightValueType
        Else
            Return ControlSizeTypeEnum.ControlSizeType_Const
        End If
    End Function

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





    Public Sub updateHeight() Implements IControl.UpdateHeight
        Dim key As String : key = Me.GetHashCode & "|" & StylePropertyEnum.StyleProperty_Height
        Dim height As Single
        '---------------------------------------------------------------------------------------------------------
        If RQ.HasKey(key) Then
            If pHeightValueType = ControlSizeTypeEnum.ControlSizeType_Auto Then
                height = Me.Height
                Me.AutoSize = True
                Me.Height = height
            ElseIf pHeightValueType = ControlSizeTypeEnum.ControlSizeType_ParentRelative Then
                'not applies for Window (it has no parent)
            End If

            '[Actual sizing]


        End If
    End Sub

    Public Sub UpdateWidth() Implements IControl.UpdateWidth
        Dim key As String : key = Me.GetHashCode & "|" & StylePropertyEnum.StyleProperty_Height
        Dim width As Single
        '---------------------------------------------------------------------------------------------------------
        If RQ.HasKey(key) Then
            If pHeightValueType = ControlSizeTypeEnum.ControlSizeType_Auto Then
                width = Me.Width
                Me.AutoSize = True
                Me.Width = width
            ElseIf pHeightValueType = ControlSizeTypeEnum.ControlSizeType_ParentRelative Then
                'not applies for Window (it has no parent)
            End If

            '[Actual sizing]


        End If
    End Sub



    Public Sub UpdateView(prop As StylePropertyEnum) Implements IControl.UpdateView
        Dim key As String : key = Me.GetHashCode & "|" & prop
        '---------------------------------------------------------------------------------------------------------
        Call VQ.removeKey(key)
        If prop = StylePropertyEnum.StyleProperty_BackgroundColor Then
        ElseIf prop = StylePropertyEnum.StyleProperty_BorderThickness Then
        ElseIf prop = StylePropertyEnum.StyleProperty_BorderColor Then
        ElseIf prop = StylePropertyEnum.StyleProperty_FontSize Then
        ElseIf prop = StylePropertyEnum.StyleProperty_FontBold Then
        ElseIf prop = StylePropertyEnum.StyleProperty_FontColor Then
        ElseIf prop = StylePropertyEnum.StyleProperty_FontFamily Then
        ElseIf prop = StylePropertyEnum.StyleProperty_HorizontalAlignment Then
        ElseIf prop = StylePropertyEnum.StyleProperty_VerticalAlignment Then
        End If
    End Sub

    Public Sub LocateInContainer(ByRef top As Single, ByRef right As Single, ByRef bottom As Single, ByRef left As Single) Implements IControl.LocateInContainer
        Stop
    End Sub

    Public Function GetWidthSizeType() As ControlSizeTypeEnum Implements IControl.GetWidthSizeType
        Return pWidthValueType
    End Function

    Public Function GetHeightSizeType() As ControlSizeTypeEnum Implements IControl.GetHeightSizeType
        Return pHeightValueType
    End Function



    'Public Sub AddToResizeQueue(prop As StylePropertyEnum) Implements IControl.AddToResizeQueue

    'End Sub

    'Public Sub UpdateView(Optional propagateDown As Boolean = False) Implements IControl.UpdateView
    '    Dim newProperties As Object() = calculateProperties()
    '    '---------------------------------------------------------------------------------------------------------
    '    Dim bordersChanged As Boolean = False               'borders
    '    Dim insidePropertiesChanged As Boolean = False      'i.e. background, font color - properties that doesn't affect any other controls.
    '    Dim imageChanged As Boolean = False
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
    '    Call compareSingleProperty(StylePropertyEnum.StyleProperty_ImageFilePath, newProperties(StylePropertyEnum.StyleProperty_ImageFilePath), imageChanged)
    '    Call compareSingleProperty(StylePropertyEnum.StyleProperty_ImageName, newProperties(StylePropertyEnum.StyleProperty_ImageName), imageChanged)
    '    Call compareSingleProperty(StylePropertyEnum.StyleProperty_ImageWidth, newProperties(StylePropertyEnum.StyleProperty_ImageWidth), imageChanged)
    '    Call compareSingleProperty(StylePropertyEnum.StyleProperty_ImageHeight, newProperties(StylePropertyEnum.StyleProperty_ImageHeight), imageChanged)
    '    Call compareSingleProperty(StylePropertyEnum.StyleProperty_ImageSize, newProperties(StylePropertyEnum.StyleProperty_ImageSize), imageChanged)

    '    If insidePropertiesChanged Then Call updateInsideProperties()
    '    If imageChanged Then Call updateImage()
    '    If bordersChanged Then Call updateBorders()

    'End Sub

    'Private Sub compareSingleProperty(propertyType As StylePropertyEnum, propertyValue As Object, ByRef changeMarker As Boolean)
    '    If Not propertyValue Is Nothing Then
    '        If propertyValue <> pCurrentProperties(propertyType) Then
    '            pCurrentProperties(propertyType) = propertyValue
    '            changeMarker = True
    '        End If
    '    End If
    'End Sub


    'Public Sub UpdateLayout(Optional propagateDown As Boolean = False) Implements IControl.UpdateLayout
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

    '    If sizeChanged Then Call UpdateSize(insideLayoutChanged)
    '    If positionChanged Then Call UpdatePosition(outsideLayoutChanged)

    '    'If sizeChanged Then Call UpdateSize(outsideLayoutChanged)
    '    'If positionChanged Then Call UpdatePosition(outsideLayoutChanged)
    '    'If outsideLayoutChanged Then Call pParent.AdjustAfterChildrenSizeChange()

    'End Sub

    'Public Sub UpdateSize(Optional ByRef anyChanges As Boolean = False) 'Implements IControl.UpdateSize
    '    Dim height As Single
    '    Dim width As Single
    '    '------------------------------------------------------------------------------------------------------------------------------------------------------------

    '    Call checkIfAnyDimensionIsAutoScaled()

    '    height = CalculateControlHeight(Me, pCurrentProperties)
    '    width = CalculateControlWidth(Me, pCurrentProperties)

    '    If height <> Me.Height Then
    '        anyChanges = True
    '        Me.Height = height
    '    End If

    '    If width <> Me.Width Then
    '        anyChanges = True
    '        Me.Width = width
    '    End If

    'End Sub

    'Private Sub checkIfAnyDimensionIsAutoScaled()
    '    pIsAutoWidth = isAutoValue(pCurrentProperties(StylePropertyEnum.StyleProperty_Width))
    '    pIsAutoHeight = isAutoValue(pCurrentProperties(StylePropertyEnum.StyleProperty_Height))
    'End Sub

    'Public Sub UpdatePosition(Optional ByRef anyChanges As Boolean = False) 'Implements IControl.UpdatePosition
    '    If pCurrentProperties(StylePropertyEnum.StyleProperty_Position) = CssPositionEnum.CssPosition_Absolute Then
    '        Me.Left = CalculateControlLeft(Me, pCurrentProperties)
    '        Me.Top = CalculateControlTop(Me, pCurrentProperties)
    '    Else

    '    End If
    'End Sub

    'Private Sub updateInsideProperties()
    '    Me.Font = createFont(pCurrentProperties)
    '    Me.ForeColor = pCurrentProperties(StylePropertyEnum.StyleProperty_FontColor)
    '    Call updateBackColor()
    'End Sub

    'Private Sub updateBackColor()
    '    Dim normalBackColor As Object : normalBackColor = If(pCurrentProperties(StylePropertyEnum.StyleProperty_BackgroundColor), Color.Transparent)
    '    Dim hoverBackColor As Object : hoverBackColor = pStylesMatrix.GetPropertyValue(StylePropertyEnum.StyleProperty_BackgroundColor, StyleNodeTypeEnum.StyleNodeType_Hover)
    '    Dim clickedBackColor As Object : clickedBackColor = pStylesMatrix.GetPropertyValue(StylePropertyEnum.StyleProperty_BackgroundColor, StyleNodeTypeEnum.StyleNodeType_Clicked)
    '    '------------------------------------------------------------------------------------------------------------------------------------------------------------
    '    Me.BackColor = normalBackColor
    '    With Me.FlatAppearance
    '        .MouseDownBackColor = If(clickedBackColor, normalBackColor)
    '        .MouseOverBackColor = If(hoverBackColor, normalBackColor)
    '    End With
    'End Sub

    'Private Sub updateImage()
    '    Const DEFAULT_IMAGE_SIZE As Single = 16
    '    '------------------------------------------------------------------------------------------------------------------------------------------------------------
    '    Dim imageFilePath As Object
    '    Dim imageWidth As Object
    '    Dim imageHeight As Object
    '    '------------------------------------------------------------------------------------------------------------------------------------------------------------

    '    imageFilePath = pCurrentProperties(StylePropertyEnum.StyleProperty_ImageFilePath)
    '    imageWidth = If(pCurrentProperties(StylePropertyEnum.StyleProperty_ImageWidth), DEFAULT_IMAGE_SIZE)
    '    imageHeight = If(pCurrentProperties(StylePropertyEnum.StyleProperty_ImageHeight), DEFAULT_IMAGE_SIZE)

    '    If Not imageFilePath Is Nothing Then
    '        If pImageList Is Nothing Then pImageList = New ImageList
    '        pImageList.ImageSize = New Size(imageWidth, imageHeight)
    '        Call pImageList.Images.Add(Bitmap.FromFile(imageFilePath))
    '        Me.ImageList = pImageList
    '        Me.ImageIndex = 1
    '        Me.TextImageRelation = TextImageRelation.ImageBeforeText
    '        Me.TextAlign = ContentAlignment.MiddleRight
    '        Me.Text = " " & pCaption
    '    Else
    '        Me.ImageList = Nothing
    '        Me.TextAlign = ContentAlignment.MiddleCenter
    '        Me.Text = pCaption
    '    End If

    'End Sub

    'Private Sub updateBorders()
    '    Dim thickness As Single
    '    Dim color As Object
    '    '------------------------------------------------------------------------------------------------------------------------------------------------------------
    '    thickness = If(pCurrentProperties(StylePropertyEnum.StyleProperty_BorderThickness), 0)
    '    If thickness Then
    '        color = If(pCurrentProperties(StylePropertyEnum.StyleProperty_BorderColor), System.Drawing.Color.Transparent)
    '        With Me.FlatAppearance
    '            .BorderSize = thickness
    '            .BorderColor = color
    '        End With
    '    End If
    'End Sub


#End Region


#Region "Events"

    Private Sub Button_GotFocus(sender As Object, e As EventArgs) Handles Me.GotFocus
        Call NotifyDefault(False)
        Call hover()
    End Sub

    Private Sub Button_KeyUp(sender As Object, e As KeyEventArgs) Handles Me.KeyUp
        If e.KeyCode = Keys.Return Then
            Call pListener.TriggerEvent("click")
        End If
    End Sub

    Private Sub Button_LostFocus(sender As Object, e As EventArgs) Handles Me.LostFocus
        Call unhover()
    End Sub

    Private Sub Button_MouseEnter(sender As Object, e As EventArgs) Handles Me.MouseEnter
        Call NotifyDefault(False)
        Call hover()
    End Sub

    Private Sub Button_MouseLeave(sender As Object, e As EventArgs) Handles Me.MouseLeave
        Call unhover()
    End Sub

    Private Sub Button_MouseUp(sender As Object, e As MouseEventArgs) Handles Me.MouseUp
        If e.X >= 0 Then
            If e.X <= Me.Width Then
                If e.Y >= 0 Then
                    If e.Y <= Me.Height Then
                        Call pListener.TriggerEvent("click")
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub Button_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        Call Invalidate()
    End Sub

#End Region



#Region "Styles - Global"
    Public Function GetStylesManager() As StylesManager Implements IControl.GetStylesManager
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
        If invalidate Then updateProperties()
    End Sub

    Public Sub AddStyleClasses(name As String, Optional invalidate As Boolean = True) Implements IControl.AddStyleClasses
        Call pStylesMatrix.AddStyleClasses(name, True)
        If invalidate Then updateProperties()
    End Sub

    Public Sub RemoveStyleClasses(name As String, Optional invalidate As Boolean = True) Implements IControl.RemoveStyleClasses
        Call pStylesMatrix.RemoveStyleClasses(name, True)
        If invalidate Then updateProperties()
    End Sub

    Public Sub SetStyleProperty(propertyType As Long, value As Object, Optional invalidate As Boolean = True) Implements IControl.SetStyleProperty
        Call pStylesMatrix.AddInlineStyle(propertyType, value, True)
        If invalidate Then updateProperties()
    End Sub

    Private Sub updateProperties()
        Stop
    End Sub

    Private Function calculateProperties() As Object
        Dim source As Object
        Dim result(0 To CountStylePropertiesEnums()) As Object
        Dim col As Long
        Dim i As Long
        '------------------------------------------------------------------------------------------------------------------------------------

        source = pStylesMatrix.GetPropertiesArray
        col = CInt(pState)
        For i = LBound(source, 2) To UBound(source, 2)
            result(i) = source(col, i)
        Next
        Return result

    End Function

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


#Region "Customize focus behaviour"

    Private Sub ButtonPaint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs)
        Dim btn = DirectCast(sender, Button)
        Using p As New Pen(Me.BackColor)
            e.Graphics.DrawRectangle(p, 1, 1, btn.Width - 3, btn.Height - 3)
        End Using
    End Sub

    Protected Overrides Sub OnMouseEnter(e As EventArgs)
        MyBase.OnMouseEnter(e)
        pShowBorder = True
    End Sub

    Protected Overrides Sub OnMouseLeave(e As EventArgs)
        MyBase.OnMouseLeave(e)
        pShowBorder = False
    End Sub

    Protected Overrides Sub OnPaint(pEvent As PaintEventArgs)
        MyBase.OnPaint(pEvent)
        If (DesignMode Or pShowBorder) Then
            Dim pen As Pen = New Pen(FlatAppearance.BorderColor, 1)
            Dim rectangle As Rectangle = New Rectangle(0, 0, Size.Width - 1, Size.Height - 1)
            Call pEvent.Graphics.DrawRectangle(pen, rectangle)
        End If
    End Sub

    Protected Overrides ReadOnly Property ShowFocusCues As Boolean
        Get
            Return False
        End Get
    End Property

#End Region


End Class
