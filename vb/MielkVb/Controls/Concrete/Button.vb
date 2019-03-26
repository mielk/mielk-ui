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

    '[Style]
    Private pStylesMatrix As StylesMatrix
    Private pCurrentProperties(0 To CountStylePropertiesEnums()) As Object
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
    End Sub

    Public Sub SetCaption(value As String)
        pCaption = value
        Me.Text = pCaption
    End Sub

#End Region


#Region "Getters"

    Public Function GetParent() As IContainer Implements IControl.GetParent
        Return pParent
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

#End Region




#Region "Rendering"

    Private Sub updateView()
        Dim newProperties As Object() = calculateProperties()
        '---------------------------------------------------------------------------------------------------------
        Dim sizeChanged As Boolean = False                  'size of this element and therefore size and layout of children elements
        Dim positionChanged As Boolean = False              'position on the screen
        Dim outsideLayoutChanged As Boolean = False         'properties that can affect layout of siblings elements
        Dim bordersChanged As Boolean = False               'borders
        Dim insidePropertiesChanged As Boolean = False      'i.e. background, font color - properties that doesn't affect any other controls.
        Dim imageChanged As Boolean = False
        '---------------------------------------------------------------------------------------------------------

        Call compareSingleProperty(StylePropertyEnum.StyleProperty_Float, newProperties(StylePropertyEnum.StyleProperty_Float), positionChanged)
        Call compareSingleProperty(StylePropertyEnum.StyleProperty_Position, newProperties(StylePropertyEnum.StyleProperty_Position), positionChanged)
        Call compareSingleProperty(StylePropertyEnum.StyleProperty_BorderBox, newProperties(StylePropertyEnum.StyleProperty_BorderBox), sizeChanged)
        Call compareSingleProperty(StylePropertyEnum.StyleProperty_Top, newProperties(StylePropertyEnum.StyleProperty_Top), positionChanged)
        Call compareSingleProperty(StylePropertyEnum.StyleProperty_Left, newProperties(StylePropertyEnum.StyleProperty_Left), positionChanged)
        Call compareSingleProperty(StylePropertyEnum.StyleProperty_Bottom, newProperties(StylePropertyEnum.StyleProperty_Bottom), positionChanged)
        Call compareSingleProperty(StylePropertyEnum.StyleProperty_Right, newProperties(StylePropertyEnum.StyleProperty_Right), positionChanged)
        Call compareSingleProperty(StylePropertyEnum.StyleProperty_Width, newProperties(StylePropertyEnum.StyleProperty_Width), positionChanged)
        Call compareSingleProperty(StylePropertyEnum.StyleProperty_MinWidth, newProperties(StylePropertyEnum.StyleProperty_MinWidth), sizeChanged)
        Call compareSingleProperty(StylePropertyEnum.StyleProperty_MaxWidth, newProperties(StylePropertyEnum.StyleProperty_MaxWidth), sizeChanged)
        Call compareSingleProperty(StylePropertyEnum.StyleProperty_Height, newProperties(StylePropertyEnum.StyleProperty_Height), sizeChanged)
        Call compareSingleProperty(StylePropertyEnum.StyleProperty_MinHeight, newProperties(StylePropertyEnum.StyleProperty_MinHeight), sizeChanged)
        Call compareSingleProperty(StylePropertyEnum.StyleProperty_MaxHeight, newProperties(StylePropertyEnum.StyleProperty_MaxHeight), sizeChanged)
        Call compareSingleProperty(StylePropertyEnum.StyleProperty_Padding, newProperties(StylePropertyEnum.StyleProperty_Padding), insidePropertiesChanged)
        Call compareSingleProperty(StylePropertyEnum.StyleProperty_PaddingTop, newProperties(StylePropertyEnum.StyleProperty_PaddingTop), insidePropertiesChanged)
        Call compareSingleProperty(StylePropertyEnum.StyleProperty_PaddingLeft, newProperties(StylePropertyEnum.StyleProperty_PaddingLeft), insidePropertiesChanged)
        Call compareSingleProperty(StylePropertyEnum.StyleProperty_PaddingBottom, newProperties(StylePropertyEnum.StyleProperty_PaddingBottom), insidePropertiesChanged)
        Call compareSingleProperty(StylePropertyEnum.StyleProperty_PaddingRight, newProperties(StylePropertyEnum.StyleProperty_PaddingRight), insidePropertiesChanged)
        Call compareSingleProperty(StylePropertyEnum.StyleProperty_Margin, newProperties(StylePropertyEnum.StyleProperty_Margin), outsideLayoutChanged)
        Call compareSingleProperty(StylePropertyEnum.StyleProperty_MarginTop, newProperties(StylePropertyEnum.StyleProperty_MarginTop), outsideLayoutChanged)
        Call compareSingleProperty(StylePropertyEnum.StyleProperty_MarginLeft, newProperties(StylePropertyEnum.StyleProperty_MarginLeft), outsideLayoutChanged)
        Call compareSingleProperty(StylePropertyEnum.StyleProperty_MarginBottom, newProperties(StylePropertyEnum.StyleProperty_MarginBottom), outsideLayoutChanged)
        Call compareSingleProperty(StylePropertyEnum.StyleProperty_MarginRight, newProperties(StylePropertyEnum.StyleProperty_MarginRight), outsideLayoutChanged)
        Call compareSingleProperty(StylePropertyEnum.StyleProperty_BackgroundColor, newProperties(StylePropertyEnum.StyleProperty_BackgroundColor), insidePropertiesChanged)
        Call compareSingleProperty(StylePropertyEnum.StyleProperty_BorderThickness, newProperties(StylePropertyEnum.StyleProperty_BorderThickness), bordersChanged)
        Call compareSingleProperty(StylePropertyEnum.StyleProperty_BorderColor, newProperties(StylePropertyEnum.StyleProperty_BorderColor), bordersChanged)
        Call compareSingleProperty(StylePropertyEnum.StyleProperty_FontSize, newProperties(StylePropertyEnum.StyleProperty_FontSize), insidePropertiesChanged)
        Call compareSingleProperty(StylePropertyEnum.StyleProperty_FontFamily, newProperties(StylePropertyEnum.StyleProperty_FontFamily), insidePropertiesChanged)
        Call compareSingleProperty(StylePropertyEnum.StyleProperty_FontColor, newProperties(StylePropertyEnum.StyleProperty_FontColor), insidePropertiesChanged)
        Call compareSingleProperty(StylePropertyEnum.StyleProperty_FontBold, newProperties(StylePropertyEnum.StyleProperty_FontBold), insidePropertiesChanged)
        Call compareSingleProperty(StylePropertyEnum.StyleProperty_HorizontalAlignment, newProperties(StylePropertyEnum.StyleProperty_HorizontalAlignment), insidePropertiesChanged)
        Call compareSingleProperty(StylePropertyEnum.StyleProperty_VerticalAlignment, newProperties(StylePropertyEnum.StyleProperty_VerticalAlignment), insidePropertiesChanged)
        Call compareSingleProperty(StylePropertyEnum.StyleProperty_ImageFilePath, newProperties(StylePropertyEnum.StyleProperty_ImageFilePath), imageChanged)
        Call compareSingleProperty(StylePropertyEnum.StyleProperty_ImageName, newProperties(StylePropertyEnum.StyleProperty_ImageName), imageChanged)
        Call compareSingleProperty(StylePropertyEnum.StyleProperty_ImageWidth, newProperties(StylePropertyEnum.StyleProperty_ImageWidth), imageChanged)
        Call compareSingleProperty(StylePropertyEnum.StyleProperty_ImageHeight, newProperties(StylePropertyEnum.StyleProperty_ImageHeight), imageChanged)
        Call compareSingleProperty(StylePropertyEnum.StyleProperty_ImageSize, newProperties(StylePropertyEnum.StyleProperty_ImageSize), imageChanged)

        If insidePropertiesChanged Then Call updateInsideProperties()
        If imageChanged Then Call updateImage()
        If bordersChanged Then Call updateBorders()
        If sizeChanged Then Call updateSize(outsideLayoutChanged)
        If positionChanged Then Call updatePosition(outsideLayoutChanged)
        If outsideLayoutChanged Then Call pParent.RearrangeControls()

    End Sub

    Private Sub compareSingleProperty(propertyType As StylePropertyEnum, propertyValue As Object, ByRef changeMarker As Boolean)
        If Not propertyValue Is Nothing Then
            If propertyValue <> pCurrentProperties(propertyType) Then
                pCurrentProperties(propertyType) = propertyValue
                changeMarker = True
            End If
        End If
    End Sub


    Public Sub UpdateSizeAndPosition(Optional ByRef anyChanges As Boolean = False) Implements IControl.UpdateSizeAndPosition
        Call updateSize(anyChanges)
        Call updatePosition(anyChanges)
    End Sub

    Private Sub updateSize(Optional ByRef anyChanges As Boolean = False)
        Dim height As Single = CalculateControlHeight(pParent, pCurrentProperties)
        Dim width As Single = CalculateControlWidth(pParent, pCurrentProperties)
        '------------------------------------------------------------------------------------------------------------------------------------------------------------

        If height <> Me.Height Then
            anyChanges = True
            Me.Height = height
        End If

        If width <> Me.Width Then
            anyChanges = True
            Me.Width = width
        End If

    End Sub

    Private Sub updatePosition(Optional ByRef anyChanges As Boolean = False)
        If pCurrentProperties(StylePropertyEnum.StyleProperty_Position) = CssPositionEnum.CssPosition_Absolute Then
            Me.Left = calculateLeft()
            Me.Top = calculateTop()
        Else

        End If
    End Sub

    Private Function calculateLeft() As Single
        Dim parentWidth As Single
        Dim left As Object
        Dim right As Object
        '------------------------------------------------------------------------------------------------------------------------------------------------------------

        left = pCurrentProperties(StylePropertyEnum.StyleProperty_Left)
        right = pCurrentProperties(StylePropertyEnum.StyleProperty_Right)

        If left Is Nothing Then
            If Not right Is Nothing Then
                parentWidth = pParent.GetWidth
                Return parentWidth - right - Me.Width
            Else
                Return 0
            End If
        Else
            Return CSng(left)
        End If

    End Function

    Private Function calculateTop() As Single
        Dim parentHeight As Single
        Dim top As Object
        Dim bottom As Object
        '------------------------------------------------------------------------------------------------------------------------------------------------------------

        top = pCurrentProperties(StylePropertyEnum.StyleProperty_Top)
        bottom = pCurrentProperties(StylePropertyEnum.StyleProperty_Bottom)

        If top Is Nothing Then
            If Not bottom Is Nothing Then
                parentHeight = pParent.GetHeight
                Return parentHeight - bottom - Me.Height
            Else
                Return 0
            End If
        Else
            Return CSng(top)
        End If

    End Function

    Private Sub updateInsideProperties()
        Me.Font = createFont()
        Me.ForeColor = pCurrentProperties(StylePropertyEnum.StyleProperty_FontColor)
        Call updateBackColor()
    End Sub

    Private Sub updateBackColor()
        Dim normalBackColor As Object : normalBackColor = If(pCurrentProperties(StylePropertyEnum.StyleProperty_BackgroundColor), Color.Transparent)
        Dim hoverBackColor As Object : hoverBackColor = pStylesMatrix.GetPropertyValue(StylePropertyEnum.StyleProperty_BackgroundColor, StyleNodeTypeEnum.StyleNodeType_Hover)
        Dim clickedBackColor As Object : clickedBackColor = pStylesMatrix.GetPropertyValue(StylePropertyEnum.StyleProperty_BackgroundColor, StyleNodeTypeEnum.StyleNodeType_Clicked)
        '------------------------------------------------------------------------------------------------------------------------------------------------------------
        Me.BackColor = normalBackColor
        With Me.FlatAppearance
            .MouseDownBackColor = If(clickedBackColor, normalBackColor)
            .MouseOverBackColor = If(hoverBackColor, normalBackColor)
        End With
    End Sub

    Private Sub updateImage()
        Const DEFAULT_IMAGE_SIZE As Single = 16
        '------------------------------------------------------------------------------------------------------------------------------------------------------------
        Dim imageFilePath As Object
        Dim imageWidth As Object
        Dim imageHeight As Object
        '------------------------------------------------------------------------------------------------------------------------------------------------------------

        imageFilePath = pCurrentProperties(StylePropertyEnum.StyleProperty_ImageFilePath)
        imageWidth = If(pCurrentProperties(StylePropertyEnum.StyleProperty_ImageWidth), DEFAULT_IMAGE_SIZE)
        imageHeight = If(pCurrentProperties(StylePropertyEnum.StyleProperty_ImageHeight), DEFAULT_IMAGE_SIZE)

        If Not imageFilePath Is Nothing Then
            If pImageList Is Nothing Then pImageList = New ImageList
            pImageList.ImageSize = New Size(imageWidth, imageHeight)
            Call pImageList.Images.Add(Bitmap.FromFile(imageFilePath))
            Me.ImageList = pImageList
            Me.ImageIndex = 1
            Me.TextImageRelation = TextImageRelation.ImageBeforeText
            Me.TextAlign = ContentAlignment.MiddleRight
            Me.Text = " " & pCaption
        Else
            Me.ImageList = Nothing
            Me.TextAlign = ContentAlignment.MiddleCenter
            Me.Text = pCaption
        End If

    End Sub

    Private Function createFont() As Font
        Dim bold As System.Drawing.FontStyle
        Dim family As String
        Dim size As Single
        '------------------------------------------------------------------------------------------------------------------------------------------------------------
        bold = IIf(pCurrentProperties(StylePropertyEnum.StyleProperty_FontBold), FontStyle.Bold, FontStyle.Regular)
        family = pCurrentProperties(StylePropertyEnum.StyleProperty_FontFamily)
        size = pCurrentProperties(StylePropertyEnum.StyleProperty_FontSize)
        Return New Font(family, size, bold)
    End Function

    Private Sub updateBorders()
        Dim thickness As Single
        Dim color As Object
        '------------------------------------------------------------------------------------------------------------------------------------------------------------
        thickness = If(pCurrentProperties(StylePropertyEnum.StyleProperty_BorderThickness), 0)
        If thickness Then
            color = If(pCurrentProperties(StylePropertyEnum.StyleProperty_BorderColor), System.Drawing.Color.Transparent)
            With Me.FlatAppearance
                .BorderSize = thickness
                .BorderColor = color
            End With
        End If
    End Sub

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

    Public Sub AddStyleClass(name As String) Implements IControl.AddStyleClass
        Call pStylesMatrix.AddStyleClass(name)
        Call updateView()
    End Sub

    Public Sub RemoveStyleClass(name As String) Implements IControl.RemoveStyleClass
        Call pStylesMatrix.RemoveStyleClass(name)
        Call updateView()
    End Sub

    Public Sub SetStyleProperty(propertyType As Long, value As Object) Implements IControl.SetStyleProperty
        Call pStylesMatrix.AddInlineStyle(propertyType, value)
        Call updateView()
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
        Call updateView()
    End Sub

    Private Sub unhover()
        pState = StyleNodeTypeEnum.StyleNodeType_Normal
        Call updateView()
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
