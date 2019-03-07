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

    '[Controls]
    Private pControls As Collection
    Private pControlsDict As Dictionary(Of String, IControl)

    '[Style]
    Private pStylesMatrix As StylesMatrix
    Private pCurrentProperties(0 To CountStylePropertiesEnums()) As Object

    '[State]
    Private pIsRendered As Boolean
    Private pState As StyleNodeTypeEnum

#End Region


#Region "Constructor"

    Public Sub New(parent As IContainer, id As String)
        pId = id
        pParent = parent
        pHandle = Me.Handle
        pControls = New Collection
        pControlsDict = New Dictionary(Of String, IControl)
        pStylesMatrix = New StylesMatrix(Me)
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

    'Public Function GetWidth() As Single
    '    Return Me.Width
    'End Function

    'Public Function GetHeight() As Single
    '    Return Me.Height
    'End Function



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

    Private Sub updateView()
        Dim newProperties As Object() = calculateProperties()
        '---------------------------------------------------------------------------------------------------------
        Dim sizeChanged As Boolean = False                  'size of this element and therefore size and layout of children elements
        Dim positionChanged As Boolean = False              'position on the screen
        Dim insideLayoutChanged As Boolean = False          'properties that can affect layout of children elements
        Dim bordersChanged As Boolean = False               'borders
        Dim insidePropertiesChanged As Boolean = False      'i.e. background, font color - properties that doesn't affect any other controls.
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
        Call compareSingleProperty(StylePropertyEnum.StyleProperty_Padding, newProperties(StylePropertyEnum.StyleProperty_Padding), insideLayoutChanged)
        Call compareSingleProperty(StylePropertyEnum.StyleProperty_PaddingTop, newProperties(StylePropertyEnum.StyleProperty_PaddingTop), insideLayoutChanged)
        Call compareSingleProperty(StylePropertyEnum.StyleProperty_PaddingLeft, newProperties(StylePropertyEnum.StyleProperty_PaddingLeft), insideLayoutChanged)
        Call compareSingleProperty(StylePropertyEnum.StyleProperty_PaddingBottom, newProperties(StylePropertyEnum.StyleProperty_PaddingBottom), insideLayoutChanged)
        Call compareSingleProperty(StylePropertyEnum.StyleProperty_PaddingRight, newProperties(StylePropertyEnum.StyleProperty_PaddingRight), insideLayoutChanged)
        Call compareSingleProperty(StylePropertyEnum.StyleProperty_Margin, newProperties(StylePropertyEnum.StyleProperty_Margin), insideLayoutChanged)
        Call compareSingleProperty(StylePropertyEnum.StyleProperty_MarginTop, newProperties(StylePropertyEnum.StyleProperty_MarginTop), insideLayoutChanged)
        Call compareSingleProperty(StylePropertyEnum.StyleProperty_MarginLeft, newProperties(StylePropertyEnum.StyleProperty_MarginLeft), insideLayoutChanged)
        Call compareSingleProperty(StylePropertyEnum.StyleProperty_MarginBottom, newProperties(StylePropertyEnum.StyleProperty_MarginBottom), insideLayoutChanged)
        Call compareSingleProperty(StylePropertyEnum.StyleProperty_MarginRight, newProperties(StylePropertyEnum.StyleProperty_MarginRight), insideLayoutChanged)
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
        If insideLayoutChanged And Not sizeChanged Then rearrangeControls()
        If sizeChanged Then Call updateSize()
        If positionChanged Then Call updatePosition()

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
        Call updatePosition()
        If anyChanges Then Call rearrangeControls()
    End Sub

    Private Sub updateSize(Optional ByRef anyChanges As Boolean = False)
        Dim heigth As Single = calculateHeigth()
        Dim width As Single = calculateWidth()
        '------------------------------------------------------------------------------------------------------------------------------------------------------------

        If Height <> Me.Height Then
            anyChanges = True
            Me.Height = Height
        End If

        If width <> Me.Width Then
            anyChanges = True
            Me.Width = width
        End If

    End Sub

    Private Sub updatePosition()

    End Sub

    Private Function calculateHeigth(Optional ByRef anyChanges As Boolean = False) As Single
        Return 40
    End Function

    Private Function calculateWidth() As Single
        Return 100
    End Function


    Private Sub updateInsideProperties()
        Me.BackColor = pCurrentProperties(StylePropertyEnum.StyleProperty_BackgroundColor)
    End Sub

    Private Sub updateBorders()

    End Sub

    Private Sub rearrangeControls()
        Dim ctrl As IControl
        For Each ctrl In pControls
            Call ctrl.UpdateSizeAndPosition()
        Next
    End Sub

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

    Private Sub Window_MouseEnter(sender As Object, e As EventArgs) Handles Me.MouseEnter
        Call hover()
        On Error Resume Next
        Call pListener.RaiseEvent_MouseEnter()
        On Error GoTo 0
    End Sub

    Private Sub Window_MouseLeave(sender As Object, e As EventArgs) Handles Me.MouseLeave
        Call unhover()
        On Error Resume Next
        Call pListener.RaiseEvent_MouseLeave()
        On Error GoTo 0
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
        Call calculateProperties()
        Call updateView()
    End Sub

    Public Sub RemoveStyleClass(name As String) Implements IControl.RemoveStyleClass
        Call pStylesMatrix.RemoveStyleClass(name)
        Call calculateProperties()
        Call updateView()
    End Sub

    Public Sub SetStyleProperty(propertyType As Long, value As Object) Implements IControl.SetStyleProperty
        Call pStylesMatrix.AddInlineStyle(propertyType, value)
        Call calculateProperties()
        Call updateView()
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
        'Call ctrl.SetParent(Me)
        'Call pControls.Add(ctrl)
        'Call pControlsDict.Add(ctrl.GetId, ctrl)
        'Call Me.Controls.Add(ctrl)
    End Sub

    Public Sub RemoveControl(ctrl As Control)

    End Sub

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


End Class
