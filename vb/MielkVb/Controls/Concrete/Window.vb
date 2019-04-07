Imports System.Windows.Forms
Imports System.Runtime.InteropServices
Imports System.Drawing

Public Class Window
    Inherits Form
    Implements IContainer
    Implements IControl


#Region "Windows API & constants"

    Private Enum GWL
        GWL_WNDPROC = -4
        GWL_HINSTANCE = -6
        GWL_HWNDPARENT = -8
        GWL_STYLE = -16
        GWL_EXSTYLE = -20
        GWL_USERDATA = -21
        GWL_ID = -12
    End Enum

    Private Enum WindowLongFlags As Integer
        GWL_EXSTYLE = -20
        GWLP_HINSTANCE = -6
        GWLP_HWNDPARENT = -8
        GWL_ID = -12
        GWL_STYLE = -16
        GWL_USERDATA = -21
        GWL_WNDPROC = -4
        DWLP_USER = &H8
        DWLP_MSGRESULT = &H0
        DWLP_DLGPROC = &H4
    End Enum

    Enum ShowWindowCommands As Integer
        ''' <summary>
        ''' Hides the window and activates another window.
        ''' </summary>
        Hide = 0
        ''' <summary>
        ''' Activates and displays a window. If the window is minimized or 
        ''' maximized, the system restores it to its original size and position.
        ''' An application should specify this flag when displaying the window 
        ''' for the first time.
        ''' </summary>
        Normal = 1
        ''' <summary>
        ''' Activates the window and displays it as a minimized window.
        ''' </summary>
        ShowMinimized = 2
        ''' <summary>
        ''' Maximizes the specified window.
        ''' </summary>
        Maximize = 3
        ' is this the right value?
        ''' <summary>
        ''' Activates the window and displays it as a maximized window.
        ''' </summary>       
        ShowMaximized = 3
        ''' <summary>
        ''' Displays a window in its most recent size and position. This value 
        ''' is similar to <see cref="Win32.ShowWindowCommands.Normal"/>, except 
        ''' the window is not actived.
        ''' </summary>
        ShowNoActivate = 4
        ''' <summary>
        ''' Activates the window and displays it in its current size and position. 
        ''' </summary>
        Show = 5
        ''' <summary>
        ''' Minimizes the specified window and activates the next top-level 
        ''' window in the Z order.
        ''' </summary>
        Minimize = 6
        ''' <summary>
        ''' Displays the window as a minimized window. This value is similar to
        ''' <see cref="Win32.ShowWindowCommands.ShowMinimized"/>, except the 
        ''' window is not activated.
        ''' </summary>
        ShowMinNoActive = 7
        ''' <summary>
        ''' Displays the window in its current size and position. This value is 
        ''' similar to <see cref="Win32.ShowWindowCommands.Show"/>, except the 
        ''' window is not activated.
        ''' </summary>
        ShowNA = 8
        ''' <summary>
        ''' Activates and displays the window. If the window is minimized or 
        ''' maximized, the system restores it to its original size and position. 
        ''' An application should specify this flag when restoring a minimized window.
        ''' </summary>
        Restore = 9
        ''' <summary>
        ''' Sets the show state based on the SW_* value specified in the 
        ''' STARTUPINFO structure passed to the CreateProcess function by the 
        ''' program that started the application.
        ''' </summary>
        ShowDefault = 10
        ''' <summary>
        '''  <b>Windows 2000/XP:</b> Minimizes a window, even if the thread 
        ''' that owns the window is not responding. This flag should only be 
        ''' used when minimizing windows from a different thread.
        ''' </summary>
        ForceMinimize = 11
    End Enum

    <DllImport("user32.dll", SetLastError:=True)> _
    Public Shared Function ShowWindow(hWnd As IntPtr, <MarshalAs(UnmanagedType.I4)> nCmdShow As ShowWindowCommands) As <MarshalAs(UnmanagedType.Bool)> Boolean
    End Function

#End Region


#Region "Private variables"

    Private Const TAG_NAME As String = "form"

    Private pListener As Object
    Private pHandle As IntPtr
    Private pId As String
    Private pStylesManager As StylesManager
    Private pLevel As Integer

    '[Main properties]
    Private pVisible As Boolean
    Private pIsModal As Boolean
    Private pBorderBox As Boolean
    Private pWindowSizeMode As Long
    Private pWindowPosition As Long
    Private pHasBorder As Boolean
    Private pHasTitleBar As Boolean
    Private pIsSizable As Boolean
    Private pTitle As String

    '[Controls]
    Private pControls As Collection
    Private pControlsDict As Dictionary(Of String, IControl)

    '[State]
    Private pIsDisplayed As Boolean
    Private pIsRendered As Boolean
    Private pState As StyleNodeTypeEnum

    '[Size and position]
    Private pStylesMatrix As StylesMatrix
    Private pCurrentProperties(0 To CountStylePropertiesEnums()) As Object
    Private pHeightValueType As ControlSizeTypeEnum
    Private pWidthValueType As ControlSizeTypeEnum
    Private pTopValueType As ControlSizeTypeEnum
    Private pLeftValueType As ControlSizeTypeEnum
    
    Private pPropertiesToUpdate As Collection


#End Region


#Region "Constructor"

    Public Sub New()
        Visible = pVisible
        pHandle = Me.Handle
        pControls = New Collection
        pControlsDict = New Dictionary(Of String, IControl)
        pPropertiesToUpdate = New Collection
        pLevel = 1
    End Sub

    Private Sub InitializeComponent()
        Me.SuspendLayout()
        '
        'Window
        '
        Me.ClientSize = New System.Drawing.Size(284, 261)
        Me.Name = "Window"
        Me.ResumeLayout(False)

    End Sub

#End Region


#Region "Window size & position"

    Public Sub Display(Optional ByVal isModal As Boolean = False)
        pIsDisplayed = True
        pIsModal = isModal
        Call updateTitleBarVisibility()
        Call Render()
        Call UpdateSize()
        'Call UpdateView(True)
        'Call UpdateLayout(True)
        If pIsModal Then
            Call Me.ShowDialog()
        Else
            Call Me.Show()
        End If
    End Sub

    Public Overloads Sub Hide()
        pVisible = False
        MyBase.Hide()
    End Sub

    Public Sub Maximize()
        If pIsDisplayed Then
            Call SetWindowSizeMode(ShowWindowCommands.ShowMaximized)
        End If
    End Sub

    Public Sub Minimize()
        If pIsDisplayed Then
            Call SetWindowSizeMode(ShowWindowCommands.Minimize)
        End If
    End Sub

    Public Sub Normalize()
        If pIsDisplayed Then
            Call SetWindowSizeMode(ShowWindowCommands.Normal)
        End If
    End Sub



    Public Sub SetWindowPosition(value As Long)
        pWindowPosition = value
        Call updatePosition()
    End Sub

    Public Function GetWindowPosition() As Long
        Return pWindowPosition
    End Function

#End Region


#Region "Window properties - Setters"

    '[Connection with VBA]
    Public Sub SetId(value As String)
        pId = value
    End Sub

    Public Sub SetListener(value As Object)
        pListener = value
    End Sub

    Public Sub SetParent(parent As IContainer) Implements IControl.SetParent
        '
    End Sub

    Public Sub SetIsModal(value As Boolean)
        pIsModal = value
    End Sub

    Public Sub SetTitle(value As String)
        pTitle = value
        Call updateTitleBarVisibility()
    End Sub

    Public Sub SetHasBorders(value As Boolean)
        pHasBorder = value
        Call updateBorders()
    End Sub

    Public Sub SetHasTitleBar(value As Boolean)
        pHasTitleBar = value
        Call updateTitleBarVisibility()
    End Sub

    Public Sub SetIsSizable(value As Boolean)
        pIsSizable = value
        Call updateBorders()
    End Sub

    Public Sub SetBorderBox(value As Boolean)
        pBorderBox = value
        Err.Raise(1)
        'Call updateSize()
        'Call updatePosition()
    End Sub

    Public Sub SetWindowSizeMode(value As ShowWindowCommands)
        pWindowSizeMode = value
        If pIsDisplayed Then
            Call ShowWindow(pHandle, value)
            Err.Raise(1)
            'Call rearrangeControls()
        End If
    End Sub

    Public Sub SetIcon(ByVal path As String)
        Me.Icon = New System.Drawing.Icon(path)
    End Sub

#End Region


#Region "Window properties - Getters"

    Public Function GetLevel() As Integer Implements IControl.GetLevel, IContainer.GetLevel
        Return pLevel
    End Function

    Public Function GetParent(Optional level As Integer = 0) As IContainer Implements IContainer.GetParent, IControl.GetParent
        Return Nothing
    End Function

    Public Function GetWindow() As Window Implements IContainer.GetWindow, IControl.GetWindow
        Return Me
    End Function

    Public Function GetId() As String Implements IControl.GetId
        Return pId
    End Function

    Public Function GetElementTag() As String Implements IControl.GetElementTag
        Return TAG_NAME
    End Function

    Public Function GetTitle() As Single
        Return Me.Text
    End Function

    Public Function HasBorder() As Boolean
        Return (Me.FormBorderStyle <> Windows.Forms.FormBorderStyle.None)
    End Function

    Public Function HasTitleBar() As Boolean
        If (Me.ControlBox = True Or Len(Me.Text) > 0) Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function IsSizable() As Boolean
        Return pIsSizable
    End Function

    Public Function GetBorderBox() As Boolean
        Return pBorderBox
    End Function

    Public Function IsModal() As Boolean
        Return pIsModal
    End Function

    Public Function GetWindowSizeMode() As Long
        Return pWindowSizeMode
    End Function

#End Region


#Region "UI properties - Getters"

    '[Position]
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

    Public Function GetWidth() As Single Implements IContainer.GetWidth, IControl.GetWidth
        Return Me.ClientSize.Width
    End Function

    Public Function GetInnerWidth() As Single Implements IContainer.GetInnerWidth
        Return Me.ClientSize.Width - GetPaddingsWidth()
    End Function

    Public Function GetHeight() As Single Implements IContainer.GetHeight, IControl.GetHeight
        Return Me.ClientSize.Height
    End Function

    Public Function GetInnerHeight() As Single Implements IContainer.GetInnerHeight
        Return Me.ClientSize.Height - GetPaddingsHeight()
    End Function

    Public Function GetSizeType(prop As StylePropertyEnum) As ControlSizeTypeEnum Implements IControl.GetSizeType, IContainer.GetSizeType
        If prop = StylePropertyEnum.StyleProperty_Width Then
            Return pWidthValueType
        ElseIf prop = StylePropertyEnum.StyleProperty_Height Then
            Return pHeightValueType
        Else
            Return ControlSizeTypeEnum.ControlSizeType_Const
        End If
    End Function


    '[Borders]
    Public Function GetBordersHeight() As Single
        Return Me.Width - Me.ClientSize.Width
    End Function

    Public Function GetBordersWidth() As Single
        Return Me.Width - Me.ClientSize.Width
    End Function

    Public Function GetTitleBarHeigth() As Single
        Return Me.Height - Me.ClientSize.Height - GetBordersHeight()
    End Function


    '[Margin and paddings]
    Public Function GetPaddingsHeight() As Single
        Return pCurrentProperties(StylePropertyEnum.StyleProperty_PaddingBottom) + pCurrentProperties(StylePropertyEnum.StyleProperty_PaddingTop)
    End Function

    Public Function GetPaddingsWidth() As Single
        Return pCurrentProperties(StylePropertyEnum.StyleProperty_PaddingLeft) + pCurrentProperties(StylePropertyEnum.StyleProperty_PaddingRight)
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

    Public Function GetBodyWidth() As Single
        Return Me.ClientSize.Width - GetPaddingsWidth()
    End Function

    Public Function GetBodyHeight() As Single
        Return Me.ClientSize.Height - GetPaddingsHeight()
    End Function

#End Region


#Region "Rendering"

    Public Function IsRendered() As Boolean Implements IControl.IsRendered
        Return pIsRendered
    End Function

    'Public Sub AddToResizeQueue(prop As StylePropertyEnum) Implements IControl.AddToResizeQueue

    'End Sub

    Public Sub Render()
        pCurrentProperties = calculateProperties()
        Call updateSizeValueTypes()
        Call updatePositionValueTypes()
        Call RQ.AddResizeTask(Me, StylePropertyEnum.StyleProperty_Width)
        Call RQ.AddResizeTask(Me, StylePropertyEnum.StyleProperty_Height)
        Call RQ.AddLayoutTask(Me)
    End Sub

    Public Sub UpdateSize()
        If pHeightValueType = ControlSizeTypeEnum.ControlSizeType_Auto Then
            Call UpdateWidth()
            Call UpdateHeight()
        Else
            Call UpdateHeight()
            Call UpdateWidth()
        End If
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

            If Not RQ.HasKey(widthUpdateKey) Then
                Call arrangeControls()
                If Not heightAlreadySet Then
                    Me.Height = CalculateControlHeight(Me, pCurrentProperties)
                End If
            End If

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
            If pHeightValueType = ControlSizeTypeEnum.ControlSizeType_Const Then
                Me.Width = CalculateControlHeight(Me, pCurrentProperties)
                widthAlreadySet = True
                For Each ctrl In pControls
                    Call ctrl.UpdateWidth()
                Next
            ElseIf pHeightValueType = ControlSizeTypeEnum.ControlSizeType_Auto Then
                For Each ctrl In pControls
                    Call ctrl.UpdateHeight()
                Next
            ElseIf pHeightValueType = ControlSizeTypeEnum.ControlSizeType_ParentRelative Then
                'not applies for Window (it has no parent)
            End If

            If Not RQ.HasKey(heightUpdateKey) Then
                Call arrangeControls()
                If Not widthAlreadySet Then
                    Me.Width = CalculateControlHeight(Me, pCurrentProperties)
                End If
            End If

        End If

    End Sub

    Private Sub arrangeControls()
        Dim key As String : key = Me.GetHashCode & "|Layout"
        '---------------------------------------------------------------------------------------------------------
        If RQ.HasKey(key) Then

            '[Actual arranging layout]



            Call RQ.removeKey(key)
        End If
    End Sub


    'Public Sub UpdateView(Optional propagateDown As Boolean = False) Implements IControl.UpdateView, IContainer.UpdateView
    '    Dim newProperties As Object() = calculateProperties()
    '    '---------------------------------------------------------------------------------------------------------
    '    Dim bordersChanged As Boolean = False              'borders
    '    Dim insidePropertiesChanged As Boolean = False     'i.e. background, font color - properties that doesn't affect any other controls.
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

    '    pCurrentProperties = newProperties

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
    '    Call compareSingleProperty(StylePropertyEnum.StyleProperty_Margin, newProperties(StylePropertyEnum.StyleProperty_Margin), positionChanged)
    '    Call compareSingleProperty(StylePropertyEnum.StyleProperty_MarginTop, newProperties(StylePropertyEnum.StyleProperty_MarginTop), positionChanged)
    '    Call compareSingleProperty(StylePropertyEnum.StyleProperty_MarginLeft, newProperties(StylePropertyEnum.StyleProperty_MarginLeft), positionChanged)
    '    Call compareSingleProperty(StylePropertyEnum.StyleProperty_MarginBottom, newProperties(StylePropertyEnum.StyleProperty_MarginBottom), positionChanged)
    '    Call compareSingleProperty(StylePropertyEnum.StyleProperty_MarginRight, newProperties(StylePropertyEnum.StyleProperty_MarginRight), positionChanged)

    '    If sizeChanged Then Call updateSize(propagateDown)
    '    If positionChanged Then Call updatePosition(propagateDown)
    '    'If insideLayoutChanged Then Call RearrangeControls()

    '    'Call updateSize(anyChanges)
    '    'Call updatePosition()
    '    'Call updateTitleBarVisibility()
    '    'If anyChanges Then Call RearrangeControls()

    '    pCurrentProperties = newProperties

    'End Sub

    'Private Sub updateSize(Optional propagateDown As Boolean = False)
    '    Dim heightChanged As Boolean
    '    Dim widthChanged As Boolean
    '    '---------------------------------------------------------------------------------------------------------

    '    'If Not pIsRendered Then
    '    '    updateActive = True
    '    'ElseIf pVisible Then
    '    '    Call ShowWindow(pHandle, pWindowSizeMode)
    '    '    updateActive = (pWindowSizeMode = ShowWindowCommands.Normal)
    '    'End If

    '    Call updateSizeValueTypes(heightChanged, widthChanged)
    '    Call updateHeight(heightChanged, propagateDown)



    'End Sub

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

    'Private Sub updateHeight(ByRef heightChanged As Boolean, Optional propagateDown As Boolean = False)
    '    Dim height As Single
    '    Dim ctrl As IControl
    '    Dim updateChildren As Boolean
    '    Dim isAutoHeight As Boolean : isAutoHeight = (pHeightValueType = ControlSizeTypeEnum.ControlSizeType_Auto)
    '    '---------------------------------------------------------------------------------------------------------

    '    If Not heightChanged Then
    '        height = CalculateControlHeight(Me, pCurrentProperties)
    '        heightChanged = (height <> Me.Height)
    '    End If

    '    If heightChanged Then
    '        'Call RQ.Add(Me, False, StylePropertyEnum.StyleProperty_Height)
    '        Call pPropertiesToUpdate.Add(StylePropertyEnum.StyleProperty_Height)

    '        If isTopDependingOnHeight() Then
    '            'Call RQ.Add(Me, False, StylePropertyEnum.StyleProperty_Height)
    '            Call pPropertiesToUpdate.Add(StylePropertyEnum.StyleProperty_Top)
    '        End If

    '        updateChildren = (isAutoHeight = False) Or propagateDown

    '    Else

    '        updateChildren = propagateDown

    '    End If

    '    If updateChildren Then
    '        For Each ctrl In pControls
    '            'Call RQ.Add(ctrl, isAutoHeight, StylePropertyEnum.StyleProperty_Height)
    '        Next
    '    End If

    'End Sub

    'Private Function isTopDependingOnHeight() As Boolean
    '    Return False
    'End Function



    Private Sub updatePosition(Optional propagateDown As Boolean = False)
        Dim topChanged As Boolean
        Dim leftChanged As Boolean
        '---------------------------------------------------------------------------------------------------------

        Call updatePositionValueTypes(topChanged, leftChanged)
        Call Err.Raise(1, Description:="Set pWindowPosition based on CSS...")

        If pWindowPosition = FormStartPosition.Manual Then
            Call locateWindowByCoordinates()
        ElseIf pWindowPosition = FormStartPosition.CenterScreen Then
            Call locateWindowInTheMiddleOfTheScreen()
        ElseIf pWindowPosition = FormStartPosition.CenterParent Then
            Call locateWindowInTheMiddleOfTheScreen()
        End If

    End Sub

    Private Sub updatePositionValueTypes(Optional ByRef topChanged As Boolean = False, Optional ByRef leftChanged As Boolean = False)
        Dim top As ControlSizeTypeEnum
        Dim left As ControlSizeTypeEnum
        '---------------------------------------------------------------------------------------------------------

        top = GetControlSizeType(pCurrentProperties(StylePropertyEnum.StyleProperty_Top))
        left = GetControlSizeType(pCurrentProperties(StylePropertyEnum.StyleProperty_Left))

        If top <> pTopValueType Then
            pTopValueType = top
            topChanged = True
        End If

        If left <> pLeftValueType Then
            pLeftValueType = left
            leftChanged = True
        End If

    End Sub






    Private Sub locateWindowByCoordinates()
        If Single.IsNaN(pCurrentProperties(StylePropertyEnum.StyleProperty_Top)) Then
            Me.Top = Screen.FromHandle(pHandle).WorkingArea.Height - pCurrentProperties(StylePropertyEnum.StyleProperty_Bottom) - pCurrentProperties(StylePropertyEnum.StyleProperty_Height)
        Else
            Me.Top = pCurrentProperties(StylePropertyEnum.StyleProperty_Top)
        End If
        If Single.IsNaN(pCurrentProperties(StylePropertyEnum.StyleProperty_Left)) Then
            Me.Left = Screen.FromHandle(pHandle).WorkingArea.Width - pCurrentProperties(StylePropertyEnum.StyleProperty_Right) - pCurrentProperties(StylePropertyEnum.StyleProperty_Width)
        Else
            Me.Left = pCurrentProperties(StylePropertyEnum.StyleProperty_Left)
        End If
    End Sub

    Private Sub locateWindowInTheMiddleOfTheScreen()
        Dim scr As System.Windows.Forms.Screen
        '---------------------------------------------------------------------------------------------------------
        scr = Screen.FromHandle(pHandle)
        Me.Top = scr.WorkingArea.Y + (scr.WorkingArea.Height - Me.Height) / 2
        Me.Left = scr.WorkingArea.X + (scr.WorkingArea.Width - Me.Width) / 2
    End Sub

    Private Sub updateTitleBarVisibility()
        If True Then 'pHasTitleBar Then
            Me.ControlBox = True
            Me.Text = pTitle
        Else
            Me.ControlBox = False
            Me.Text = String.Empty
        End If
    End Sub

    Private Sub updateInsideProperties()
        Me.BackColor = pCurrentProperties(StylePropertyEnum.StyleProperty_BackgroundColor)
    End Sub

    Private Sub updateBorders()
        If pHasBorder Then
            If pIsSizable Then
                Me.FormBorderStyle = Windows.Forms.FormBorderStyle.Sizable
            Else
                Me.FormBorderStyle = FormBorderStyle.FixedSingle
            End If
        Else
            Me.FormBorderStyle = FormBorderStyle.None
        End If
    End Sub




    'Public Sub RearrangeControls() 'Implements IContainer.RearrangeControls
    '    Stop
    '    'Err.Raise(1)
    'End Sub

    'Public Sub ResizeControls() 'Implements IContainer.ResizeControls
    '    '    Err.Raise(1)
    'End Sub


#End Region


#Region "Events"

    Private Sub Window_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        On Error Resume Next
        Call pListener.RaiseEvent_Load()
        On Error GoTo 0
    End Sub

    Private Sub Window_Disposed(sender As Object, e As EventArgs) Handles Me.Disposed
        On Error Resume Next
        Call pListener.RaiseEvent_Disposed()
        On Error GoTo 0
    End Sub

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

    Private Sub Window_GotFocus(sender As Object, e As EventArgs) Handles Me.GotFocus
        On Error Resume Next
        Call pListener.RaiseEvent_GotFocus()
        On Error GoTo 0
    End Sub

    Private Sub Window_LostFocus(sender As Object, e As EventArgs) Handles Me.LostFocus
        On Error Resume Next
        Call pListener.RaiseEvent_LostFocus()
        On Error GoTo 0
    End Sub

    Private Sub Window_ResizeEnd(sender As Object, e As EventArgs) Handles Me.ResizeEnd
        On Error Resume Next
        'Call rearrangeControls()
        Call pListener.RaiseEvent_SizeChanged(Me.Width, Me.Height, Me.Left, Me.Top)
        On Error GoTo 0
    End Sub

    Private Sub Window_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        'Call rearrangeControls()
    End Sub

    Private Sub Window_Scroll(sender As Object, e As ScrollEventArgs) Handles Me.Scroll
        On Error Resume Next
        Call pListener.RaiseEvent_Scroll()
        On Error GoTo 0
    End Sub

#End Region

#Region "Managing children"

    Public Sub AddControl(ctrl As IControl)
        Call ctrl.SetParent(Me)
        Call pControls.Add(ctrl)
        Call pControlsDict.Add(ctrl.GetId, ctrl)
        Call Me.Controls.Add(ctrl)
    End Sub

    Public Sub RemoveControl(ctrl As Control)
        Err.Raise(1)
    End Sub

#End Region

#Region "Styles - Global"
    Public Sub SetStylesManager(value As StylesManager)
        pStylesManager = value
    End Sub
    Public Function GetStylesManager() As StylesManager Implements IContainer.GetStylesManager, IControl.GetStylesManager
        Return pStylesManager
    End Function
#End Region

#Region "Styles - Local"

    Public Sub SetStyleClasses(classes As String, Optional invalidate As Boolean = True) Implements IControl.SetStyleClasses
        If pStylesMatrix Is Nothing Then pStylesMatrix = New StylesMatrix(Me)
        With pStylesMatrix
            Call .RemoveAllStyleClasses(False)
            Call .AddStyleClasses(classes, True)
        End With
        If invalidate Then Render()
    End Sub

    Public Sub AddStyleClasses(name As String, Optional invalidate As Boolean = True) Implements IControl.AddStyleClasses
        If pStylesMatrix Is Nothing Then pStylesMatrix = New StylesMatrix(Me)
        Call pStylesMatrix.AddStyleClasses(name, True)
        If invalidate Then Render()
    End Sub

    Public Sub RemoveStyleClasses(name As String, Optional invalidate As Boolean = True) Implements IControl.RemoveStyleClasses
        If pStylesMatrix Is Nothing Then pStylesMatrix = New StylesMatrix(Me)
        Call pStylesMatrix.RemoveStyleClasses(name, True)
        If invalidate Then Render()
    End Sub

    Public Sub SetStyleProperty(propertyType As Long, value As Object, Optional invalidate As Boolean = True) Implements IControl.SetStyleProperty
        If pStylesMatrix Is Nothing Then pStylesMatrix = New StylesMatrix(Me)
        Call pStylesMatrix.AddInlineStyle(propertyType, value, True)
        If invalidate Then Render()
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


#Region "Images"

    Public Sub AddImage(key As String, filepath As String)
        Dim image As Image
        image = Bitmap.FromFile(filepath)
        'Call pImageList.Images.Add(key, image)
    End Sub

    Public Function GetImageList() As ImageList
        'Return pImageList
    End Function

#End Region


End Class