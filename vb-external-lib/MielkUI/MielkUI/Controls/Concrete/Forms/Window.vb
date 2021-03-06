﻿Imports System.Windows.Forms
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
    Private pBody As Body

    '[State]
    Private pIsDisplayed As Boolean
    Private pIsRendered As Boolean

    '[Size and position]
    Private pStylesMatrix As StylesMatrix
    Private pCurrentProperties(0 To CountStylePropertiesEnums()) As Object

#End Region


#Region "Constructor"

    Public Sub New()
        Visible = pVisible
        pHandle = Me.Handle
        pStylesMatrix = New StylesMatrix(Me, False)
        'Call createBody()
    End Sub

#End Region



#Region "Window size & position"

    Public Sub Display(Optional ByVal isModal As Boolean = False)
        Call Redraw()
        pIsDisplayed = True
        pIsModal = isModal
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
        Call updateSizeAndPosition()
    End Sub

    Public Sub SetWindowSizeMode(value As ShowWindowCommands)
        pWindowSizeMode = value
        If pIsDisplayed Then
            Call ShowWindow(pHandle, value)
            Call redrawControls()
        End If
    End Sub

#End Region


#Region "Window properties - Getters"

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




#Region "UI properties - Setters"


    '[Position]
    Public Sub SetTop(value As Object, Optional redraw As Boolean = True)
        Call pStylesMatrix.AddInlineStyle(StylePropertyEnum.StyleProperty_Top, value)
        If redraw Then updatePosition()
    End Sub

    Public Sub SetLeft(value As Object, Optional redraw As Boolean = True)
        Call pStylesMatrix.AddInlineStyle(StylePropertyEnum.StyleProperty_Left, value)
        If redraw Then updatePosition()
    End Sub

    Public Sub SetBottom(value As Object, Optional redraw As Boolean = True)
        Call pStylesMatrix.AddInlineStyle(StylePropertyEnum.StyleProperty_Bottom, value)
        If redraw Then updatePosition()
        'pBottom = value
        'pTop = Single.NaN
        'Call updatePosition()
    End Sub

    Public Sub SetRight(value As Single)
        'pRight = value
        'pLeft = Single.NaN
        'Call updatePosition()
    End Sub


    '[Size]
    Public Sub SetWidth(value As Single)
        'pWidth = value
        'Call updateSizeAndPosition()
    End Sub

    Public Sub SetAutoWidth(value As Boolean)
        'pAutoWidth = value
        'If pAutoWidth Then
        '    pWidth = Single.NaN
        'Else
        '    pWidth = Me.Width
        'End If
        'Call updateSizeAndPosition()
    End Sub

    Public Sub SetHeight(value As Single)
        'pHeight = value
        'pAutoHeight = False
        'Call updateSizeAndPosition()
    End Sub

    Public Sub SetAutoHeight(value As Boolean)
        'pAutoHeight = value
        'If pAutoHeight Then
        '    pHeight = Single.NaN
        'Else
        '    pHeight = Me.Height
        'End If
        'Call updateSizeAndPosition()
    End Sub





    '[Paddings]
    Public Sub SetPaddings(ByVal value As Single)
        'pPaddingTop = value
        'pPaddingRight = value
        'pPaddingBottom = value
        'pPaddingLeft = value
        'Call updateSizeAndPosition()
    End Sub

    Public Sub SetPaddingTop(value As Single)
        'pPaddingTop = value
        'Call updateSizeAndPosition()
    End Sub

    Public Sub SetPaddingRight(value As Single)
        'pPaddingRight = value
        'Call updateSizeAndPosition()
    End Sub

    Public Sub SetPaddingBottom(value As Single)
        'pPaddingBottom = value
        'Call updateSizeAndPosition()
    End Sub

    Public Sub SetPaddingLeft(value As Single)
        'pPaddingLeft = value
        'Call updateSizeAndPosition()
    End Sub



    '[Background]
    Public Sub SetBackgroundColorByNumeric(value As Long, Optional opacity As Single = 1)
        Dim blue As Long = value \ 65536
        Dim green As Long = (value - blue * 65536) \ 256
        Dim red As Long = value - blue * 65536 - green * 256
        Dim color As Color
        '------------------------------------------------------------------------------------------------------------------------------------------------------------
        color = System.Drawing.Color.FromArgb(Math.Max(Math.Min(opacity * 255, 255), 0), red, green, blue)
        'Me.BackColor = pBackgroundColor
    End Sub

    Public Sub SetBackgroundColorByRgb(color As String, Optional opacity As Single = 1)
        'Me.BackColor = pBackgroundColor
    End Sub


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

    Public Function GetWidth() As Single
        Return Me.Width
    End Function

    Public Function GetHeight() As Single
        Return Me.Height
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

    Public Function GetPaddingTop() As Single
        Return pCurrentProperties(StylePropertyEnum.StyleProperty_PaddingTop)
    End Function

    Public Function GetPaddingRight() As Single
        Return pCurrentProperties(StylePropertyEnum.StyleProperty_PaddingRight)
    End Function

    Public Function GetPaddingBottom() As Single
        Return pCurrentProperties(StylePropertyEnum.StyleProperty_PaddingBottom)
    End Function

    Public Function GetPaddingLeft() As Single
        Return pCurrentProperties(StylePropertyEnum.StyleProperty_PaddingLeft)
    End Function

#End Region





#Region "Rendering"

    Public Sub Redraw()
        Call updateSize()
        Call updateBorders()
        Call updateBackground()
        Call updateTitleBarVisibility()
        Call updatePosition()
    End Sub

    Private Sub updateSizeAndPosition()
        updateSize()
        updatePosition()
    End Sub

    Private Sub updateSize()
        If pVisible Then
            Call ShowWindow(pHandle, pWindowSizeMode)
            If pWindowSizeMode = ShowWindowCommands.Normal Then
                Me.Height = pCurrentProperties(StylePropertyEnum.StyleProperty_Height) + IIf(pBorderBox, 0, Math.Max(GetPaddingTop, 0) + Math.Max(GetPaddingBottom, 0) + GetTitleBarHeigth())
                Me.Width = pCurrentProperties(StylePropertyEnum.StyleProperty_Width) + IIf(pBorderBox, 0, Math.Max(GetPaddingLeft, 0) + Math.Max(GetPaddingRight, 0))
            End If
        End If
    End Sub

    Private Sub updatePosition()
        If pWindowPosition = FormStartPosition.Manual Then
            Call locateWindowByCoordinates()
        ElseIf pWindowPosition = FormStartPosition.CenterScreen Then
            Call locateWindowInTheMiddleOfTheScreen()
        ElseIf pWindowPosition = FormStartPosition.CenterParent Then
            Call locateWindowInTheMiddleOfTheScreen()
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
        If pHasTitleBar Then
            Me.ControlBox = True
            Me.Text = pTitle
        Else
            Me.ControlBox = False
            Me.Text = String.Empty
        End If
    End Sub

    Private Sub updateBackground()
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

    Private Sub redrawControls()

    End Sub

#End Region


#Region "Events"

    Private Sub Window_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If Not pListener Is Nothing Then
            On Error Resume Next
            Call pListener.RaiseEvent_Load()
        End If
    End Sub

    Private Sub Window_Disposed(sender As Object, e As EventArgs) Handles Me.Disposed
        If Not pListener Is Nothing Then
            On Error Resume Next
            Call pListener.RaiseEvent_Disposed()
        End If
    End Sub

    Private Sub Window_MouseEnter(sender As Object, e As EventArgs) Handles Me.MouseEnter
        If Not pListener Is Nothing Then
            On Error Resume Next
            Call pListener.RaiseEvent_MouseEnter()
        End If
    End Sub

    Private Sub Window_MouseLeave(sender As Object, e As EventArgs) Handles Me.MouseLeave
        If Not pListener Is Nothing Then
            On Error Resume Next
            Call pListener.RaiseEvent_MouseLeave()
        End If
    End Sub

    Private Sub Window_GotFocus(sender As Object, e As EventArgs) Handles Me.GotFocus
        If Not pListener Is Nothing Then
            On Error Resume Next
            Call pListener.RaiseEvent_GotFocus()
        End If
    End Sub

    Private Sub Window_LostFocus(sender As Object, e As EventArgs) Handles Me.LostFocus
        If Not pListener Is Nothing Then
            On Error Resume Next
            Call pListener.RaiseEvent_LostFocus()
        End If
    End Sub

    Private Sub Window_ResizeEnd(sender As Object, e As EventArgs) Handles Me.ResizeEnd
        If Not pListener Is Nothing Then
            On Error Resume Next
            Call pListener.RaiseEvent_SizeChanged(Me.Width, Me.Height, Me.Left, Me.Top)
        End If
    End Sub

    Private Sub Window_Scroll(sender As Object, e As ScrollEventArgs) Handles Me.Scroll
        If Not pListener Is Nothing Then
            On Error Resume Next
            Call pListener.RaiseEvent_Scroll()
        End If
    End Sub

#End Region


    Private Sub InitializeComponent()
        Me.SuspendLayout()
        '
        'Window
        '
        Me.ClientSize = New System.Drawing.Size(284, 261)
        Me.Name = "Window"
        Me.ResumeLayout(False)

    End Sub


#Region "Managing children"

    'Public Sub CreateBody()
    '    pBody = New Body(Me)
    '    Me.Controls.Add(pBody)
    'End Sub

    'Public Sub AddControl(ctrl As Control)
    '    Call pBody.AddControl(ctrl)
    'End Sub

    'Public Sub RemoveControl(ctrl As Control)
    '    Call pBody.RemoveControl(ctrl)
    'End Sub

#End Region


#Region "Styles - Global"
    Public Sub SetStylesManager(value As StylesManager)
        pStylesManager = value
        Call pStylesMatrix.LoadStyles()
    End Sub
    Public Function GetStylesManager() As StylesManager Implements IContainer.GetStylesManager, IControl.GetStylesManager
        Return pStylesManager
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

#End Region



    Public Function GetLog() As String
        Return 1
    End Function



End Class