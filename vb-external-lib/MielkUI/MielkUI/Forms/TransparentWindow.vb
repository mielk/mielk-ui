Imports System.Windows.Forms
Imports System.Drawing
Imports System.Runtime.InteropServices

Friend Class TransparentWindow
    Inherits Form

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

    <Flags()> _
    Private Enum RedrawWindowFlags As UInteger
        ''' <summary>
        ''' Invalidates the rectangle or region that you specify in lprcUpdate or hrgnUpdate.
        ''' You can set only one of these parameters to a non-NULL value. If both are NULL, RDW_INVALIDATE invalidates the entire window.
        ''' </summary>
        Invalidate = &H1

        ''' <summary>Causes the OS to post a WM_PAINT message to the window regardless of whether a portion of the window is invalid.</summary>
        InternalPaint = &H2

        ''' <summary>
        ''' Causes the window to receive a WM_ERASEBKGND message when the window is repainted.
        ''' Specify this value in combination with the RDW_INVALIDATE value; otherwise, RDW_ERASE has no effect.
        ''' </summary>
        [Erase] = &H4

        ''' <summary>
        ''' Validates the rectangle or region that you specify in lprcUpdate or hrgnUpdate.
        ''' You can set only one of these parameters to a non-NULL value. If both are NULL, RDW_VALIDATE validates the entire window.
        ''' This value does not affect internal WM_PAINT messages.
        ''' </summary>
        Validate = &H8

        NoInternalPaint = &H10

        ''' <summary>Suppresses any pending WM_ERASEBKGND messages.</summary>
        NoErase = &H20

        ''' <summary>Excludes child windows, if any, from the repainting operation.</summary>
        NoChildren = &H40

        ''' <summary>Includes child windows, if any, in the repainting operation.</summary>
        AllChildren = &H80

        ''' <summary>Causes the affected windows, which you specify by setting the RDW_ALLCHILDREN and RDW_NOCHILDREN values, to receive WM_ERASEBKGND and WM_PAINT messages before the RedrawWindow returns, if necessary.</summary>
        UpdateNow = &H100

        ''' <summary>
        ''' Causes the affected windows, which you specify by setting the RDW_ALLCHILDREN and RDW_NOCHILDREN values, to receive WM_ERASEBKGND messages before RedrawWindow returns, if necessary.
        ''' The affected windows receive WM_PAINT messages at the ordinary time.
        ''' </summary>
        EraseNow = &H200

        Frame = &H400

        NoFrame = &H800
    End Enum

    Private Const WS_EX_LAYERED As Long = &H80000
    Private Const WS_EX_TRANSPARENT As Long = &H20
    Private Const LWA_ALPHA = 2
    Private Const LWA_COLORKEY = 1


    'Windows API functions
    Private Declare Auto Function SetLayeredWindowAttributes Lib "User32.Dll" (ByVal hWnd As IntPtr, ByVal crKey As Integer, ByVal Alpha As Byte, ByVal dwFlags As Integer) As Boolean

    <System.Runtime.InteropServices.DllImport("user32.dll", EntryPoint:="SetWindowLong")> _
    Private Shared Function SetWindowLong32(ByVal hWnd As IntPtr, <MarshalAs(UnmanagedType.I4)> nIndex As WindowLongFlags, ByVal dwNewLong As Integer) As Integer
    End Function
    <System.Runtime.InteropServices.DllImport("user32.dll", EntryPoint:="SetWindowLongPtr")> _
    Private Shared Function SetWindowLongPtr64(ByVal hWnd As IntPtr, <MarshalAs(UnmanagedType.I4)> nIndex As WindowLongFlags, ByVal dwNewLong As IntPtr) As IntPtr
    End Function
    Private Shared Function SetWindowLongPtr(ByVal hWnd As IntPtr, nIndex As WindowLongFlags, ByVal dwNewLong As IntPtr) As IntPtr
        If IntPtr.Size = 8 Then
            Return SetWindowLongPtr64(hWnd, nIndex, dwNewLong)
        Else
            Return New IntPtr(SetWindowLong32(hWnd, nIndex, dwNewLong.ToInt32))
        End If
    End Function

    <DllImport("user32.dll", EntryPoint:="GetWindowLong")> _
    Private Shared Function GetWindowLongPtr32(ByVal hWnd As IntPtr, ByVal nIndex As Integer) As IntPtr
    End Function
    <DllImport("user32.dll", EntryPoint:="GetWindowLongPtr")> _
    Private Shared Function GetWindowLongPtr64(ByVal hWnd As IntPtr, ByVal nIndex As Integer) As IntPtr
    End Function
    Private Shared Function GetWindowLongPtr(ByVal hWnd As IntPtr, ByVal nIndex As Integer) As IntPtr
        If IntPtr.Size = 8 Then
            Return GetWindowLongPtr64(hWnd, nIndex)
        Else
            Return GetWindowLongPtr32(hWnd, nIndex)
        End If
    End Function

    <DllImport("user32.dll")> _
    Private Shared Function RedrawWindow(hWnd As IntPtr, <[In]> ByRef lprcUpdate As RECT, hrgnUpdate As IntPtr, flags As RedrawWindowFlags) As Boolean
    End Function
    <DllImport("user32.dll")> _
    Private Shared Function RedrawWindow(hWnd As IntPtr, lprcUpdate As IntPtr, hrgnUpdate As IntPtr, flags As RedrawWindowFlags) As Boolean
    End Function

    <StructLayout(LayoutKind.Sequential)> _
    Private Structure RECT
        Private _Left As Integer, _Top As Integer, _Right As Integer, _Bottom As Integer

        Public Sub New(ByVal Rectangle As Rectangle)
            Me.New(Rectangle.Left, Rectangle.Top, Rectangle.Right, Rectangle.Bottom)
        End Sub
        Public Sub New(ByVal Left As Integer, ByVal Top As Integer, ByVal Right As Integer, ByVal Bottom As Integer)
            _Left = Left
            _Top = Top
            _Right = Right
            _Bottom = Bottom
        End Sub

        Public Property X As Integer
            Get
                Return _Left
            End Get
            Set(ByVal value As Integer)
                _Right = _Right - _Left + value
                _Left = value
            End Set
        End Property
        Public Property Y As Integer
            Get
                Return _Top
            End Get
            Set(ByVal value As Integer)
                _Bottom = _Bottom - _Top + value
                _Top = value
            End Set
        End Property
        Public Property Left As Integer
            Get
                Return _Left
            End Get
            Set(ByVal value As Integer)
                _Left = value
            End Set
        End Property
        Public Property Top As Integer
            Get
                Return _Top
            End Get
            Set(ByVal value As Integer)
                _Top = value
            End Set
        End Property
        Public Property Right As Integer
            Get
                Return _Right
            End Get
            Set(ByVal value As Integer)
                _Right = value
            End Set
        End Property
        Public Property Bottom As Integer
            Get
                Return _Bottom
            End Get
            Set(ByVal value As Integer)
                _Bottom = value
            End Set
        End Property
        Public Property Height() As Integer
            Get
                Return _Bottom - _Top
            End Get
            Set(ByVal value As Integer)
                _Bottom = value + _Top
            End Set
        End Property
        Public Property Width() As Integer
            Get
                Return _Right - _Left
            End Get
            Set(ByVal value As Integer)
                _Right = value + _Left
            End Set
        End Property
        Public Property Location() As Point
            Get
                Return New Point(Left, Top)
            End Get
            Set(ByVal value As Point)
                _Right = _Right - _Left + value.X
                _Bottom = _Bottom - _Top + value.Y
                _Left = value.X
                _Top = value.Y
            End Set
        End Property
        Public Property Size() As Size
            Get
                Return New Size(Width, Height)
            End Get
            Set(ByVal value As Size)
                _Right = value.Width + _Left
                _Bottom = value.Height + _Top
            End Set
        End Property

        Public Shared Widening Operator CType(ByVal Rectangle As RECT) As Rectangle
            Return New Rectangle(Rectangle.Left, Rectangle.Top, Rectangle.Width, Rectangle.Height)
        End Operator
        Public Shared Widening Operator CType(ByVal Rectangle As Rectangle) As RECT
            Return New RECT(Rectangle.Left, Rectangle.Top, Rectangle.Right, Rectangle.Bottom)
        End Operator
        Public Shared Operator =(ByVal Rectangle1 As RECT, ByVal Rectangle2 As RECT) As Boolean
            Return Rectangle1.Equals(Rectangle2)
        End Operator
        Public Shared Operator <>(ByVal Rectangle1 As RECT, ByVal Rectangle2 As RECT) As Boolean
            Return Not Rectangle1.Equals(Rectangle2)
        End Operator

        Public Overrides Function ToString() As String
            Return "{Left: " & _Left & "; " & "Top: " & _Top & "; Right: " & _Right & "; Bottom: " & _Bottom & "}"
        End Function

        Public Overloads Function Equals(ByVal Rectangle As RECT) As Boolean
            Return Rectangle.Left = _Left AndAlso Rectangle.Top = _Top AndAlso Rectangle.Right = _Right AndAlso Rectangle.Bottom = _Bottom
        End Function
        Public Overloads Overrides Function Equals(ByVal [Object] As Object) As Boolean
            If TypeOf [Object] Is RECT Then
                Return Equals(DirectCast([Object], RECT))
            ElseIf TypeOf [Object] Is Rectangle Then
                Return Equals(New RECT(DirectCast([Object], Rectangle)))
            End If

            Return False
        End Function
    End Structure

#End Region

#Region "Private variables"

    Private pHandle As IntPtr
    Private pIsLayered As Boolean
    Private pSeekSpeedInterval As Integer = 300
    Private pCurrentOpacity As Byte = 255
    Private pDestinationOpacity As Byte = 255

    '[Fading properties]
    Private pStartVal As Integer = 0
    Private pIncrementVal As Integer = 0
    Private pTotalSteps As Integer = 10
    Private pUp As Boolean = True

    '[Timers]
    Private WithEvents pFadingTimer As Timer

#End Region

#Region "Constructor"

    Public Sub New()
        pHandle = Me.Handle
        Call initializeTimer()
    End Sub

    Private Sub initializeTimer()
        pFadingTimer = New Timer()
        pFadingTimer.Interval = pSeekSpeedInterval
    End Sub

#End Region

#Region "Events"
    Friend Event OnFadeComplete()
#End Region

#Region "Event handlers"

    Private Sub pTimer_Tick(sender As Object, e As EventArgs) Handles pFadingTimer.Tick

        If pUp Then
            If pStartVal + pIncrementVal < pDestinationOpacity Then
                pStartVal += pIncrementVal
            Else
                pStartVal = pDestinationOpacity
            End If
            Call updateOpacity(CByte(pStartVal), True)
        Else
            If pStartVal + pIncrementVal > pDestinationOpacity Then
                pStartVal += pIncrementVal
            Else
                pStartVal = pDestinationOpacity
            End If
            Call updateOpacity(CByte(pStartVal), True)
        End If

        If (pCurrentOpacity = pDestinationOpacity) Then
            pFadingTimer.Stop()
            RaiseEvent OnFadeComplete()
        End If

    End Sub

#End Region

#Region "Public methods"

    Public Sub RefreshView()
        Call RedrawWindow(pHandle, IntPtr.Zero, IntPtr.Zero, RedrawWindowFlags.Frame OrElse RedrawWindowFlags.UpdateNow OrElse RedrawWindowFlags.Invalidate OrElse RedrawWindowFlags.AllChildren)
    End Sub

    Public Sub setTransparentLayeredWindow()
        Dim oldStyle As Long
        '------------------------------------------------------------------------------------------------------------------------------------------------------
        If Not pIsLayered Then
            oldStyle = GetWindowLongPtr(Me.Handle, GWL.GWL_EXSTYLE)
            Call SetWindowLongPtr(pHandle, GWL.GWL_EXSTYLE, oldStyle Or WS_EX_LAYERED)
            Call SetLayeredWindowAttributes(pHandle, 0, pCurrentOpacity, LWA_ALPHA)
            pIsLayered = True
        End If
    End Sub

    Public Sub clearTransparentLayeredWindow()
        Dim oldStyle As Long
        '------------------------------------------------------------------------------------------------------------------------------------------------------
        If pIsLayered Then
            oldStyle = GetWindowLongPtr(Me.Handle, GWL.GWL_EXSTYLE)
            Call SetWindowLongPtr(Me.Handle, GWL.GWL_EXSTYLE, oldStyle Or WS_EX_LAYERED)
            pIsLayered = False
        End If
    End Sub

    Public Sub updateOpacity(opacity As Byte, forceRefresh As Boolean)
        If pIsLayered Then
            Call SetLayeredWindowAttributes(pHandle, 0, opacity, LWA_ALPHA)
            pCurrentOpacity = opacity
            If (forceRefresh) Then RefreshView()
        End If
    End Sub

    Public Sub updateOpacityInPercent(opacityAsPercent As Integer, forceRefresh As Boolean)
        Dim opacity As Byte
        '------------------------------------------------------------------------------------------------------------------------------------------------------
        If pIsLayered Then
            opacity = CByte((255 * Math.Max(0, Math.Min(opacityAsPercent, 100))) / 100)
            Call SetLayeredWindowAttributes(pHandle, 0, opacity, LWA_ALPHA)
            pCurrentOpacity = opacity
            If (forceRefresh) Then RefreshView()
        End If
    End Sub

    Public Sub seekToDestination()

        If pIsLayered Then
            If pCurrentOpacity < pDestinationOpacity Then
                pUp = True
            ElseIf pCurrentOpacity > pDestinationOpacity Then
                pUp = False
            Else
                Return
            End If
        End If

        pIncrementVal = (CInt(pDestinationOpacity) - CInt(pCurrentOpacity)) / pTotalSteps
        pFadingTimer.Interval = pSeekSpeedInterval
        pFadingTimer.Start()

    End Sub

    Public Sub seekTo(destinationOpacity As Byte)
        If pIsLayered Then
            pDestinationOpacity = destinationOpacity

            If (pCurrentOpacity < pDestinationOpacity) Then
                pUp = True
            ElseIf (pCurrentOpacity > pDestinationOpacity) Then
                pUp = False
            Else
                Return
            End If

        End If

        pStartVal = pCurrentOpacity
        pIncrementVal = (CInt(pDestinationOpacity) - CInt(pCurrentOpacity)) / pTotalSteps
        pFadingTimer.Interval = pSeekSpeedInterval
        pFadingTimer.Start()

    End Sub


    Public Sub seekToPercentage(destinationOpacityAsPercent As Integer)
        If pIsLayered Then
            DestinationOpacityPercent = destinationOpacityAsPercent

            If (pCurrentOpacity < pDestinationOpacity) Then
                pUp = True
            ElseIf (pCurrentOpacity > pDestinationOpacity) Then
                pUp = False
            Else
                Return
            End If

        End If

        pStartVal = pCurrentOpacity
        pIncrementVal = (CInt(pDestinationOpacity) - CInt(pCurrentOpacity)) / pTotalSteps
        pFadingTimer.Interval = pSeekSpeedInterval
        pFadingTimer.Start()

    End Sub

    Public Sub stopTimer()
        Call pFadingTimer.Stop()
        Call updateOpacity(255, True)
    End Sub

#End Region

#Region "Public properties"

    Public Property IsLayered() As Boolean
        Get
            Return pIsLayered
        End Get
        Set(value As Boolean)
            pIsLayered = value
        End Set
    End Property

    Public Property SeekSpeed() As Integer
        Get
            Return pSeekSpeedInterval
        End Get
        Set(value As Integer)
            pSeekSpeedInterval = value
            pFadingTimer.Interval = value
        End Set
    End Property

    Public Property DestinationOpacity() As Byte
        Get
            Return pDestinationOpacity
        End Get
        Set(value As Byte)
            pDestinationOpacity = value
        End Set
    End Property

    Public Property DestinationOpacityPercent() As Integer
        Get
            Return (pDestinationOpacity * 100) / 255
        End Get
        Set(value As Integer)
            pDestinationOpacity = CByte((255 * Math.Max(0, Math.Min(value, 100))) / 100)
        End Set
    End Property

    Public ReadOnly Property CurrentOpacity() As Byte
        Get
            Return pCurrentOpacity
        End Get
    End Property

    Public Property StepsToFade() As Integer
        Get
            Return pTotalSteps
        End Get
        Set(value As Integer)
            pTotalSteps = value
        End Set
    End Property

#End Region


    Private Sub InitializeComponent()
        Me.SuspendLayout()
        '
        'TransparentWindow
        '
        Me.ClientSize = New System.Drawing.Size(284, 261)
        Me.Name = "TransparentWindow"
        Me.ResumeLayout(False)

    End Sub


    Private Sub TransparentWindow_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class
