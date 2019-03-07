Public Class TransparentPanel
    Inherits Panel

    Private Const WS_EX_TRANSPARENT As Long = &H20
    Private pOpacity As Integer = 50
    Private pGraphics As Graphics
    Private pText As String

    Public Sub TransparentPanel()
        SetStyle(ControlStyles.Opaque, True)
    End Sub

    Public Property Opacity() As Integer
        Get
            Return pOpacity
        End Get
        Set(value As Integer)
            If (value < 0 Or value > 100) Then
                Throw New ArgumentException("value must be between 0 and 100")
            End If
            pOpacity = value
        End Set
    End Property

    Public Sub setText(value As String)
        pText = value
    End Sub

    Protected Overrides ReadOnly Property CreateParams As CreateParams
        Get
            Dim cp As CreateParams
            cp = MyBase.CreateParams
            cp.ExStyle = cp.ExStyle Or WS_EX_TRANSPARENT
            Return cp
        End Get
    End Property

    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        pGraphics = e.Graphics
        With pGraphics
            .TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAlias
            .InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBilinear
            .PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.HighQuality
            .SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality
        End With

        Dim brush As SolidBrush
        brush = New SolidBrush(Color.FromArgb(pOpacity * 255 / 100, Me.BackColor))
        e.Graphics.FillRectangle(brush, Me.ClientRectangle)
        MyBase.OnPaint(e)

        If Len(pText) > 0 Then
            drawText(pText)
        End If

    End Sub


    Public Sub drawText(text As String)
        Dim font As Font = New Font("Century Gothic", 13, FontStyle.Bold)
        Dim textSize As SizeF = pGraphics.MeasureString(text, font)
        Dim width As Integer = CInt(Math.Ceiling(textSize.Width))
        Dim height As Integer = CInt(Math.Ceiling(textSize.Height))
        Dim size As Size = New Size(width, height)
        Dim position As Point = New Point(100, 100)
        Dim rectangle As Rectangle = New Rectangle(position, size)
        Dim color As Brush = Brushes.Black
        Call pGraphics.DrawString(text, font, color, rectangle)
    End Sub

End Class
