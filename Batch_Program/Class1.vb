Imports System.Drawing
Imports System.Windows.Forms
Imports System.ComponentModel
Imports System.Drawing.Drawing2D

'Namespace VistaStyleProgressBar

' A replacement for the default ProgressBar control.
<DefaultEvent("ValueChanged")> _
Public Class MSVistaProgressBar
    Inherits System.Windows.Forms.UserControl

#Region "- Enums -"
    Public Enum PBarOrientation
        LeftToRight = 0
        RightToLeft = 1
        'TopToBottom = 2
        'BottomToTop = 3
    End Enum
    Public Enum BarStyle
        Continuous = 0
        Marquee = 1
        Blocks = 2
    End Enum
    Public Enum BarColorsWhen
        BarProgress = 0
        OnThreshold = 1
        None = 2
    End Enum
    Public Enum BackGradiant
        None = 0
        Horizontal = 1
        Vertical = 2
        ForwardDiagonal = 3
        BackwardDiagonal = 4
    End Enum
#End Region

#Region "- Designer -"

    ' Required designer variable.
    Private components As System.ComponentModel.Container = Nothing

    ' Create the control and initialize it.
    Public Sub New()
        ' This call is required by the Windows.Forms Form Designer.
        InitializeComponent()
        Me.SetStyle(ControlStyles.AllPaintingInWmPaint, True)
        Me.SetStyle(ControlStyles.DoubleBuffer, True)
        Me.SetStyle(ControlStyles.ResizeRedraw, True)
        Me.SetStyle(ControlStyles.Selectable, True)
        Me.SetStyle(ControlStyles.SupportsTransparentBackColor, True)
        Me.SetStyle(ControlStyles.UserPaint, True)
        Me.BackColor = Color.Transparent

        'If Not InDesignMode() Then
        AddHandler mGlowAnimation.Tick, AddressOf mGlowAnimation_Tick
        AddHandler mMarqueeAnimation.Tick, AddressOf mMarqueeAnimation_Tick
        mMarqueeAnimation.Interval = 80
        mGlowAnimation.Interval = 15
        If Value < MaxValue Then
            mGlowAnimation.Start()
        End If
        'End If
    End Sub

    ' Clean up any resources being used.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If components IsNot Nothing Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

#Region "- Component Designer -"

    ' Required method for Designer support - do not modify
    ' the contents of this method with the code editor.
    Private Sub InitializeComponent()
        Me.Name = "ProgressBar"
        Me.Size = New System.Drawing.Size(264, 32)
        AddHandler Me.Paint, AddressOf ProgressBar_Paint
    End Sub
#End Region

#End Region

#Region "- Properties -"

    Private mGlowPosition As Integer = -325
    Private mGlowAnimation As New Timer()
    Private mMarqueeAnimation As New Timer()
    Private mMarqueePosition As Integer = -120
    Private ExcessPixels As Integer = 0

#Region "- Percentage Text -"

    Dim McShowText As Boolean = True
    <Description("Sets whether the Progress Text Is Shown"), Category("Display Text")> _
    Public Property ShowText() As Boolean
        Get
            Return McShowText
        End Get
        Set(ByVal value As Boolean)
            McShowText = value
            Me.Invalidate()
        End Set
    End Property

    Dim McDispText As String = "%P%"
    <Description("Sets The Text Displayed With the Percentage"), Category("Display Text")> _
    Public Property DisplayText() As String
        Get
            Return McDispText
        End Get
        Set(ByVal value As String)
            McDispText = value
            Me.Invalidate()
        End Set
    End Property

    Dim McDispTextFont As Font = New Font("Arial", 8, FontStyle.Regular, GraphicsUnit.Point)
    <Description("Sets The Font of The Text"), Category("Display Text")> _
    Public Property DisplayTextFont() As Font
        Get
            Return McDispTextFont
        End Get
        Set(ByVal value As Font)
            McDispTextFont = value
            Me.Invalidate()
        End Set
    End Property

    Dim McDispTextColor As Color = Color.FromKnownColor(KnownColor.ControlText)
    <Description("Sets The Color of The Text"), Category("Display Text")> _
    Public Property DisplayTextColor() As Color
        Get
            Return McDispTextColor
        End Get
        Set(ByVal value As Color)
            McDispTextColor = value
            Me.Invalidate()
        End Set
    End Property
#End Region

#Region "- Value -"

    Private mValue As Integer = 0
    ''' <summary>
    ''' The value that is displayed on the progress bar.
    ''' </summary>
    <Category("Value"), DefaultValue(0), Description("The value that is displayed on the progress bar.")> _
    Public Property Value() As Integer
        Get
            Return mValue
        End Get
        Set(ByVal value As Integer)
            If value > MaxValue OrElse value < MinValue Then
                Exit Property
            End If
            mValue = value
            If value < MaxValue And Not Me.GlowWhenFinished Then
                mGlowAnimation.Start()
            End If
            If value = MaxValue And Not Me.GlowWhenFinished Then
                mGlowAnimation.Stop()
                mGlowPosition = -320
            End If

            RaiseEvent ValueChanged(Me, New System.EventArgs)
            Me.Invalidate()
        End Set
    End Property

    Private mMaxValue As Integer = 100
    ''' <summary>
    ''' The maximum value for the Value property.
    ''' </summary>
    <Category("Value"), DefaultValue(100), Description("The maximum value for the Value property.")> _
    Public Property MaxValue() As Integer
        Get
            Return mMaxValue
        End Get
        Set(ByVal value As Integer)
            mMaxValue = value
            If value > MaxValue Then
                value = MaxValue
            End If
            If value < MaxValue And Not Me.GlowWhenFinished Then
                mGlowAnimation.Start()
            End If
            RaiseEvent MaxChanged(Me, New System.EventArgs)
            Me.Invalidate()
        End Set
    End Property

    Private mMinValue As Integer = 0
    ''' <summary>
    ''' The minimum value for the Value property.
    ''' </summary>
    <Category("Value"), DefaultValue(0), Description("The minimum value for the Value property.")> _
    Public Property MinValue() As Integer
        Get
            Return mMinValue
        End Get
        Set(ByVal value As Integer)
            mMinValue = value
            If value < MinValue Then
                value = MinValue
            End If
            RaiseEvent MinChanged(Me, New System.EventArgs())
            Me.Invalidate()
        End Set
    End Property

#End Region

#Region "- Threshholds -"
    Private mcThresholds As New List(Of ThresholdItem)
    ''' <summary>
    ''' The minimum value for the Value property.
    ''' </summary>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Content), _
Category("Thresholds")> _
    Public Property BarThresholds() As List(Of ThresholdItem)
        Get
            Return mcThresholds
        End Get
        Set(ByVal value As List(Of ThresholdItem))
            mcThresholds = value
            Me.Invalidate()
        End Set
    End Property
#End Region

#Region "- Bar -"

    Private mStartColor As Color = Color.FromArgb(210, 0, 0)
    ' The start color for the progress bar.
    ' 210, 000, 000 = Red
    ' 210, 202, 000 = Yellow
    ' 000, 163, 211 = Blue
    ' 000, 211, 040 = Green
    <Category("Bar"), DefaultValue(GetType(Color), "210, 0, 0"), Description("The start color for the progress bar." & "210, 000, 000 = Red" & vbLf & "210, 202, 000 = Yellow" & vbLf & "000, 163, 211 = Blue" & vbLf & "000, 211, 040 = Green" & vbLf)> _
    Public Property StartColor() As Color
        Get
            Return mStartColor
        End Get
        Set(ByVal value As Color)
            mStartColor = value
            Me.Invalidate()
        End Set
    End Property

    Private mEndColor As Color = Color.FromArgb(0, 211, 40)
    ' The end color for the progress bar.
    ' 210, 000, 000 = Red
    ' 210, 202, 000 = Yellow
    ' 000, 163, 211 = Blue
    ' 000, 211, 040 = Green
    <Category("Bar"), DefaultValue(GetType(Color), "0, 211, 40"), Description("The end color for the progress bar." & "210, 000, 000 = Red" & vbLf & "210, 202, 000 = Yellow" & vbLf & "000, 163, 211 = Blue" & vbLf & "000, 211, 040 = Green" & vbLf)> _
    Public Property EndColor() As Color
        Get
            Return mEndColor
        End Get
        Set(ByVal value As Color)
            mEndColor = value
            Me.Invalidate()
        End Set
    End Property

    Dim McBarOriant As PBarOrientation = 0
    <Description("The Progress Bar Direction."), Category("Bar"), DefaultValue(GetType(PBarOrientation), "LeftToRight")> _
    Public Property ProgressDirection() As PBarOrientation
        Get
            Return McBarOriant
        End Get
        Set(ByVal value As PBarOrientation)
            McBarOriant = value
            Me.Invalidate()
        End Set
    End Property

    Dim McBarStyle As BarStyle = 0
    <Description("The Style Of the Progress Bar."), Category("Bar"), DefaultValue(GetType(BarStyle), "Continuous")> _
    Public Property ProgressBarStyle() As BarStyle
        Get
            Return McBarStyle
        End Get
        Set(ByVal value As BarStyle)
            McBarStyle = value
            If value = BarStyle.Continuous OrElse value = BarStyle.Blocks Then
                mMarqueeAnimation.Enabled = False
                mGlowAnimation.Enabled = True
            Else
                mMarqueeAnimation.Enabled = True
                mGlowAnimation.Enabled = False
                mGlowPosition = -320
            End If
            Me.Invalidate()
        End Set
    End Property

    Dim McBarColorMethod As BarColorsWhen = 0
    <Description("Sets how the Bar's Color Changes."), Category("Bar"), DefaultValue(GetType(BarColorsWhen), "BarProgress")> _
    Public Property BarColorMethod() As BarColorsWhen
        Get
            Return McBarColorMethod
        End Get
        Set(ByVal value As BarColorsWhen)
            McBarColorMethod = value
            Me.Invalidate()
        End Set
    End Property

    Private mMarqueeSpeed As Integer = 80
    ''' <summary>
    ''' The color of the glow.
    ''' </summary>
    <Category("Bar"), DefaultValue(80), Description("The Speed of the Marquee. (Lower Value = Higher Speed)")> _
    Public Property MarqueeSpeed() As Integer
        Get
            Return mMarqueeSpeed
        End Get
        Set(ByVal value As Integer)
            mMarqueeSpeed = value
            mMarqueeAnimation.Interval = value
            Me.Invalidate()
        End Set
    End Property
#End Region

#Region "- Background -"

    Private mOuterStrokeColor As Color = Color.FromArgb(178, 178, 178)
    <Category("Background"), DefaultValue(GetType(Color), "178,178,178"), Description("The color of the Outr Stroke.")> _
    Public Property OuterStrokeColor() As Color
        Get
            Return mOuterStrokeColor
        End Get
        Set(ByVal value As Color)
            mOuterStrokeColor = value
            Me.Invalidate()
        End Set
    End Property

    Private mBackgroundColor As Color = Color.FromArgb(201, 201, 201)
    ''' <summary>
    ''' The color of the background.
    ''' </summary>
    <Category("Background"), DefaultValue(GetType(Color), "201,201,201"), Description("The color of the background.")> _
    Public Property BackgroundColor() As Color
        Get
            Return mBackgroundColor
        End Get
        Set(ByVal value As Color)
            mBackgroundColor = value
            Me.Invalidate()
        End Set
    End Property

    Private mBackgroundColor2 As Color = Color.FromArgb(201, 201, 201)
    ''' <summary>
    ''' The color of the background.
    ''' </summary>
    <Category("Background"), DefaultValue(GetType(Color), "201,201,201"), Description("The Second color of the background. Used Only if a Gradiant Is Applied.")> _
    Public Property BackgroundColor2() As Color
        Get
            Return mBackgroundColor2
        End Get
        Set(ByVal value As Color)
            mBackgroundColor2 = value
            Me.Invalidate()
        End Set
    End Property

    Private mGradStyle As BackGradiant = BackGradiant.None
    <Category("Background"), DefaultValue(GetType(Color), "201,201,201"), Description("The Style of the Background Gradiant.")> _
    Public Property GradiantStyle() As BackGradiant
        Get
            Return mGradStyle
        End Get
        Set(ByVal value As BackGradiant)
            mGradStyle = value
            Me.Invalidate()
        End Set
    End Property
#End Region

#Region "- Highlights and Glow -"

    Private mHighlightColor As Color = Color.White
    ''' <summary>
    ''' The color of the highlights.
    ''' </summary>
    <Category("Highlight and Glow"), DefaultValue(GetType(Color), "White"), Description("The color of the Bottom Highlight.")> _
    Public Property HighlightColor() As Color
        Get
            Return mHighlightColor
        End Get
        Set(ByVal value As Color)
            mHighlightColor = value
            Me.Invalidate()
        End Set
    End Property

    Private mTopHighlightColor As Color = Color.White
    ''' <summary>
    ''' The color of the highlights.
    ''' </summary>
    <Category("Highlight and Glow"), DefaultValue(GetType(Color), "White"), Description("The color of the Top Highlight.")> _
    Public Property TopHighlightColor() As Color
        Get
            Return mTopHighlightColor
        End Get
        Set(ByVal value As Color)
            mTopHighlightColor = value
            Me.Invalidate()
        End Set
    End Property

    Private mAnimate As Boolean = True
    ''' <summary>
    ''' Whether the glow is animated.
    ''' </summary>
    <Category("Highlight and Glow"), DefaultValue(GetType(Boolean), "true"), Description("Whether the glow is animated or not.")> _
    Public Property Animate() As Boolean
        Get
            Return mAnimate
        End Get
        Set(ByVal value As Boolean)
            mAnimate = value
            If value Then
                mGlowAnimation.Start()
            Else
                mGlowAnimation.Stop()
                mGlowPosition = -320
            End If
            Me.Invalidate()
        End Set
    End Property

    Private mGlowColor As Color = Color.FromArgb(150, 255, 255, 255)
    ''' <summary>
    ''' The color of the glow.
    ''' </summary>
    <Category("Highlight and Glow"), DefaultValue(GetType(Color), "150, 255, 255, 255"), Description("The color of the glow.")> _
    Public Property GlowColor() As Color
        Get
            Return mGlowColor
        End Get
        Set(ByVal value As Color)
            mGlowColor = value
            Me.Invalidate()
        End Set
    End Property

    Private mGlowSpeed As Integer = 15
    ''' <summary>
    ''' The color of the glow.
    ''' </summary>
    <Category("Highlight and Glow"), DefaultValue(15), Description("The Speed of the glow. (Lower Value = Higher Speed)")> _
    Public Property GlowTmrInterval() As Integer
        Get
            Return mGlowSpeed
        End Get
        Set(ByVal value As Integer)
            mGlowSpeed = value
            mGlowAnimation.Interval = value
            Me.Invalidate()
        End Set
    End Property

    Private McGlowOnComplete As Boolean = True
    ''' <summary>
    ''' Whether the glow is animated.
    ''' </summary>
    <Category("Highlight and Glow"), DefaultValue(GetType(Boolean), "true"), Description("Sets Whether the glow will Continue After The Bar Is Full.")> _
    Public Property GlowWhenFinished() As Boolean
        Get
            Return McGlowOnComplete
        End Get
        Set(ByVal value As Boolean)
            McGlowOnComplete = value
            If value = True And Me.Animate Then
                mGlowAnimation.Enabled = True
            End If
            If Not value And Me.Value = MaxValue Then
                mGlowAnimation.Enabled = False
                mGlowPosition = -320
            End If
            Me.Invalidate()
        End Set
    End Property
#End Region

#Region "- Block Style -"

    Dim mBlockSize As Integer = 10
    <Description("The Block Size for when Style is Blocks."), Category("Blocks")> _
    Public Property BlockSize() As Integer
        Get
            Return mBlockSize
        End Get
        Set(ByVal value As Integer)
            mBlockSize = value
            Me.Invalidate()
        End Set
    End Property

    Dim mBlockSpacing As Integer = 10
    <Description("The Block Spacing for when Style is Blocks."), Category("Blocks")> _
    Public Property BlockSpacing() As Integer
        Get
            Return mBlockSpacing
        End Get
        Set(ByVal value As Integer)
            mBlockSpacing = value
            Me.Invalidate()
        End Set
    End Property

#End Region

#End Region

#Region "- Drawing -"

    Private Sub DrawBackground(ByVal g As Graphics)
        Dim r As Rectangle = Me.ClientRectangle
        r.Width -= 1
        r.Height -= 1
        Dim rr As GraphicsPath = RoundRect(r, 2, 2, 2, 2)
        Dim BGBrush As Object = New SolidBrush(Me.BackgroundColor)

        If Not Me.GradiantStyle = BackGradiant.None Then
            Dim gDir As LinearGradientMode
            Select Case Me.GradiantStyle
                Case BackGradiant.Horizontal
                    gDir = LinearGradientMode.Horizontal
                Case BackGradiant.Vertical
                    gDir = LinearGradientMode.Vertical
                Case BackGradiant.BackwardDiagonal
                    gDir = LinearGradientMode.BackwardDiagonal
                Case Else
                    gDir = LinearGradientMode.ForwardDiagonal
            End Select
            BGBrush = New LinearGradientBrush(rr.GetBounds, Me.BackgroundColor, Me.BackgroundColor2, gDir)
        End If
        g.FillPath(BGBrush, rr)
    End Sub

    Private Sub DrawOuterStroke(ByVal g As Graphics)
        Dim r As Rectangle = Me.ClientRectangle
        r.Width -= 1
        r.Height -= 1
        Dim rr As GraphicsPath = RoundRect(r, 2, 2, 2, 2)
        g.DrawPath(New Pen(Me.OuterStrokeColor), rr)
    End Sub

    Private Sub DrawBackgroundShadows(ByVal g As Graphics)
        Dim lr As New Rectangle(2, 2, 10, Me.Height - 5)
        Dim lg As New LinearGradientBrush(lr, Color.FromArgb(30, 0, 0, 0), Color.Transparent, LinearGradientMode.Horizontal)
        lr.X -= 1
        g.FillRectangle(lg, lr)

        Dim rr As New Rectangle(Me.Width - 12, 2, 10, Me.Height - 5)
        Dim rg As New LinearGradientBrush(rr, Color.Transparent, Color.FromArgb(20, 0, 0, 0), LinearGradientMode.Horizontal)
        g.FillRectangle(rg, rr)
    End Sub

    Private Sub DrawBar(ByVal g As Graphics)

        Dim r As New Rectangle(1, 2, Me.Width - 3, Me.Height - 3)
        r.Width = CInt((Value * 1.0F / (MaxValue - MinValue) * Me.Width))
        If Me.ProgressDirection = PBarOrientation.RightToLeft Then
            r.X = Me.Width - 1 - r.Width
        End If

        Dim CurBarColor As Color = GetIntermediateColor()
        If Me.BarColorMethod = BarColorsWhen.None Then
            CurBarColor = Me.StartColor
        ElseIf Me.BarColorMethod = BarColorsWhen.OnThreshold Then
            Dim LastHiTH As Integer
            For Each Threshold As ThresholdItem In Me.BarThresholds
                If Threshold.ThresholdPosition > LastHiTH And Value >= Threshold.ThresholdPosition And Threshold.ColorizeBar Then
                    LastHiTH = Threshold.ThresholdPosition
                    CurBarColor = Threshold.ThresholdColor
                End If
            Next
            If LastHiTH = 0 Then
                CurBarColor = Me.StartColor
            End If
        End If

        If ProgressBarStyle = BarStyle.Continuous Then
            g.FillRectangle(New SolidBrush(CurBarColor), r)
            DrawBarShadows(g)
        Else
            Dim Blocks() As Rectangle = GetBlocks()
            For Each b As Rectangle In Blocks
                g.FillRectangle(New SolidBrush(CurBarColor), b)
            Next
        End If
    End Sub

    Private Sub DrawMarqueeBar(ByVal g As Graphics)
        Dim r As New Rectangle(mMarqueePosition, 0, 120, Me.Height)
        Dim lgb As New LinearGradientBrush(r, Color.White, Color.White, LinearGradientMode.Horizontal)
        Dim MarqeeColor As Color
        Dim MarqueeCenterPx As Integer = r.X + (r.Width / 2)
        If Me.BarColorMethod = BarColorsWhen.OnThreshold Then
            Dim LastHiTH As Integer = 0
            For Each Threshold As ThresholdItem In Me.BarThresholds
                If CInt((Threshold.ThresholdPosition * 1.0F / (MaxValue - MinValue) * Me.Width)) > LastHiTH And MarqueeCenterPx >= CInt((Threshold.ThresholdPosition * 1.0F / (MaxValue - MinValue) * Me.Width)) And Threshold.ColorizeBar Then
                    LastHiTH = CInt((Threshold.ThresholdPosition * 1.0F / (MaxValue - MinValue) * Me.Width))
                    MarqeeColor = Threshold.ThresholdColor
                End If
            Next
            If LastHiTH = 0 Then
                MarqeeColor = Me.StartColor
            End If
        Else
            MarqeeColor = Me.StartColor
        End If

        Dim cb As New ColorBlend(4)
        cb.Colors = New Color() {Color.Transparent, MarqeeColor, MarqeeColor, Color.Transparent}
        cb.Positions = New Single() {0.0F, 0.5F, 0.6F, 1.0F}
        lgb.InterpolationColors = cb

        Dim clip As New Rectangle(1, 1, Me.Width - 2, Me.Height - 2)
        g.SetClip(clip)
        g.FillRectangle(lgb, r)
        g.ResetClip()
    End Sub

    Private Sub DrawBarShadows(ByVal g As Graphics)
        Dim lr As New Rectangle(1, 2, 15, Me.Height - 3)
        If Me.ProgressDirection = PBarOrientation.RightToLeft Then
            lr.X = ExcessPixels - 1
        Else
            lr.X = 1
        End If
        Dim lg As New LinearGradientBrush(lr, Color.White, Color.White, LinearGradientMode.Horizontal)
        Dim lc As New ColorBlend(3)
        lc.Colors = New Color() {Color.Transparent, Color.FromArgb(40, 0, 0, 0), Color.Transparent}
        lc.Positions = New Single() {0.0F, 0.2F, 1.0F}
        lg.InterpolationColors = lc
        lr.X -= 1
        g.FillRectangle(lg, lr)


        Dim rr As New Rectangle(Me.Width - 3, 2, 15, Me.Height - 3)
        If Me.ProgressDirection = PBarOrientation.RightToLeft Then
            rr.X = Me.Width - 16
        Else
            rr.X = CInt((Value * 1.0F / (MaxValue - MinValue) * Me.Width)) - 14
        End If
        Dim rg As New LinearGradientBrush(rr, Color.Black, Color.Black, LinearGradientMode.Horizontal)
        Dim rc As New ColorBlend(3)
        rc.Colors = New Color() {Color.Transparent, Color.FromArgb(40, 0, 0, 0), Color.Transparent}
        rc.Positions = New Single() {0.0F, 0.8F, 1.0F}
        rg.InterpolationColors = rc
        g.FillRectangle(rg, rr)
    End Sub

    Private Sub DrawHighlight(ByVal g As Graphics)
        Dim tr As New Rectangle(1, 1, Me.Width - 1, 6)
        Dim tp As GraphicsPath = RoundRect(tr, 2, 2, 0, 0)

        g.SetClip(tp)
        Dim tg As New LinearGradientBrush(tr, Me.TopHighlightColor, Color.FromArgb(128, Me.TopHighlightColor), LinearGradientMode.Vertical)
        g.FillPath(tg, tp)
        g.ResetClip()

        Dim br As New Rectangle(1, Me.Height - 8, Me.Width - 1, 6)
        Dim bp As GraphicsPath = RoundRect(br, 0, 0, 2, 2)

        g.SetClip(bp)
        Dim bg As New LinearGradientBrush(br, Color.Transparent, Color.FromArgb(100, Me.HighlightColor), LinearGradientMode.Vertical)
        g.FillPath(bg, bp)
        g.ResetClip()
    End Sub

    Private Sub DrawInnerStroke(ByVal g As Graphics)
        Dim r As Rectangle = Me.ClientRectangle
        r.X += 1
        r.Y += 1
        r.Width -= 3
        r.Height -= 3
        Dim rr As GraphicsPath = RoundRect(r, 2, 2, 2, 2)
        g.DrawPath(New Pen(Color.FromArgb(100, Color.White)), rr)
    End Sub

    Private Sub DrawGlow(ByVal g As Graphics)
        Dim ClipperPosX As Integer = 1
        Dim r As New Rectangle(mGlowPosition, 0, 60, Me.Height)
        If Me.ProgressDirection = PBarOrientation.RightToLeft Then
            ClipperPosX = ExcessPixels
        End If
        Dim lgb As New LinearGradientBrush(r, Color.White, Color.White, LinearGradientMode.Horizontal)

        Dim cb As New ColorBlend(4)
        cb.Colors = New Color() {Color.Transparent, Me.GlowColor, Me.GlowColor, Color.Transparent}
        cb.Positions = New Single() {0.0F, 0.5F, 0.6F, 1.0F}
        lgb.InterpolationColors = cb

        Dim clip As New Rectangle(1, 2, Me.Width - 3, Me.Height - 3)
        clip.Width = CInt((Value * 1.0F / (MaxValue - MinValue) * Me.Width))
        clip.X = ClipperPosX
        g.SetClip(clip)
        g.FillRectangle(lgb, r)
        g.ResetClip()
    End Sub

    Private Sub DrawText(ByVal g As Graphics)
        If ShowText = True Then
            Dim showString As String = DisplayText.Replace("%P", Value)

            Dim Br As Brush = New SolidBrush(DisplayTextColor)
            Dim Rect As Rectangle = New Rectangle(0, 0, Width, Height)
            Dim tFormat As StringFormat = New StringFormat()
            tFormat.LineAlignment = StringAlignment.Center
            tFormat.Alignment = StringAlignment.Center
            g.DrawString(showString, DisplayTextFont, Br, Rect, tFormat)
        End If
    End Sub

    Private Sub DrawThresholds(ByVal g As Graphics)
        For Each Threshold As ThresholdItem In mcThresholds
            If Threshold.ShowThreshold Then
                Dim linePen As Pen = Nothing
                Dim pos As Integer

                If Threshold.ColorizeThreshold Then
                    linePen = New Pen(Threshold.ThresholdColor, 1)
                Else
                    linePen = New Pen(Color.FromArgb(178, 178, 178), 1)
                End If
                Dim THIPos As Integer = Threshold.ThresholdPosition
                If Threshold.ThresholdPosition > MaxValue Then
                    THIPos = MaxValue
                ElseIf Threshold.ThresholdPosition < MinValue Then
                    THIPos = MinValue
                End If
                If McBarOriant = PBarOrientation.RightToLeft Then
                    pos = Me.Width - CInt(((CDbl(Me.Width) - 2) * (THIPos - MinValue)) / (MaxValue - MinValue))
                Else
                    pos = CInt(((CDbl(Me.Width) - 2) * (THIPos - MinValue)) / (MaxValue - MinValue))
                End If
                g.DrawLine(linePen, pos, 0, pos, Me.Height)
            End If
        Next
    End Sub
#End Region

#Region "- Functions -"

    Private Function RoundRect(ByVal r As RectangleF, ByVal r1 As Single, ByVal r2 As Single, ByVal r3 As Single, ByVal r4 As Single) As GraphicsPath
        Dim x As Single = r.X, y As Single = r.Y, w As Single = r.Width, h As Single = r.Height
        Dim rr As New GraphicsPath()
        rr.AddBezier(x, y + r1, x, y, x + r1, y, _
        x + r1, y)
        rr.AddLine(x + r1, y, x + w - r2, y)
        rr.AddBezier(x + w - r2, y, x + w, y, x + w, y + r2, _
        x + w, y + r2)
        rr.AddLine(x + w, y + r2, x + w, y + h - r3)
        rr.AddBezier(x + w, y + h - r3, x + w, y + h, x + w - r3, y + h, _
        x + w - r3, y + h)
        rr.AddLine(x + w - r3, y + h, x + r4, y + h)
        rr.AddBezier(x + r4, y + h, x, y + h, x, y + h - r4, _
        x, y + h - r4)
        rr.AddLine(x, y + h - r4, x, y + r1)
        Return rr
    End Function

    Private Function InDesignMode() As Boolean
        Return (LicenseManager.UsageMode = LicenseUsageMode.Designtime)
    End Function

    Private Function GetIntermediateColor() As Color
        Dim c As Color = Me.StartColor
        Dim c2 As Color = Me.EndColor

        Dim pc As Single = Me.Value * 1.0F / (Me.MaxValue - Me.MinValue)

        Dim ca As Integer = c.A, cr As Integer = c.R, cg As Integer = c.G, cb As Integer = c.B
        Dim c2a As Integer = c2.A, c2r As Integer = c2.R, c2g As Integer = c2.G, c2b As Integer = c2.B

        Dim a As Integer = CInt(Math.Abs(ca + (ca - c2a) * pc))
        Dim r As Integer = CInt(Math.Abs(cr - ((cr - c2r) * pc)))
        Dim g As Integer = CInt(Math.Abs(cg - ((cg - c2g) * pc)))
        Dim b As Integer = CInt(Math.Abs(cb - ((cb - c2b) * pc)))

        If a > 255 Then
            a = 255
        End If
        If r > 255 Then
            r = 255
        End If
        If g > 255 Then
            g = 255
        End If
        If b > 255 Then
            b = 255
        End If

        Return (Color.FromArgb(a, r, g, b))
    End Function

    Private Function GetBlocks() As Rectangle()
        Dim b As New List(Of Rectangle)

        Dim w As Integer = Me.BlockSize
        Dim h As Integer = Me.Height - 3
        Dim r As Rectangle

        Dim x As Integer = 1
        Dim stopX As Integer = CInt(Me.Value / (Me.MaxValue - Me.MinValue) * Me.Width)
        While (x + w <= stopX)
            r = New Rectangle(x, 2, w, h)
            b.Add(r)

            x += Me.BlockSize + Me.BlockSpacing
        End While

        Return b.ToArray
    End Function
#End Region

#Region "- Other -"

    Private Sub ProgressBar_Paint(ByVal sender As Object, ByVal e As PaintEventArgs)
        ExcessPixels = Me.Width - CInt((Value * 1.0F / (MaxValue - MinValue) * Me.Width))
        e.Graphics.SmoothingMode = SmoothingMode.AntiAlias
        e.Graphics.InterpolationMode = InterpolationMode.HighQualityBicubic
        DrawBackground(e.Graphics)
        DrawBackgroundShadows(e.Graphics)
        If Me.ProgressBarStyle = BarStyle.Marquee Then
            DrawMarqueeBar(e.Graphics)
        Else
            DrawBar(e.Graphics)
        End If
        DrawHighlight(e.Graphics)
        DrawInnerStroke(e.Graphics)
        DrawGlow(e.Graphics)
        DrawOuterStroke(e.Graphics)
        DrawThresholds(e.Graphics)
        DrawText(e.Graphics)
    End Sub

    Private Sub mGlowAnimation_Tick(ByVal sender As Object, ByVal e As EventArgs)
        If Me.Animate Then
            If Me.ProgressDirection = PBarOrientation.LeftToRight Then
                mGlowPosition += 4
                If mGlowPosition > Me.Width Then
                    mGlowPosition = -300
                End If
            Else
                mGlowPosition -= 4
                If mGlowPosition < -60 Then
                    mGlowPosition = Me.Width + 300
                End If
            End If
        Else
            mGlowAnimation.Stop()
            mGlowPosition = -320
        End If
        Me.Invalidate()
    End Sub

    Private Sub mMarqueeAnimation_Tick(ByVal sender As Object, ByVal e As EventArgs)
        If Me.ProgressBarStyle = BarStyle.Marquee Then
            If Me.ProgressDirection = PBarOrientation.LeftToRight Then
                mMarqueePosition += 4
                If mMarqueePosition > Me.Width Then
                    mMarqueePosition = -120
                End If
            Else
                mMarqueePosition -= 4
                If mMarqueePosition < -120 Then
                    mMarqueePosition = Me.Width + 120
                End If
            End If
        Else
            mMarqueeAnimation.Stop()
            mMarqueePosition = -120
        End If
        Me.Invalidate()
    End Sub
#End Region

#Region "- Events -"
    ''' When the Value property is changed.
    Public Delegate Sub ValueChangedHandler(ByVal sender As Object, ByVal e As EventArgs)
    ''' When the Value property is changed.
    Public Event ValueChanged As ValueChangedHandler
    ''' When the MinValue property is changed.
    Public Delegate Sub MinChangedHandler(ByVal sender As Object, ByVal e As EventArgs)
    ''' When the MinValue property is changed.
    Public Event MinChanged As MinChangedHandler
    ''' When the MaxValue property is changed.
    Public Delegate Sub MaxChangedHandler(ByVal sender As Object, ByVal e As EventArgs)
    ''' When the MaxValue property is changed.
    Public Event MaxChanged As MaxChangedHandler
#End Region

End Class

<Serializable()> _
Public Class ThresholdItem

    Dim McTHColor As Color = Color.FromArgb(0, 211, 40)
    <Description("The Threshold Color"), Category("Settings")> _
    Public Property ThresholdColor() As Color
        Get
            Return McTHColor
        End Get
        Set(ByVal value As Color)
            McTHColor = value
        End Set
    End Property

    Dim McTHColorize As Boolean = False
    <Description("The Threshold Color"), Category("Settings"), DefaultValue(GetType(Boolean), "False")> _
    Public Property ColorizeThreshold() As Boolean
        Get
            Return McTHColorize
        End Get
        Set(ByVal value As Boolean)
            McTHColorize = value
        End Set
    End Property

    Dim McShowThreshold As Boolean = False
    <Description("Sets Whether to Show This Threshold or not."), Category("Settings"), DefaultValue(GetType(Boolean), "False")> _
    Public Property ShowThreshold() As Boolean
        Get
            Return McShowThreshold
        End Get
        Set(ByVal value As Boolean)
            McShowThreshold = value
        End Set
    End Property

    Dim McColorizeOnTH As Boolean = False
    <Description("Sets whether the Progress Bar Will Change Color When It Hits this Threshold"), Category("Settings"), DefaultValue(GetType(Boolean), "False")> _
    Public Property ColorizeBar() As Boolean
        Get
            Return McColorizeOnTH
        End Get
        Set(ByVal value As Boolean)
            McColorizeOnTH = value
        End Set
    End Property

    Dim McThresholdPos As Integer = 50
    <Description("Sets The Position Of The Threshold"), Category("Settings"), DefaultValue(GetType(Boolean), "False")> _
    Public Property ThresholdPosition() As Integer
        Get
            Return McThresholdPos
        End Get
        Set(ByVal value As Integer)
            McThresholdPos = value
        End Set
    End Property

End Class
'End Namespace