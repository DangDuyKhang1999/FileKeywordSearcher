using System.Drawing.Drawing2D;

namespace FileKeywordSearcher;

public class RoundedPanel : Panel
{
    public int CornerRadius { get; set; } = 16;
    public Color BorderColor { get; set; } = Color.Transparent;
    public Color HighlightColor { get; set; } = Color.FromArgb(255, 255, 255);
    public Color ShadowColor { get; set; } = Color.FromArgb(198, 216, 203);
    public bool ShowLeafPattern { get; set; }
    public RoundedPanel() { DoubleBuffered = true; ResizeRedraw = true; }
    protected override void OnResize(EventArgs eventargs)
    {
        base.OnResize(eventargs);
        if (ClientRectangle.Width <= 1 || ClientRectangle.Height <= 1) return;
        using GraphicsPath path = RoundedPath(ClientRectangle, CornerRadius);
        Region = new Region(path);
    }
    protected override void OnPaint(PaintEventArgs e)
    {
        base.OnPaint(e);
        using GraphicsPath path = RoundedPath(ClientRectangle, CornerRadius);
        using Pen borderPen = new(BorderColor, 1F);
        e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
        e.Graphics.DrawPath(borderPen, path);
        if (ShowLeafPattern) DrawLeaves(e.Graphics);
    }

    private void DrawLeaves(Graphics graphics)
    {
        using Pen stemPen = new(Color.FromArgb(38, 86, 151, 105), 2F);
        using SolidBrush leafBrush = new(Color.FromArgb(32, 91, 169, 116));
        DrawLeafBranch(graphics, new Point(28, Height - 18), -1, stemPen, leafBrush, 1.35F);
        DrawLeafBranch(graphics, new Point(Width - 26, 28), 1, stemPen, leafBrush, 1.25F);
        DrawLeafBranch(graphics, new Point(Width - 28, Height - 18), 1, stemPen, leafBrush, 1.05F);
        DrawLeafBranch(graphics, new Point(28, 22), -1, stemPen, leafBrush, 0.95F);
    }

    private static void DrawLeafBranch(Graphics graphics, Point origin, int direction, Pen stemPen, Brush leafBrush, float scale = 1F)
    {
        Point end = new(origin.X + direction * (int)(92 * scale), origin.Y - (int)(72 * scale));
        graphics.DrawBezier(stemPen, origin, new Point(origin.X + direction * (int)(30 * scale), origin.Y - (int)(12 * scale)), new Point(end.X - direction * (int)(24 * scale), end.Y + (int)(18 * scale)), end);
        for (int i = 1; i <= 6; i++)
        {
            float t = i / 7F;
            int x = (int)(origin.X + (end.X - origin.X) * t);
            int y = (int)(origin.Y + (end.Y - origin.Y) * t);
            int side = i % 2 == 0 ? 1 : -1;
            int leafWidth = Math.Max(9, (int)(22 * scale));
            int leafHeight = Math.Max(6, (int)(13 * scale));
            Rectangle leaf = new(x + side * 2 - leafWidth / 2, y + side * (int)(8 * scale) - leafHeight / 2, leafWidth, leafHeight);
            graphics.FillEllipse(leafBrush, leaf);
        }
    }
    internal static GraphicsPath RoundedPath(Rectangle bounds, int radius)
    {
        bounds.Width--; bounds.Height--;
        int d = Math.Max(2, radius * 2);
        GraphicsPath path = new();
        path.AddArc(bounds.Left, bounds.Top, d, d, 180, 90);
        path.AddArc(bounds.Right - d, bounds.Top, d, d, 270, 90);
        path.AddArc(bounds.Right - d, bounds.Bottom - d, d, d, 0, 90);
        path.AddArc(bounds.Left, bounds.Bottom - d, d, d, 90, 90);
        path.CloseFigure();
        return path;
    }
}

public class ModernButton : Button
{
    public int CornerRadius { get; set; } = 10;
    public Color BorderColor { get; set; } = Color.Transparent;
    public ModernButton()
    {
        FlatStyle = FlatStyle.Flat;
        FlatAppearance.BorderSize = 0;
        UseVisualStyleBackColor = false;
        Cursor = Cursors.Hand;
        TabStop = false;
    }
    protected override void OnResize(EventArgs e)
    {
        base.OnResize(e);
        using GraphicsPath path = RoundedPanel.RoundedPath(ClientRectangle, CornerRadius);
        Region?.Dispose();
        Region = new Region(path);
    }
    protected override void OnPaint(PaintEventArgs e)
    {
        base.OnPaint(e);
        using GraphicsPath path = RoundedPanel.RoundedPath(ClientRectangle, CornerRadius);
        using Pen pen = new(BorderColor);
        e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
        e.Graphics.DrawPath(pen, path);
    }
}

public class ModernProgressBar : Control
{
    private int _value;
    private int _displayValue;
    private int _shimmerOffset;
    private bool _isIndeterminate;
    private readonly System.Windows.Forms.Timer _animationTimer;
    public int Minimum { get; set; }
    public int Maximum { get; set; } = 100;
    public int Step { get; set; } = 1;
    public int Value
    {
        get => _value;
        set
        {
            _value = Math.Clamp(value, Minimum, Maximum);
            if (!_animationTimer.Enabled) _animationTimer.Start();
        }
    }
    public Color TrackColor { get; set; } = Color.FromArgb(218, 234, 222);
    public Color ProgressColor { get; set; } = Color.FromArgb(111, 187, 137);
    public bool IsIndeterminate
    {
        get => _isIndeterminate;
        set
        {
            _isIndeterminate = value;
            if (value && !_animationTimer.Enabled) _animationTimer.Start();
            Invalidate();
        }
    }

    public ModernProgressBar()
    {
        DoubleBuffered = true;
        ResizeRedraw = true;
        _animationTimer = new System.Windows.Forms.Timer { Interval = 16 };
        _animationTimer.Tick += (_, _) =>
        {
            int distance = _value - _displayValue;
            if (distance != 0)
                _displayValue += Math.Sign(distance) * Math.Max(1, Math.Abs(distance) / 7);
            _shimmerOffset = (_shimmerOffset + 7) % Math.Max(1, Width + 90);
            Invalidate();
            if (!_isIndeterminate && _displayValue == _value && (_value == Minimum || _value == Maximum))
                _animationTimer.Stop();
        };
    }
    protected override void OnPaint(PaintEventArgs e)
    {
        e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
        Rectangle track = new(0, 0, Width - 1, Height - 1);
        using GraphicsPath trackPath = RoundedPanel.RoundedPath(track, Height / 2);
        using SolidBrush trackBrush = new(TrackColor);
        e.Graphics.FillPath(trackBrush, trackPath);
        if (_isIndeterminate)
        {
            int segmentWidth = Math.Max(80, Width / 4);
            int segmentX = _shimmerOffset - segmentWidth;
            Rectangle segment = new(segmentX, 0, segmentWidth, Height - 1);
            using GraphicsPath segmentPath = RoundedPanel.RoundedPath(segment, Height / 2);
            using LinearGradientBrush segmentBrush = new(segment, Color.FromArgb(80, ProgressColor), ProgressColor, LinearGradientMode.Horizontal);
            e.Graphics.FillPath(segmentBrush, segmentPath);
            return;
        }

        int fillWidth = Maximum <= Minimum ? 0 : (int)((Width - 1) * (_displayValue - Minimum) / (double)(Maximum - Minimum));
        if (fillWidth < 2) return;
        Rectangle fill = new(0, 0, fillWidth, Height - 1);
        using GraphicsPath fillPath = RoundedPanel.RoundedPath(fill, Height / 2);
        using SolidBrush fillBrush = new(ProgressColor);
        e.Graphics.FillPath(fillBrush, fillPath);

        e.Graphics.SetClip(fillPath);
        int shimmerX = _shimmerOffset - 90;
        using LinearGradientBrush shimmer = new(
            new Rectangle(shimmerX, 0, 90, Height),
            Color.FromArgb(0, Color.White),
            Color.FromArgb(145, Color.White),
            LinearGradientMode.Horizontal);
        e.Graphics.FillRectangle(shimmer, shimmerX, 0, 90, Height);
        e.Graphics.ResetClip();
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing) _animationTimer.Dispose();
        base.Dispose(disposing);
    }
}
