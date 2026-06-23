using Microsoft.Maui.Graphics;

namespace EKGViewer.Graphics;

public sealed class HeartModelDrawable : IDrawable
{
    private static readonly Color BackgroundTop = Color.FromArgb("#111827");
    private static readonly Color BackgroundBottom = Color.FromArgb("#020617");

    public void Draw(ICanvas canvas, RectF dirtyRect)
    {
        DrawBackground(canvas, dirtyRect);

        var size = Math.Min(dirtyRect.Width, dirtyRect.Height) * 0.74f;
        if (size <= 0)
            return;

        var centerX = dirtyRect.X + dirtyRect.Width / 2f;
        var centerY = dirtyRect.Y + dirtyRect.Height / 2f + size * 0.03f;
        var radiusX = size * 0.42f;
        var radiusY = size * 0.36f;
        var overlap = size * 0.04f;

        DrawQuarterEllipsoid(canvas, centerX, centerY, -1, -1, radiusX, radiusY, overlap, Color.FromArgb("#D9467A"), "Left atrium");
        DrawQuarterEllipsoid(canvas, centerX, centerY, 1, -1, radiusX, radiusY, overlap, Color.FromArgb("#F9736B"), "Right atrium");
        DrawQuarterEllipsoid(canvas, centerX, centerY, -1, 1, radiusX, radiusY * 1.12f, overlap, Color.FromArgb("#7C3AED"), "Left ventricle");
        DrawQuarterEllipsoid(canvas, centerX, centerY, 1, 1, radiusX, radiusY * 1.12f, overlap, Color.FromArgb("#2563EB"), "Right ventricle");

        DrawCenterSeams(canvas, centerX, centerY, radiusX, radiusY * 1.1f);
        DrawTitle(canvas, dirtyRect);
    }

    private static void DrawBackground(ICanvas canvas, RectF rect)
    {
        var paint = new LinearGradientPaint
        {
            StartColor = BackgroundTop,
            EndColor = BackgroundBottom,
            StartPoint = new Point(0, 0),
            EndPoint = new Point(0, 1)
        };

        canvas.SetFillPaint(paint, rect);
        canvas.FillRectangle(rect);
    }

    private static void DrawQuarterEllipsoid(
        ICanvas canvas,
        float centerX,
        float centerY,
        int horizontalDirection,
        int verticalDirection,
        float radiusX,
        float radiusY,
        float overlap,
        Color baseColor,
        string label)
    {
        var path = CreateQuarterPath(centerX, centerY, horizontalDirection, verticalDirection, radiusX, radiusY, overlap);

        canvas.SaveState();
        canvas.SetShadow(new SizeF(horizontalDirection * 5, verticalDirection * 5), 12, Colors.Black.WithAlpha(0.35f));
        canvas.FillColor = baseColor.WithAlpha(0.96f);
        canvas.FillPath(path);
        canvas.RestoreState();

        canvas.StrokeColor = Colors.White.WithAlpha(0.24f);
        canvas.StrokeSize = 1.5f;
        canvas.DrawPath(path);

        DrawLayeredHighlight(canvas, centerX, centerY, horizontalDirection, verticalDirection, radiusX, radiusY, baseColor);
        DrawLabel(canvas, centerX, centerY, horizontalDirection, verticalDirection, radiusX, radiusY, label);
    }

    private static PathF CreateQuarterPath(float centerX, float centerY, int horizontalDirection, int verticalDirection, float radiusX, float radiusY, float overlap)
    {
        var xEdge = centerX + horizontalDirection * radiusX;
        var yEdge = centerY + verticalDirection * radiusY;
        var xInner = centerX - horizontalDirection * overlap;
        var yInner = centerY - verticalDirection * overlap;
        var controlX = centerX + horizontalDirection * radiusX;
        var controlY = centerY + verticalDirection * radiusY;

        var path = new PathF();
        path.MoveTo(xInner, centerY);
        path.LineTo(xEdge, centerY);
        path.CurveTo(controlX, centerY, controlX, yEdge, centerX, yEdge);
        path.LineTo(centerX, yInner);
        path.CurveTo(
            centerX + horizontalDirection * radiusX * 0.18f,
            centerY + verticalDirection * radiusY * 0.18f,
            centerX + horizontalDirection * radiusX * 0.18f,
            centerY + verticalDirection * radiusY * 0.02f,
            xInner,
            centerY);
        path.Close();
        return path;
    }

    private static void DrawLayeredHighlight(ICanvas canvas, float centerX, float centerY, int horizontalDirection, int verticalDirection, float radiusX, float radiusY, Color baseColor)
    {
        for (var i = 0; i < 7; i++)
        {
            var inset = i * 0.055f;
            var alpha = 0.12f - i * 0.012f;
            canvas.FillColor = Colors.White.WithAlpha(Math.Max(0.02f, alpha));
            canvas.FillEllipse(
                centerX + horizontalDirection * radiusX * (0.22f + inset),
                centerY + verticalDirection * radiusY * (0.20f + inset),
                radiusX * (0.52f - inset),
                radiusY * (0.42f - inset));
        }

        canvas.StrokeColor = baseColor.WithAlpha(0.75f);
        canvas.StrokeSize = 4;
        canvas.DrawEllipse(
            centerX + horizontalDirection * radiusX * 0.08f,
            centerY + verticalDirection * radiusY * 0.08f,
            radiusX * 0.82f,
            radiusY * 0.72f);
    }

    private static void DrawCenterSeams(ICanvas canvas, float centerX, float centerY, float radiusX, float radiusY)
    {
        canvas.StrokeColor = Colors.White.WithAlpha(0.30f);
        canvas.StrokeSize = 2;
        canvas.DrawLine(centerX - radiusX * 0.95f, centerY, centerX + radiusX * 0.95f, centerY);
        canvas.DrawLine(centerX, centerY - radiusY * 0.9f, centerX, centerY + radiusY * 0.98f);

        canvas.FillColor = Colors.White.WithAlpha(0.20f);
        canvas.FillCircle(centerX, centerY, 5);
    }

    private static void DrawTitle(ICanvas canvas, RectF rect)
    {
        canvas.FontColor = Colors.White.WithAlpha(0.78f);
        canvas.FontSize = 14;
        canvas.DrawString("Prototype 3D heart model", rect.X, rect.Y + 10, rect.Width, 20, HorizontalAlignment.Center, VerticalAlignment.Center);
    }

    private static void DrawLabel(ICanvas canvas, float centerX, float centerY, int horizontalDirection, int verticalDirection, float radiusX, float radiusY, string label)
    {
        canvas.FontColor = Colors.White.WithAlpha(0.88f);
        canvas.FontSize = 11;
        canvas.DrawString(
            label,
            centerX + horizontalDirection * radiusX * 0.23f - 55,
            centerY + verticalDirection * radiusY * 0.28f - 8,
            110,
            16,
            HorizontalAlignment.Center,
            VerticalAlignment.Center);
    }
}
