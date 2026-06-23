using Microsoft.Maui.Graphics;

namespace EKGViewer.Graphics;

public sealed class HeartModelDrawable : IDrawable
{
    private const float CameraDistance = 4.2f;
    private const int LatitudeSteps = 10;
    private const int LongitudeSteps = 10;

    private static readonly Color BackgroundTop = Color.FromArgb("#111827");
    private static readonly Color BackgroundBottom = Color.FromArgb("#020617");

    private float _rotationX = -0.25f;
    private float _rotationY = 0.45f;

    public void Rotate(float deltaX, float deltaY)
    {
        _rotationY += deltaX * 0.012f;
        _rotationX = Math.Clamp(_rotationX + deltaY * 0.012f, -1.35f, 1.35f);
    }

    public void Draw(ICanvas canvas, RectF dirtyRect)
    {
        DrawBackground(canvas, dirtyRect);

        var size = Math.Min(dirtyRect.Width, dirtyRect.Height) * 0.72f;
        if (size <= 0)
            return;

        var centerX = dirtyRect.X + dirtyRect.Width / 2f;
        var centerY = dirtyRect.Y + dirtyRect.Height / 2f + size * 0.06f;
        var scale = size * 0.46f;

        var patches = BuildHeartPatches(centerX, centerY, scale)
            .OrderBy(patch => patch.AverageDepth)
            .ToList();

        foreach (var patch in patches)
            DrawPatch(canvas, patch);

        DrawCenterSeams(canvas, centerX, centerY, scale);
        DrawTitle(canvas, dirtyRect);
    }

    private IEnumerable<ProjectedPatch> BuildHeartPatches(float centerX, float centerY, float scale)
    {
        foreach (var chamber in HeartChambers)
        {
            for (var lat = 0; lat < LatitudeSteps; lat++)
            {
                for (var lon = 0; lon < LongitudeSteps; lon++)
                {
                    var points = new[]
                    {
                        Project(CreatePoint(chamber, lat, lon), centerX, centerY, scale),
                        Project(CreatePoint(chamber, lat + 1, lon), centerX, centerY, scale),
                        Project(CreatePoint(chamber, lat + 1, lon + 1), centerX, centerY, scale),
                        Project(CreatePoint(chamber, lat, lon + 1), centerX, centerY, scale)
                    };

                    var normal = CalculateNormal(points[0].Rotated, points[1].Rotated, points[2].Rotated);
                    var light = Math.Clamp(0.35f + normal.Z * 0.42f - normal.X * 0.16f - normal.Y * 0.10f, 0.24f, 0.95f);
                    yield return new ProjectedPatch(points, chamber.Color, light);
                }
            }
        }
    }

    private ChamberPoint CreatePoint(HeartChamber chamber, int latitudeIndex, int longitudeIndex)
    {
        var latT = latitudeIndex / (float)LatitudeSteps;
        var lonT = longitudeIndex / (float)LongitudeSteps;
        var theta = latT * MathF.PI / 2f;
        var phi = lonT * MathF.PI / 2f;

        var horizontal = chamber.HorizontalDirection;
        var vertical = chamber.VerticalDirection;
        var x = chamber.CenterX + horizontal * chamber.RadiusX * MathF.Sin(theta) * MathF.Cos(phi);
        var y = chamber.CenterY + vertical * chamber.RadiusY * MathF.Sin(theta) * MathF.Sin(phi);
        var z = chamber.CenterZ + chamber.RadiusZ * MathF.Cos(theta);

        return new ChamberPoint(x, y, z);
    }

    private ProjectedPoint Project(ChamberPoint point, float centerX, float centerY, float scale)
    {
        var rotated = RotatePoint(point);
        var perspective = CameraDistance / (CameraDistance - rotated.Z);
        var screenX = centerX + rotated.X * scale * perspective;
        var screenY = centerY + rotated.Y * scale * perspective;

        return new ProjectedPoint(screenX, screenY, rotated);
    }

    private ChamberPoint RotatePoint(ChamberPoint point)
    {
        var cosY = MathF.Cos(_rotationY);
        var sinY = MathF.Sin(_rotationY);
        var x = point.X * cosY + point.Z * sinY;
        var z = -point.X * sinY + point.Z * cosY;

        var cosX = MathF.Cos(_rotationX);
        var sinX = MathF.Sin(_rotationX);
        var y = point.Y * cosX - z * sinX;
        z = point.Y * sinX + z * cosX;

        return new ChamberPoint(x, y, z);
    }

    private static ChamberPoint CalculateNormal(ChamberPoint a, ChamberPoint b, ChamberPoint c)
    {
        var ux = b.X - a.X;
        var uy = b.Y - a.Y;
        var uz = b.Z - a.Z;
        var vx = c.X - a.X;
        var vy = c.Y - a.Y;
        var vz = c.Z - a.Z;

        var nx = uy * vz - uz * vy;
        var ny = uz * vx - ux * vz;
        var nz = ux * vy - uy * vx;
        var length = MathF.Max(0.0001f, MathF.Sqrt(nx * nx + ny * ny + nz * nz));

        return new ChamberPoint(nx / length, ny / length, nz / length);
    }

    private static void DrawPatch(ICanvas canvas, ProjectedPatch patch)
    {
        var path = new PathF();
        path.MoveTo(patch.Points[0].X, patch.Points[0].Y);
        path.LineTo(patch.Points[1].X, patch.Points[1].Y);
        path.LineTo(patch.Points[2].X, patch.Points[2].Y);
        path.LineTo(patch.Points[3].X, patch.Points[3].Y);
        path.Close();

        canvas.FillColor = Shade(patch.Color, patch.Light);
        canvas.FillPath(path);

        canvas.StrokeColor = Colors.White.WithAlpha(0.05f);
        canvas.StrokeSize = 0.6f;
        canvas.DrawPath(path);
    }

    private static Color Shade(Color color, float light)
    {
        return new Color(
            Math.Clamp(color.Red * light, 0f, 1f),
            Math.Clamp(color.Green * light, 0f, 1f),
            Math.Clamp(color.Blue * light, 0f, 1f),
            1f);
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

    private void DrawCenterSeams(ICanvas canvas, float centerX, float centerY, float scale)
    {
        var seamColor = Colors.White.WithAlpha(0.45f);
        DrawProjectedLine(canvas, centerX, centerY, scale, new ChamberPoint(-0.78f, 0, 0.04f), new ChamberPoint(0.78f, 0, 0.04f), seamColor);
        DrawProjectedLine(canvas, centerX, centerY, scale, new ChamberPoint(0, -0.72f, 0.04f), new ChamberPoint(0, 0.86f, 0.04f), seamColor);
    }

    private void DrawProjectedLine(ICanvas canvas, float centerX, float centerY, float scale, ChamberPoint start, ChamberPoint end, Color color)
    {
        var projectedStart = Project(start, centerX, centerY, scale);
        var projectedEnd = Project(end, centerX, centerY, scale);

        canvas.StrokeColor = color;
        canvas.StrokeSize = 2.2f;
        canvas.DrawLine(projectedStart.X, projectedStart.Y, projectedEnd.X, projectedEnd.Y);
    }

    private static void DrawTitle(ICanvas canvas, RectF rect)
    {
        canvas.FontColor = Colors.White.WithAlpha(0.78f);
        canvas.FontSize = 14;
        canvas.DrawString("Drag to rotate 3D heart model", rect.X, rect.Y + 10, rect.Width, 20, HorizontalAlignment.Center, VerticalAlignment.Center);
    }

    private static readonly HeartChamber[] HeartChambers =
    {
        new(-1, -1, -0.08f, -0.08f, 0f, 0.82f, 0.68f, 0.70f, Color.FromArgb("#F05A8A")),
        new(1, -1, 0.08f, -0.08f, 0f, 0.82f, 0.68f, 0.70f, Color.FromArgb("#FF7A63")),
        new(-1, 1, -0.08f, 0.08f, 0f, 0.86f, 0.88f, 0.82f, Color.FromArgb("#8B5CF6")),
        new(1, 1, 0.08f, 0.08f, 0f, 0.86f, 0.88f, 0.82f, Color.FromArgb("#3B82F6"))
    };

    private readonly record struct HeartChamber(
        int HorizontalDirection,
        int VerticalDirection,
        float CenterX,
        float CenterY,
        float CenterZ,
        float RadiusX,
        float RadiusY,
        float RadiusZ,
        Color Color);

    private readonly record struct ChamberPoint(float X, float Y, float Z);

    private readonly record struct ProjectedPoint(float X, float Y, ChamberPoint Rotated);

    private sealed class ProjectedPatch
    {
        public ProjectedPatch(ProjectedPoint[] points, Color color, float light)
        {
            Points = points;
            Color = color;
            Light = light;
            AverageDepth = points.Average(point => point.Rotated.Z);
        }

        public ProjectedPoint[] Points { get; }

        public Color Color { get; }

        public float Light { get; }

        public float AverageDepth { get; }
    }
}
