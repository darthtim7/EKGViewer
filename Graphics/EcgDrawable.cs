using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.Maui.Graphics;

namespace EKGViewer.Graphics;

public sealed class ECGDrawable : IDrawable
{
    private double[] _samples = Array.Empty<double>();
    private double _sampleRateHz = 250.0;
    private string _signalName = "";

    public double SecondsToDisplay { get; set; } = 10.0;

    public void SetSignal(double[] samples, double sampleRateHz, string signalName)
    {
        _samples = samples;
        _sampleRateHz = sampleRateHz > 0 ? sampleRateHz : 250.0;
        _signalName = signalName;
    }

    public void Draw(ICanvas canvas, RectF dirtyRect)
    {
        canvas.FillColor = Colors.White;
        canvas.FillRectangle(dirtyRect);

        DrawGrid(canvas, dirtyRect);

        if (_samples.Length < 2)
        {
            DrawCenteredText(canvas, dirtyRect, "No signal loaded");
            return;
        }

        var sampleCountToDisplay = Math.Min(
            _samples.Length,
            Math.Max(2, (int)Math.Round(_sampleRateHz * SecondsToDisplay)));

        var left = 10f;
        var right = 10f;
        var top = 10f;
        var bottom = 28f;

        var plotWidth = Math.Max(1, dirtyRect.Width - left - right);
        var plotHeight = Math.Max(1, dirtyRect.Height - top - bottom);
        var midY = top + plotHeight / 2f;

        var visible = _samples.Take(sampleCountToDisplay).ToArray();

        var min = visible.Min();
        var max = visible.Max();
        var center = (min + max) / 2.0;
        var range = Math.Max(max - min, 0.001);

        var path = new PathF();

        for (var i = 0; i < visible.Length; i++)
        {
            var x = left + (float)i / (visible.Length - 1) * plotWidth;

            var normalized = (visible[i] - center) / range;
            var y = midY - (float)(normalized * plotHeight * 0.85);

            if (i == 0)
                path.MoveTo(x, y);
            else
                path.LineTo(x, y);
        }

        canvas.StrokeColor = Colors.Black;
        canvas.StrokeSize = 2;
        canvas.DrawPath(path);

        DrawTimeLabels(canvas, dirtyRect, left, plotWidth);
    }

    private static void DrawGrid(ICanvas canvas, RectF rect)
    {
        const float minor = 10f;
        const float major = minor * 5f;

        canvas.StrokeSize = 1;

        canvas.StrokeColor = Color.FromArgb("#FFE6E6");
        for (float x = 0; x <= rect.Width; x += minor)
            canvas.DrawLine(x, 0, x, rect.Height);

        for (float y = 0; y <= rect.Height; y += minor)
            canvas.DrawLine(0, y, rect.Width, y);

        canvas.StrokeColor = Color.FromArgb("#FFB8B8");
        for (float x = 0; x <= rect.Width; x += major)
            canvas.DrawLine(x, 0, x, rect.Height);

        for (float y = 0; y <= rect.Height; y += major)
            canvas.DrawLine(0, y, rect.Width, y);
    }

    private void DrawTimeLabels(ICanvas canvas, RectF rect, float left, float plotWidth)
    {
        canvas.FontColor = Colors.DimGray;
        canvas.FontSize = 11;

        var seconds = (int)Math.Floor(SecondsToDisplay);

        for (var s = 0; s <= seconds; s++)
        {
            var x = left + (float)(s / SecondsToDisplay) * plotWidth;

            canvas.DrawString(
                $"{s}s",
                x - 12,
                rect.Height - 22,
                24,
                16,
                HorizontalAlignment.Center,
                VerticalAlignment.Center);
        }
    }

    private static void DrawCenteredText(ICanvas canvas, RectF rect, string text)
    {
        canvas.FontColor = Colors.Gray;
        canvas.FontSize = 16;

        canvas.DrawString(
            text,
            rect.X,
            rect.Y + rect.Height / 2 - 10,
            rect.Width,
            20,
            HorizontalAlignment.Center,
            VerticalAlignment.Center);
    }
}
