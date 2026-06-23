namespace EKGViewer;
using EKGViewer.File;
using EKGViewer.Graphics;
using Microsoft.Maui.Graphics;

public partial class MainPage : ContentPage
{
	int count = 0;
	private readonly WFDBReader _reader = new("Ekg");
    private readonly ECGDrawable _drawable = new();
    private readonly HeartModelDrawable _heartDrawable = new();
    private bool _loading;

 
	public MainPage()
	{
		InitializeComponent();

		EcgGraph.Drawable = _drawable;
        HeartModel.Drawable = _heartDrawable;
        HeartModel.StartInteraction += OnHeartStartInteraction;
        HeartModel.DragInteraction += OnHeartDragInteraction;
        HeartModel.EndInteraction += OnHeartEndInteraction;
        Loaded += OnLoaded;

	}

    private PointF? _lastHeartTouch;

    private void OnHeartStartInteraction(object? sender, TouchEventArgs e)
    {
        _lastHeartTouch = e.Touches.FirstOrDefault();
    }

    private void OnHeartDragInteraction(object? sender, TouchEventArgs e)
    {
        if (_lastHeartTouch is not PointF lastTouch || e.Touches.Length == 0)
            return;

        var currentTouch = e.Touches[0];
        _heartDrawable.Rotate(currentTouch.X - lastTouch.X, currentTouch.Y - lastTouch.Y);
        _lastHeartTouch = currentTouch;
        HeartModel.Invalidate();
    }

    private void OnHeartEndInteraction(object? sender, TouchEventArgs e)
    {
        _lastHeartTouch = null;
    }

	private async void OnLoaded(object? sender, EventArgs e)
    {
        Loaded -= OnLoaded;
        await LoadRecordListAsync();
    }

    private async Task LoadRecordListAsync()
    {
        try
        {
            _loading = true;

            var records = await _reader.ReadRecordManifestAsync();

            RecordPicker.ItemsSource = records;
            InfoLabel.Text = records.Count == 0
                ? "No records listed in Resources/Raw/Ekg/records.txt."
                : "Select a record.";

            _loading = false;

            if (records.Count > 0)
                RecordPicker.SelectedIndex = 0;
        }
        catch (Exception ex)
        {
            _loading = false;
            InfoLabel.Text = $"Startup error: {ex.Message}";
        }
    }

    private async void OnRecordChanged(object? sender, EventArgs e)
    {
        if (_loading || RecordPicker.SelectedItem is not string recordName)
            return;

        await LoadHeaderAndSignalAsync(recordName, 0);
    }

    private async void OnSignalChanged(object? sender, EventArgs e)
    {
        if (_loading || RecordPicker.SelectedItem is not string || SignalPicker.SelectedIndex < 0)
            return;

        await LoadSignalOnlyAsync();
    }

    private async Task LoadHeaderAndSignalAsync(string recordName, int signalIndex)
    {
        try
        {
            _loading = true;

            var header = await _reader.ReadHeaderAsync(recordName);
            if (header.SignalCount == 0)
                throw new InvalidOperationException("The WFDB header contains no signals.");

            SignalPicker.ItemsSource = header.Signals.Select(s => s.SignalName).ToList();
            SignalPicker.SelectedIndex = Math.Clamp(signalIndex, 0, header.SignalCount - 1);

            _loading = false;

            await LoadSignalOnlyAsync();
        }
        catch (Exception ex)
        {
            _loading = false;
            InfoLabel.Text = $"Record load error: {ex.Message}";
        }
    }

    private async Task LoadSignalOnlyAsync()
    {
        if (RecordPicker.SelectedItem is not string recordName)
            return;

        try
        {
            InfoLabel.Text = "Loading signal...";

            var signal = await _reader.LoadSignalAsync(recordName, SignalPicker.SelectedIndex);

            _drawable.SetSignal(signal.Samples, signal.SampleRateHz, signal.SelectedSignalName);
            EcgGraph.Invalidate();

            var shownSeconds = Math.Min(_drawable.SecondsToDisplay,
                signal.Samples.Length / signal.SampleRateHz);

            InfoLabel.Text =
                $"{signal.Name} | {signal.SelectedSignalName} | {signal.SampleRateHz:g} Hz | " +
                $"{signal.Samples.Length:n0} samples loaded | showing first {shownSeconds:0.#} s | " +
                $"{signal.Units}, auto-scaled";
        }
        catch (Exception ex)
        {
            InfoLabel.Text = $"Signal load error: {ex.Message}";
        }
    }

	private void OnCounterClicked(object? sender, EventArgs e)
	{
		count++;
		/*
		if (count == 1)
			CounterBtn.Text = $"Clicked {count} time";
		else
			CounterBtn.Text = $"Clicked {count} times";

		SemanticScreenReader.Announce(CounterBtn.Text);
		*/
	}
}
