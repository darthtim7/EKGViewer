using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Buffers.Binary;
using System.Globalization;
using System.Text.RegularExpressions;
using Microsoft.Maui.Storage;

namespace EKGViewer.File
{
    public sealed class WFDBReader
    {
        private readonly string _assetFolder;

        public WFDBReader(string assetFolder)
        {
            _assetFolder = assetFolder.Trim('/', '\\');
        }

        public async Task<List<string>> ReadRecordManifestAsync()
        {
            var path = $"{_assetFolder}/records.txt";

            await using var stream = await FileSystem.Current.OpenAppPackageFileAsync(path);
            using var reader = new StreamReader(stream);

            var records = new List<string>();

            while (await reader.ReadLineAsync() is { } line)
            {
                var clean = line.Split('#')[0].Trim();

                if (!string.IsNullOrWhiteSpace(clean))
                    records.Add(clean);
            }

            return records;
        }

        public async Task<WFDBHeader> ReadHeaderAsync(string recordName)
        {
            var path = $"{_assetFolder}/{recordName}.hea";

            await using var stream = await FileSystem.Current.OpenAppPackageFileAsync(path);
            using var reader = new StreamReader(stream);

            var lines = new List<string>();

            while (await reader.ReadLineAsync() is { } line)
            {
                var trimmed = line.Trim();

                if (trimmed.Length == 0 || trimmed.StartsWith('#'))
                    continue;

                lines.Add(trimmed);
            }

            if (lines.Count == 0)
                throw new InvalidOperationException($"Header {path} is empty.");

            var recordFields = SplitFields(lines[0]);

            if (recordFields.Length < 2)
                throw new FormatException("WFDB record line must contain at least record name and signal count.");

            var nameToken = recordFields[0];

            if (nameToken.Contains('/'))
            {
                throw new NotSupportedException(
                    "Multi-segment WFDB headers are not supported in this first minimal reader.");
            }

            var signalCount = int.Parse(recordFields[1], CultureInfo.InvariantCulture);
            var sampleRateHz = recordFields.Length >= 3
                ? ParseSampleRate(recordFields[2])
                : 250.0;

            var sampleCount = 0;
            if (recordFields.Length >= 4)
                int.TryParse(recordFields[3], NumberStyles.Integer, CultureInfo.InvariantCulture, out sampleCount);
            
            if (lines.Count < signalCount + 1)
            {
                throw new FormatException(
                    $"Header says {signalCount} signals, but only {lines.Count - 1} signal lines were found.");
            }

            var signals = new List<WFDBSignalInfo>();

            for (var i = 0; i < signalCount; i++)
                signals.Add(ParseSignalLine(lines[i + 1], i));

            return new WFDBHeader(
                Name: nameToken,
                SampleRateHz: sampleRateHz,
                SignalCount: signalCount,
                SampleCount: sampleCount,
                Signals: signals);
        }

        public async Task<WFDBSignal> LoadSignalAsync(WFDBHeader header, int signalIndex)
        {
            
            if (signalIndex < 0 || signalIndex >= header.SignalCount)
                throw new ArgumentOutOfRangeException(nameof(signalIndex));

            return await LoadSignalAsync_Internal(header, signalIndex);

        }

        public async Task<WFDBSignal> LoadSignalAsync(string recordName, int signalIndex)
        {
            var header = await ReadHeaderAsync(recordName);

            if (signalIndex < 0 || signalIndex >= header.SignalCount)
                throw new ArgumentOutOfRangeException(nameof(signalIndex));

            return await LoadSignalAsync_Internal(header, signalIndex);

        }

        private async Task<WFDBSignal> LoadSignalAsync_Internal(WFDBHeader header, int signalIndex)
        {
            var selected = header.Signals[signalIndex];

            if (selected.FileName == "-" || selected.FileName == "~")
                throw new NotSupportedException($"Signal file name '{selected.FileName}' is not supported.");

            var group = header.Signals
                .Select((signal, index) => new { Signal = signal, Index = index })
                .Where(x => x.Signal.FileName == selected.FileName)
                .ToList();

            var signalPositionInFile = group.FindIndex(x => x.Index == signalIndex);
            var signalsInFile = group.Count;

            if (signalPositionInFile < 0)
                throw new InvalidOperationException("Could not locate the selected signal in its signal file group.");

            if (group.Any(x => x.Signal.Format != selected.Format))
                throw new NotSupportedException("Mixed formats within one signal file are not supported.");

            var dataPath = $"{_assetFolder}/{selected.FileName}";
            var bytes = await ReadAllBytesFromPackageAsync(dataPath);

            var samples = selected.Format switch
            {
                16 => DecodeFormat16(
                    bytes,
                    selected.ByteOffset,
                    header.SampleCount,
                    signalsInFile,
                    signalPositionInFile,
                    selected.Gain,
                    selected.Baseline),

                212 => DecodeFormat212(
                    bytes,
                    selected.ByteOffset,
                    header.SampleCount,
                    signalsInFile,
                    signalPositionInFile,
                    selected.Gain,
                    selected.Baseline),

                _ => throw new NotSupportedException(
                    $"WFDB format {selected.Format} is not supported yet. Add a decoder for it.")
            };

            return new WFDBSignal(
                Name: header.Name,
                SelectedSignalName: selected.SignalName,
                Units: selected.Units,
                SampleRateHz: header.SampleRateHz,
                Samples: samples);
        }

        private static async Task<byte[]> ReadAllBytesFromPackageAsync(string path)
        {
            await using var stream = await FileSystem.Current.OpenAppPackageFileAsync(path);
            using var memory = new MemoryStream();

            await stream.CopyToAsync(memory);
            return memory.ToArray();
        }

        private static double[] DecodeFormat16(
            byte[] data,
            int byteOffset,
            int headerSampleCount,
            int signalsInFile,
            int signalPositionInFile,
            double gain,
            int baseline)
        {
            if (byteOffset < 0 || byteOffset >= data.Length)
                byteOffset = 0;

            var availableFlatSamples = (data.Length - byteOffset) / 2;
            var availableFrames = availableFlatSamples / signalsInFile;

            var frameCount = headerSampleCount > 0
                ? Math.Min(headerSampleCount, availableFrames)
                : availableFrames;

            var output = new double[frameCount];
            var safeGain = gain == 0 ? 200.0 : gain;

            for (var frame = 0; frame < frameCount; frame++)
            {
                var flatSampleIndex = frame * signalsInFile + signalPositionInFile;
                var byteIndex = byteOffset + flatSampleIndex * 2;

                var raw = BinaryPrimitives.ReadInt16LittleEndian(data.AsSpan(byteIndex, 2));
                output[frame] = (raw - baseline) / safeGain;
            }

            return output;
        }

        private static double[] DecodeFormat212(
            byte[] data,
            int byteOffset,
            int headerSampleCount,
            int signalsInFile,
            int signalPositionInFile,
            double gain,
            int baseline)
        {
            if (byteOffset < 0 || byteOffset >= data.Length)
                byteOffset = 0;

            var availableBytes = data.Length - byteOffset;
            var availableFlatSamples = (availableBytes / 3) * 2;
            var availableFrames = availableFlatSamples / signalsInFile;

            var frameCount = headerSampleCount > 0
                ? Math.Min(headerSampleCount, availableFrames)
                : availableFrames;

            var output = new double[frameCount];
            var safeGain = gain == 0 ? 200.0 : gain;

            for (var frame = 0; frame < frameCount; frame++)
            {
                var flatSampleIndex = frame * signalsInFile + signalPositionInFile;
                var raw = ReadFormat212Sample(data, byteOffset, flatSampleIndex);

                output[frame] = (raw - baseline) / safeGain;
            }

            return output;
        }

        private static int ReadFormat212Sample(byte[] data, int byteOffset, int flatSampleIndex)
        {
            var tripletIndex = flatSampleIndex / 2;
            var byteIndex = byteOffset + tripletIndex * 3;

            if (byteIndex + 2 >= data.Length)
                return 0;

            var b0 = data[byteIndex];
            var b1 = data[byteIndex + 1];
            var b2 = data[byteIndex + 2];

            int value;

            if (flatSampleIndex % 2 == 0)
            {
                value = b0 | ((b1 & 0x0F) << 8);
            }
            else
            {
                value = ((b1 & 0xF0) >> 4) | (b2 << 4);
            }

            if ((value & 0x800) != 0)
                value -= 0x1000;

            return value;
        }

        private static WFDBSignalInfo ParseSignalLine(string line, int signalIndex)
        {
            var fields = SplitFields(line);

            if (fields.Length < 2)
                throw new FormatException($"Invalid signal line: {line}");

            var fileName = FileBaseName(fields[0]);
            var formatToken = fields[1];

            var format = ParseBaseFormat(formatToken);
            var byteOffset = ParseByteOffset(formatToken);

            var gain = 200.0;
            int? baselineFromGain = null;
            var units = "mV";

            if (fields.Length >= 3)
                ParseGainBaselineUnits(fields[2], out gain, out baselineFromGain, out units);

            var adcZero = 0;
            if (fields.Length >= 5)
                int.TryParse(fields[4], NumberStyles.Integer, CultureInfo.InvariantCulture, out adcZero);

            var baseline = baselineFromGain ?? adcZero;

            var signalName = fields.Length >= 9
                ? string.Join(" ", fields.Skip(8))
                : $"Signal {signalIndex}";

            return new WFDBSignalInfo(
                FileName: fileName,
                Format: format,
                ByteOffset: byteOffset,
                Gain: gain,
                Baseline: baseline,
                Units: units,
                SignalName: signalName);
        }

        private static string[] SplitFields(string line)
        {
            return Regex.Split(line.Trim(), @"\s+")
                .Where(x => x.Length > 0)
                .ToArray();
        }

        private static double ParseSampleRate(string token)
        {
            var main = token.Split('/')[0];
            main = main.Split('(')[0];

            return double.Parse(main, CultureInfo.InvariantCulture);
        }

        private static int ParseBaseFormat(string token)
        {
            var match = Regex.Match(token, @"^\d+");

            if (!match.Success)
                throw new FormatException($"Could not parse WFDB format token '{token}'.");

            return int.Parse(match.Value, CultureInfo.InvariantCulture);
        }

        private static int ParseByteOffset(string token)
        {
            var match = Regex.Match(token, @"\+(\d+)");

            return match.Success
                ? int.Parse(match.Groups[1].Value, CultureInfo.InvariantCulture)
                : 0;
        }

        private static void ParseGainBaselineUnits(
            string token,
            out double gain,
            out int? baseline,
            out string units)
        {
            gain = 200.0;
            baseline = null;
            units = "mV";

            var unitSplit = token.Split('/', 2);
            var gainAndBaseline = unitSplit[0];

            if (unitSplit.Length == 2 && !string.IsNullOrWhiteSpace(unitSplit[1]))
                units = unitSplit[1];

            var parenStart = gainAndBaseline.IndexOf('(');
            var gainText = gainAndBaseline;

            if (parenStart >= 0)
            {
                gainText = gainAndBaseline[..parenStart];

                var parenEnd = gainAndBaseline.IndexOf(')', parenStart + 1);
                if (parenEnd > parenStart)
                {
                    var baselineText = gainAndBaseline[(parenStart + 1)..parenEnd];

                    if (int.TryParse(baselineText, NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsedBaseline))
                        baseline = parsedBaseline;
                }
            }

            if (double.TryParse(gainText, NumberStyles.Float, CultureInfo.InvariantCulture, out var parsedGain)
                && parsedGain != 0)
            {
                gain = parsedGain;
            }
        }

        private static string FileBaseName(string value)
        {
            var normalized = value.Replace('\\', '/');
            var lastSlash = normalized.LastIndexOf('/');

            return lastSlash >= 0
                ? normalized[(lastSlash + 1)..]
                : normalized;
        }
    }
}