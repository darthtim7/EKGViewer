using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace EKGViewer.File
{
    public class WFDBModels
    {
        
    }

    public sealed record WFDBHeader(
        string Name,
        double SampleRateHz,
        int SignalCount,
        int SampleCount,
        IReadOnlyList<WFDBSignalInfo> Signals);

    public sealed record WFDBSignalInfo(
        string FileName,
        int Format,
        int ByteOffset,
        double Gain,
        int Baseline,
        string Units,
        string SignalName);

    public sealed record WFDBSignal(
        string Name,
        string SelectedSignalName,
        string Units,
        double SampleRateHz,
        double[] Samples);

}