# EkgViewerMaui

A simple .NET MAUI ECG/EKG waveform viewer for iOS and Android.

The initial goal is narrow: bundle a small set of WFDB records with the app, select one record, select one signal/lead, and display the waveform as a graph.

This is an early prototype. It currently focuses on waveform loading and rendering, not diagnostic interpretation.

## Project status

Current prototype features:

* .NET MAUI single-project app
* iOS and Android targets
* Bundled WFDB files as app resources
* WFDB header parsing
* WFDB signal data loading
* Basic ECG-style graph rendering with `GraphicsView`
* Record picker
* Signal/lead picker
* Initial support for WFDB format `16`
* Initial support for WFDB format `212`

Not yet implemented:

* Calibrated ECG paper scale
* 25 mm/s paper speed rendering
* 10 mm/mV vertical calibration
* Multi-lead display
* Horizontal scrolling
* Zooming
* Annotation display
* Multi-segment WFDB records
* Remote WFDB downloads
* Clinical interpretation
* 3D heart animation

## Local development stack

Known intended development stack for this project:

| Component            | Version          |
| -------------------- | ---------------- |
| macOS                | Development host |
| Xcode                | 26.3             |
| .NET SDK             | 10.0.203         |
| .NET workload set    | 10.0.203         |
| .NET MAUI workload   | 10.0.20          |
| iOS workload         | 26.2.10233       |
| Android SDK platform | API 36           |
| JDK                  | 21               |
| Editor               | VS Code          |

The project should avoid later Apple workload packs that require Xcode 26.5 unless Xcode is also upgraded.

## Target platforms

The project targets:

```xml
<TargetFrameworks>net10.0-android;net10.0-ios</TargetFrameworks>
```

Mac Catalyst is intentionally omitted during the initial setup to reduce Apple SDK compatibility problems.

## Repository layout

Expected high-level structure:

```text
EkgViewerMaui/
├── App.xaml
├── App.xaml.cs
├── AppShell.xaml
├── AppShell.xaml.cs
├── MainPage.xaml
├── MainPage.xaml.cs
├── EkgViewerMaui.csproj
├── global.json
├── WfdbModels.cs
├── WfdbReader.cs
├── EcgDrawable.cs
├── Resources/
│   ├── AppIcon/
│   ├── Fonts/
│   ├── Images/
│   ├── Raw/
│   │   └── Ekg/
│   │       ├── records.txt
│   │       ├── 100.hea
│   │       └── 100.dat
│   └── Splash/
└── README.md
```

## WFDB resource layout

WFDB records are bundled under:

```text
Resources/Raw/Ekg/
```

Each record should have at least:

```text
record-name.hea
record-name.dat
```

The app also expects a manifest file:

```text
Resources/Raw/Ekg/records.txt
```

Example:

```text
100
101
102
```

Use one record name per line, without the file extension.

For example, this line:

```text
100
```

expects:

```text
Resources/Raw/Ekg/100.hea
Resources/Raw/Ekg/100.dat
```

## Supported WFDB formats

The initial reader supports:

| WFDB format | Description                     |
| ----------- | ------------------------------- |
| `16`        | 16-bit two's-complement samples |
| `212`       | Packed 12-bit samples           |

Unsupported WFDB formats will throw a `NotSupportedException`.

Planned future support may include additional WFDB formats such as `24`, `32`, `80`, `310`, `311`, and `516`, depending on the datasets selected for the app.

## Setup

From the project directory:

```bash
dotnet --version
dotnet workload --version
dotnet workload list
```

Expected local SDK lane:

```text
10.0.203
```

Expected workload set:

```text
10.0.203
```

Install required workloads if needed:

```bash
dotnet workload install maui ios android \
  --version 10.0.203 \
  --skip-manifest-update
```

## Android setup

Expected Android SDK location on macOS:

```text
~/Library/Android/sdk
```

Expected JDK major version:

```text
21
```

Check Java:

```bash
java -version
javac -version
echo "$JAVA_HOME"
```

Build Android:

```bash
dotnet build -f net10.0-android
```

Run on an Android emulator:

```bash
dotnet run -f net10.0-android -p:AdbTarget=-e
```

## iOS setup

Check Xcode:

```bash
xcodebuild -version
```

Expected:

```text
Xcode 26.3
```

Build iOS:

```bash
dotnet build -f net10.0-ios
```

If the build resolves to an iOS pack like this:

```text
Microsoft.iOS.Sdk.net10.0_26.5/26.5.10284
```

that is the wrong iOS workload for Xcode 26.3. The intended compatible iOS pack is:

```text
Microsoft.iOS.Sdk.net10.0_26.2/26.2.10233
```

## Running from VS Code

Open the project folder directly:

```bash
cd ~/Projects/EkgViewerMaui
code .
```

Open a C# file, such as:

```text
MainPage.xaml.cs
```

Then use the VS Code status bar target selector to choose:

```text
Android emulator
```

or:

```text
iOS simulator
```

Then press:

```text
F5
```

If VS Code opens the top command box asking for a debugger, choose:

```text
C#
```

## Build verification

Run these before debugging through VS Code:

```bash
dotnet build -f net10.0-android
dotnet build -f net10.0-ios
```

If command-line builds fail, VS Code debugging will also fail.

## Current rendering behavior

The graph currently:

* Displays one selected signal
* Shows the first several seconds of data
* Auto-scales vertically
* Draws a simple ECG-style background grid
* Does not yet preserve clinical ECG paper calibration

This is appropriate for initial waveform inspection but not yet appropriate for measurement-sensitive ECG review.

## Planned next steps

Near-term development priorities:

1. Add fixed ECG paper speed rendering at 25 mm/s.
2. Add vertical calibration at 10 mm/mV.
3. Add horizontal scrolling.
4. Add a standard 12-lead layout.
5. Add support for more WFDB formats.
6. Add annotation display.
7. Add sample record metadata.
8. Add unit tests for WFDB parsing.
9. Add known-good waveform screenshots for regression testing.

Later project directions:

* Deterministic ECG measurements
* Rule-based interpretation scaffolding
* App-based teaching mode
* Optional AI-assisted explanation layer
* 3D animated cardiac electrical/mechanical model

## Development notes

Keep small sample WFDB records in the repository only if they are intentionally bundled with the app.

Do not commit large external ECG datasets unless their license allows redistribution and they are intentionally part of the app package.

Recommended ignored local dataset folders:

```text
Data/
Datasets/
LocalData/
WFDB/
```

## License

License not yet selected.

Before adding public ECG datasets or sample records to the repository, verify their redistribution and commercial-use terms.
