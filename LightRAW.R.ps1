[CmdletBinding()]
param(
    [string]$FolderPath,
    [switch]$ScanOnly
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

[ComImport]
[Guid("bcc18b79-ba16-442f-80c4-8a59c30c463b")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IShellItemImageFactory {
    void GetImage(NativeSize size, ShellItemImageFactoryFlags flags, out IntPtr phbm);
}

[StructLayout(LayoutKind.Sequential)]
public struct NativeSize {
    public int cx;
    public int cy;
}

[Flags]
public enum ShellItemImageFactoryFlags {
    ResizeToFit = 0x0,
    BiggerSizeOk = 0x1,
    MemoryOnly = 0x2,
    IconOnly = 0x4,
    ThumbnailOnly = 0x8,
    InCacheOnly = 0x10,
    CropToSquare = 0x20,
    WideThumbnails = 0x40,
    IconBackground = 0x80,
    ScaleUp = 0x100
}

public static class NativeShell {
    [DllImport("shell32.dll", CharSet = CharSet.Unicode, PreserveSig = false)]
    public static extern void SHCreateItemFromParsingName(
        string path,
        IntPtr pbc,
        ref Guid riid,
        [MarshalAs(UnmanagedType.Interface)] out IShellItemImageFactory ppv);

    [DllImport("gdi32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool DeleteObject(IntPtr hObject);
}

public static class NativeShellThumbnailProvider {
    public static IntPtr GetHBitmap(string path, int size) {
        Guid iid = typeof(IShellItemImageFactory).GUID;
        IShellItemImageFactory factory;
        NativeShell.SHCreateItemFromParsingName(path, IntPtr.Zero, ref iid, out factory);
        try {
            IntPtr hBitmap;
            factory.GetImage(
                new NativeSize { cx = size, cy = size },
                ShellItemImageFactoryFlags.ThumbnailOnly | ShellItemImageFactoryFlags.BiggerSizeOk,
                out hBitmap
            );
            return hBitmap;
        } finally {
            Marshal.FinalReleaseComObject(factory);
        }
    }
}
"@

$scriptRoot = if ($PSScriptRoot) {
    $PSScriptRoot
} elseif ($PSCommandPath) {
    Split-Path -Parent $PSCommandPath
} else {
    (Get-Location).Path
}

$script:SettingsPath = Join-Path $scriptRoot ".LightRAW.R-settings.json"
$script:QualityPresets = [ordered]@{
    Fast = 1024
    Balanced = 1280
    Sharp = 1600
}
$script:QualityPresetName = "Balanced"
$script:PreviewSourceMode = "JPG"
$script:ShowInfoPanel = $true
$script:ShowPreviewOverlay = $true
$script:PreviewBackgroundGray = 35
$savedFolderPath = $null
$savedPreferredPath = $null
$savedPreviewZoom = 1.0
$savedInfoPanelHeight = 260.0
$savedFailDestinationFolder = $null
$savedShowPreviewOverlay = $true
$savedPreviewBackgroundGray = 35

if (Test-Path -LiteralPath $script:SettingsPath -PathType Leaf) {
    try {
        $settingsData = Get-Content -LiteralPath $script:SettingsPath -Raw | ConvertFrom-Json
        if ($settingsData.QualityPreset -and $script:QualityPresets.Contains([string]$settingsData.QualityPreset)) {
            $script:QualityPresetName = [string]$settingsData.QualityPreset
        }
        if ($settingsData.PreviewSourceMode) {
            $savedPreviewSourceMode = [string]$settingsData.PreviewSourceMode
            if ($savedPreviewSourceMode -in @("JPG", "RAW")) {
                $script:PreviewSourceMode = $savedPreviewSourceMode
            }
        }
        if ($settingsData.LastFolderPath) {
            $savedFolderPath = [string]$settingsData.LastFolderPath
        }
        if ($settingsData.LastPreferredPath) {
            $savedPreferredPath = [string]$settingsData.LastPreferredPath
        }
        if ($settingsData.LastPreviewZoom) {
            $zoomValue = 1.0
            if ([double]::TryParse([string]$settingsData.LastPreviewZoom, [ref]$zoomValue)) {
                $savedPreviewZoom = [Math]::Max(1.0, [Math]::Min(32.0, $zoomValue))
            }
        }
        if ($null -ne $settingsData.ShowInfoPanel) {
            $script:ShowInfoPanel = [bool]$settingsData.ShowInfoPanel
        } elseif ($null -ne $settingsData.ShowExif -or $null -ne $settingsData.ShowAnalysis) {
            $showExifSetting = if ($null -ne $settingsData.ShowExif) { [bool]$settingsData.ShowExif } else { $true }
            $showAnalysisSetting = if ($null -ne $settingsData.ShowAnalysis) { [bool]$settingsData.ShowAnalysis } else { $true }
            $script:ShowInfoPanel = ($showExifSetting -or $showAnalysisSetting)
        }
        if ($settingsData.InfoPanelHeight) {
            $infoPanelHeight = 260.0
            if ([double]::TryParse([string]$settingsData.InfoPanelHeight, [ref]$infoPanelHeight)) {
                $savedInfoPanelHeight = [Math]::Max(140.0, $infoPanelHeight)
            }
        }
        if ($settingsData.FailDestinationFolder) {
            $savedFailDestinationFolder = [string]$settingsData.FailDestinationFolder
        }
        if ($null -ne $settingsData.ShowPreviewOverlay) {
            $savedShowPreviewOverlay = [bool]$settingsData.ShowPreviewOverlay
            $script:ShowPreviewOverlay = $savedShowPreviewOverlay
        }
        if ($null -ne $settingsData.PreviewBackgroundGray) {
            $grayValue = $savedPreviewBackgroundGray
            if ([int]::TryParse([string]$settingsData.PreviewBackgroundGray, [ref]$grayValue)) {
                $savedPreviewBackgroundGray = [Math]::Max(0, [Math]::Min(255, $grayValue))
            }
        }
    } catch {
    }
}

if ([string]::IsNullOrWhiteSpace($FolderPath)) {
    if ($savedFolderPath -and (Test-Path -LiteralPath $savedFolderPath -PathType Container)) {
        $FolderPath = $savedFolderPath
    }
}

$script:PreviewExtensions = @(".jpg", ".jpeg", ".png", ".bmp", ".tif", ".tiff")
$script:RawExtensions = @(".arw", ".cr2", ".cr3", ".nef", ".dng", ".rw2", ".orf", ".raf")
$script:MainPreviewDecodeWidth = [int]$script:QualityPresets[$script:QualityPresetName]
$script:MainRawPreviewSize = [int]$script:QualityPresets[$script:QualityPresetName]
$script:CurrentFolder = $null
$script:Items = @()
$script:Window = $null
$script:PreviewImage = $null
$script:PreviewOverlayBorder = $null
$script:PreviewOverlayText = $null
$script:ShowOverlayCheckBox = $null
$script:EmptyText = $null
$script:DetailsText = $null
$script:SourceText = $null
$script:ExifText = $null
$script:CountText = $null
$script:FileList = $null
$script:PreviewBorder = $null
$script:PreviewHost = $null
$script:PreviewCache = @{}
$script:MetadataCache = @{}
$script:AnalysisCache = @{}
$script:ShellApp = $null
$script:RenderedPreviewKey = $null
$script:PreviewScaleTransform = $null
$script:PreviewTranslateTransform = $null
$script:PreviewZoom = 1.0
$script:IsPanningPreview = $false
$script:IsPreviewMouseDown = $false
$script:PanStartPoint = $null
$script:PreviewMouseDownPoint = $null
$script:PanStartTranslateX = 0.0
$script:PanStartTranslateY = 0.0
$script:PreviewPanThreshold = 4.0
$script:PreviewZoomFactor = 2.0
$script:MaxPreviewZoom = 32.0
$script:StableRawPreviewSize = [int]$script:QualityPresets[$script:QualityPresetName]
$script:AnalysisPreviewSize = 1024
$script:AppStatusText = $null
$script:AnalysisText = $null
$script:QualityFastButton = $null
$script:QualityBalancedButton = $null
$script:QualitySharpButton = $null
$script:PreviewSourceSwitch = $null
$script:ToggleFailButton = $null
$script:MoveFailedButton = $null
$script:FailFolderButton = $null
$script:FailFolderText = $null
$script:ShowInfoCheckBox = $null
$script:InfoPanelBorder = $null
$script:InfoPanelSplitter = $null
$script:InfoSplitterRow = $null
$script:InfoPanelRow = $null
$script:LastPreferredPath = if ($savedPreferredPath) { $savedPreferredPath } else { $null }
$script:PendingRestoreSelectionPath = if ($savedPreferredPath) { $savedPreferredPath } else { $null }
$script:PendingRestorePreviewZoom = [Math]::Max(1.0, $savedPreviewZoom)
$script:ShouldRestoreZoomForSelection = $false
$script:InfoPanelHeight = [Math]::Max(140.0, $savedInfoPanelHeight)
$script:FailDestinationFolder = if ($savedFailDestinationFolder) { $savedFailDestinationFolder } else { $null }
$script:FailedMarksByFolder = @{}
$script:ShowPreviewOverlay = [bool]$savedShowPreviewOverlay
$script:PreviewBackgroundGray = [Math]::Max(0, [Math]::Min(255, [int]$savedPreviewBackgroundGray))
$script:SkipPendingFailPromptOnClose = $false
$script:ThumbnailLoadTimer = $null
$script:ThumbnailLoadIndex = 0
$script:ThumbnailBatchSize = 256
$script:PreviewPreloadTimer = $null
$script:PreviewPreloadQueue = @()
$script:PreviewPreloadRadius = 2
$script:FullResolutionPreviewQueue = @()
$script:FullResolutionPreviewTimer = $null
$script:FullResolutionPreviewPath = $null
$script:FullResolutionRawPath = $null


$script:FunctionsScriptPath = Join-Path $scriptRoot "app\\LightRAW.R.functions.ps1"
$script:UiBootstrapScriptPath = Join-Path $scriptRoot "app\\LightRAW.R.ui.ps1"
$script:MainWindowXamlPath = Join-Path $scriptRoot "ui\\main-window.xaml"

if (-not (Test-Path -LiteralPath $script:FunctionsScriptPath -PathType Leaf)) {
    throw "Missing script file: $script:FunctionsScriptPath"
}

if (-not (Test-Path -LiteralPath $script:UiBootstrapScriptPath -PathType Leaf)) {
    throw "Missing script file: $script:UiBootstrapScriptPath"
}

if (-not (Test-Path -LiteralPath $script:MainWindowXamlPath -PathType Leaf)) {
    throw "Missing XAML file: $script:MainWindowXamlPath"
}

. $script:FunctionsScriptPath

if ($ScanOnly) {
    if ([string]::IsNullOrWhiteSpace($FolderPath)) {
        throw "FolderPath is required when using -ScanOnly."
    }

    $items = Get-GroupedImages -Path $FolderPath
    $items | ForEach-Object {
        [pscustomobject]@{
            BaseName      = $_.BaseName
            PreviewSource = if ($_.PreviewPath) { [System.IO.Path]::GetFileName($_.PreviewPath) } else { "" }
            RawSource     = if ($_.RawPath) { [System.IO.Path]::GetFileName($_.RawPath) } else { "" }
            Extensions    = $_.Extensions
        }
    }
    return
}

[xml]$xaml = Get-Content -LiteralPath $script:MainWindowXamlPath -Raw
. $script:UiBootstrapScriptPath
