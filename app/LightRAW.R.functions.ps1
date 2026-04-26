function Get-GroupedImages {
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    $resolved = [System.IO.Path]::GetFullPath($Path)
    if (-not (Test-Path -LiteralPath $resolved -PathType Container)) {
        throw "Folder not found: $resolved"
    }

    $groups = @{}
    $supportedExtensions = @($script:PreviewExtensions + $script:RawExtensions)

    Get-ChildItem -LiteralPath $resolved -File | ForEach-Object {
        $baseKey = [System.IO.Path]::GetFileNameWithoutExtension($_.Name).ToLowerInvariant()
        $extension = $_.Extension.ToLowerInvariant()

        if ($supportedExtensions -notcontains $extension) {
            return
        }

        if (-not $groups.ContainsKey($baseKey)) {
            $groups[$baseKey] = [ordered]@{
                BaseName = [System.IO.Path]::GetFileNameWithoutExtension($_.Name)
                Files    = @{}
            }
        }

        $groups[$baseKey].Files[$extension] = $_.FullName
    }

    $items = @(
        foreach ($group in ($groups.GetEnumerator() | Sort-Object Key)) {
        $previewPath = $null
        foreach ($extension in $script:PreviewExtensions) {
            if ($group.Value.Files.ContainsKey($extension)) {
                $previewPath = $group.Value.Files[$extension]
                break
            }
        }

        $rawPath = $null
        foreach ($extension in $script:RawExtensions) {
            if ($group.Value.Files.ContainsKey($extension)) {
                $rawPath = $group.Value.Files[$extension]
                break
            }
        }

        $sortedExtensions = $group.Value.Files.Keys | Sort-Object
        $tag = if ($previewPath -and $rawPath) {
            "RAW+PREVIEW"
        } elseif ($previewPath) {
            "PREVIEW ONLY"
        } else {
            "RAW ONLY"
        }

        $primaryName = if ($rawPath) {
            [System.IO.Path]::GetFileName($rawPath)
        } elseif ($previewPath) {
            [System.IO.Path]::GetFileName($previewPath)
        } else {
            $group.Value.BaseName
        }

        $label = if ($rawPath -or $previewPath) {
            $primaryName
        } else {
            $group.Value.BaseName
        }

        $defaultPreferredPath = if ($previewPath) { $previewPath } else { $rawPath }

        [pscustomobject]@{
            GroupKey         = [string]$group.Key
            BaseName        = $group.Value.BaseName
            Label           = "{0} [{1}]" -f $label, $tag
            PrimaryName     = $primaryName
            PreviewPath     = $previewPath
            RawPath         = $rawPath
            PreferredPath   = $defaultPreferredPath
            DefaultPreferredPath = $defaultPreferredPath
            HasQuickPreview = [bool]($previewPath -or $rawPath)
            Extensions      = ($sortedExtensions -join ", ")
            FilePaths       = @($sortedExtensions | ForEach-Object { $group.Value.Files[$_] })
            IsFailed        = $false
        }
    }
    )

    return $items
}

function Ensure-Directory {
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path -PathType Container)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
    }
}

function Get-UniquePathInFolder {
    param(
        [Parameter(Mandatory)]
        [string]$FolderPath,
        [Parameter(Mandatory)]
        [string]$FileName
    )

    $candidatePath = Join-Path $FolderPath $FileName
    if (-not (Test-Path -LiteralPath $candidatePath)) {
        return $candidatePath
    }

    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($FileName)
    $extension = [System.IO.Path]::GetExtension($FileName)

    for ($i = 1; $i -lt 10000; $i++) {
        $candidateName = "{0} ({1}){2}" -f $baseName, $i, $extension
        $candidatePath = Join-Path $FolderPath $candidateName
        if (-not (Test-Path -LiteralPath $candidatePath)) {
            return $candidatePath
        }
    }

    throw "Could not create a unique filename for $FileName in $FolderPath"
}

function New-BitmapImage {
    param(
        [Parameter(Mandatory)]
        [string]$Path,
        [double]$TargetWidth,
        [int]$MinimumDecodeWidth = 128
    )

    $bitmap = New-Object System.Windows.Media.Imaging.BitmapImage
    $bitmap.BeginInit()
    $bitmap.CacheOption = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
    $bitmap.UriSource = [Uri]$Path

    $decodeWidth = [int][Math]::Max($MinimumDecodeWidth, [Math]::Min(2200, [Math]::Round($TargetWidth)))
    if ($decodeWidth -gt 0) {
        $bitmap.DecodePixelWidth = $decodeWidth
    }

    $bitmap.EndInit()
    $bitmap.Freeze()
    return $bitmap
}

function New-FullBitmapImage {
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    $bitmap = New-Object System.Windows.Media.Imaging.BitmapImage
    $bitmap.BeginInit()
    $bitmap.CacheOption = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
    $bitmap.UriSource = [Uri]$Path
    $bitmap.EndInit()
    $bitmap.Freeze()
    return $bitmap
}

function Get-OrientedBitmapImage {
    param(
        [Parameter(Mandatory)]
        [string]$Path,
        [double]$TargetWidth,
        [int]$MinimumDecodeWidth = 128
    )

    $bitmap = New-BitmapImage -Path $Path -TargetWidth $TargetWidth -MinimumDecodeWidth $MinimumDecodeWidth
    $rotation = Get-BitmapRotationFromMetadata -Path $Path
    if ($rotation -eq 0) {
        return $bitmap
    }

    $rotateTransform = New-Object System.Windows.Media.RotateTransform($rotation)
    $rotatedBitmap = New-Object System.Windows.Media.Imaging.TransformedBitmap($bitmap, $rotateTransform)
    $rotatedBitmap.Freeze()
    return $rotatedBitmap
}

function New-ShellThumbnailImage {
    param(
        [Parameter(Mandatory)]
        [string]$Path,
        [double]$TargetWidth,
        [double]$TargetHeight
    )

    $size = [Math]::Max(256, [Math]::Min(2200, [int][Math]::Round([Math]::Max($TargetWidth, $TargetHeight))))
    $hBitmap = [IntPtr]::Zero
    try {
        $hBitmap = [NativeShellThumbnailProvider]::GetHBitmap($Path, $size)

        $bitmap = [System.Windows.Interop.Imaging]::CreateBitmapSourceFromHBitmap(
            $hBitmap,
            [IntPtr]::Zero,
            [System.Windows.Int32Rect]::Empty,
            [System.Windows.Media.Imaging.BitmapSizeOptions]::FromEmptyOptions()
        )
        $bitmap.Freeze()
        return $bitmap
    } finally {
        if ($hBitmap -ne [IntPtr]::Zero) {
            [void][NativeShell]::DeleteObject($hBitmap)
        }
    }
}

function Flush-Ui {
    if (-not $script:Window -or -not $script:Window.Dispatcher) {
        return
    }

    $script:Window.Dispatcher.Invoke([action]{}, [System.Windows.Threading.DispatcherPriority]::Render)
}

function Get-CurrentPreferredPath {
    $selectedItem = if ($script:FileList) { $script:FileList.SelectedItem } else { $null }
    if ($selectedItem -and $selectedItem.PreferredPath) {
        return [string]$selectedItem.PreferredPath
    }

    if ($script:LastPreferredPath) {
        return [string]$script:LastPreferredPath
    }

    return $null
}

function Save-AppSettings {
    try {
        if ($script:InfoPanelRow -and $script:ShowInfoPanel) {
            $currentHeight = $script:InfoPanelRow.Height.Value
            if ($currentHeight -gt 0) {
                $script:InfoPanelHeight = [Math]::Max(140.0, [double]$currentHeight)
            }
        }

        $data = [ordered]@{
            LastFolderPath    = $script:CurrentFolder
            LastPreferredPath = Get-CurrentPreferredPath
            LastPreviewZoom   = [Math]::Round([double]$script:PreviewZoom, 3)
            QualityPreset     = $script:QualityPresetName
            PreviewSourceMode = $script:PreviewSourceMode
            ShowInfoPanel     = [bool]$script:ShowInfoPanel
            InfoPanelHeight   = [Math]::Round([double]$script:InfoPanelHeight, 1)
            FailDestinationFolder = $script:FailDestinationFolder
            ShowPreviewOverlay = [bool]$script:ShowPreviewOverlay
            PreviewBackgroundGray = [int]$script:PreviewBackgroundGray
        }
        $json = $data | ConvertTo-Json -Depth 3
        Set-Content -LiteralPath $script:SettingsPath -Value $json -Encoding UTF8
    } catch {
    }
}

function Update-PreviewBackground {
    $gray = [int][Math]::Round([double]$script:PreviewBackgroundGray)
    $gray = [Math]::Max(0, [Math]::Min(255, $gray))
    $script:PreviewBackgroundGray = $gray

    if ($script:PreviewBackgroundSlider) {
        $sliderValue = [double]$script:PreviewBackgroundSlider.Value
        if ([Math]::Abs($sliderValue - [double]$gray) -gt 0.01) {
            $script:PreviewBackgroundSlider.Value = [double]$gray
        }
    }

    if ($script:PreviewBackgroundValueText) {
        $script:PreviewBackgroundValueText.Text = $gray.ToString()
    }

    if ($script:PreviewBorder) {
        $hex = $gray.ToString("X2")
        $script:PreviewBorder.Background = [System.Windows.Media.Brushes]::Transparent
        $script:PreviewBorder.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#FF{0}{0}{0}" -f $hex)
    }
}

function Get-DefaultFailFolderPath {
    if ($script:CurrentFolder) {
        return (Join-Path $script:CurrentFolder "fail")
    }

    return $null
}

function Get-ActiveFailFolderPath {
    if (-not [string]::IsNullOrWhiteSpace($script:FailDestinationFolder)) {
        return [System.IO.Path]::GetFullPath($script:FailDestinationFolder)
    }

    return Get-DefaultFailFolderPath
}

function Update-FailFolderText {
    if (-not $script:FailFolderText) {
        return
    }

    $activeFailFolder = Get-ActiveFailFolderPath
    if ([string]::IsNullOrWhiteSpace($activeFailFolder)) {
        $script:FailFolderText.Text = "Reject folder: not set"
        return
    }

    $label = if (-not [string]::IsNullOrWhiteSpace($script:FailDestinationFolder)) {
        "Reject folder: {0}" -f $activeFailFolder
    } else {
        "Reject folder: {0} (default)" -f $activeFailFolder
    }
    $script:FailFolderText.Text = $label
}

function Update-PreviewOverlayVisibility {
    if ($script:ShowOverlayCheckBox) {
        $script:ShowOverlayCheckBox.IsChecked = [bool]$script:ShowPreviewOverlay
    }

    if ($script:PreviewOverlayBorder) {
        $hasOverlayContent = $false
        if ($script:PreviewOverlayText) {
            if ($script:PreviewOverlayText.Inlines.Count -gt 0) {
                $hasOverlayContent = $true
            } elseif (-not [string]::IsNullOrWhiteSpace([string]$script:PreviewOverlayText.Text)) {
                $hasOverlayContent = $true
            }
        }

        $shouldShow = $script:ShowPreviewOverlay -and $hasOverlayContent
        $script:PreviewOverlayBorder.Visibility = if ($shouldShow) { "Visible" } else { "Collapsed" }
    }
}

function Clear-HistogramView {
    if ($script:HistogramCanvas) {
        $script:HistogramCanvas.Children.Clear()
    }
    if ($script:HistogramStatusText) {
        $script:HistogramStatusText.Visibility = "Visible"
        $script:HistogramStatusText.Text = "Histogram unavailable"
    }
    if ($script:HistogramSummaryText) {
        $script:HistogramSummaryText.Text = ""
    }
}

function Get-HistogramData {
    param(
        [Parameter(Mandatory)]
        [object]$BitmapSource,
        [int]$Bins = 64
    )

    $targetWidth = 256
    $targetHeight = 256

    $scaleX = $targetWidth / [double]$BitmapSource.PixelWidth
    $scaleY = $targetHeight / [double]$BitmapSource.PixelHeight
    $scale = [Math]::Min(1.0, [Math]::Min($scaleX, $scaleY))

    $thumbnail = if ($scale -lt 0.9999) {
        $tb = New-Object System.Windows.Media.Imaging.TransformedBitmap
        $tb.BeginInit()
        $tb.Source = $BitmapSource
        $tb.Transform = New-Object System.Windows.Media.ScaleTransform($scale, $scale)
        $tb.EndInit()
        $tb
    } else {
        $BitmapSource
    }

    $formatted = New-Object System.Windows.Media.Imaging.FormatConvertedBitmap
    $formatted.BeginInit()
    $formatted.Source = $thumbnail
    $formatted.DestinationFormat = [System.Windows.Media.PixelFormats]::Bgra32
    $formatted.EndInit()

    $width = [int]$formatted.PixelWidth
    $height = [int]$formatted.PixelHeight
    if ($width -le 0 -or $height -le 0) {
        return $null
    }

    $stride = $width * 4
    $bytes = New-Object byte[] ($stride * $height)
    $formatted.CopyPixels($bytes, $stride, 0)

    $hist = New-Object int[] $Bins
    $histR = New-Object int[] $Bins
    $histG = New-Object int[] $Bins
    $histB = New-Object int[] $Bins
    $maxCount = 0
    $step = 256.0 / $Bins

    for ($i = 0; $i -lt $bytes.Length; $i += 4) {
        $b = [int]$bytes[$i]
        $g = [int]$bytes[$i + 1]
        $r = [int]$bytes[$i + 2]
        $luma = [int][Math]::Round((0.2126 * $r) + (0.7152 * $g) + (0.0722 * $b))
        $index = [int][Math]::Floor($luma / $step)
        $rIndex = [int][Math]::Floor($r / $step)
        $gIndex = [int][Math]::Floor($g / $step)
        $bIndex = [int][Math]::Floor($b / $step)

        if ($index -ge $Bins) { $index = $Bins - 1 }
        if ($rIndex -ge $Bins) { $rIndex = $Bins - 1 }
        if ($gIndex -ge $Bins) { $gIndex = $Bins - 1 }
        if ($bIndex -ge $Bins) { $bIndex = $Bins - 1 }

        $hist[$index]++
        $histR[$rIndex]++
        $histG[$gIndex]++
        $histB[$bIndex]++

        if ($hist[$index] -gt $maxCount) { $maxCount = $hist[$index] }
        if ($histR[$rIndex] -gt $maxCount) { $maxCount = $histR[$rIndex] }
        if ($histG[$gIndex] -gt $maxCount) { $maxCount = $histG[$gIndex] }
        if ($histB[$bIndex] -gt $maxCount) { $maxCount = $histB[$bIndex] }
    }

    $totalPixels = $width * $height
    $weightedSum = 0.0
    $shadowPixels = 0
    $highlightPixels = 0
    for ($b = 0; $b -lt $Bins; $b++) {
        $binCenter = (($b + 0.5) * $step)
        $count = $hist[$b]
        $weightedSum += ($binCenter * $count)
        if ($b -lt [int]($Bins * 0.1)) { $shadowPixels += $count }
        if ($b -ge [int]($Bins * 0.9)) { $highlightPixels += $count }
    }

    $mean = if ($totalPixels -gt 0) { $weightedSum / $totalPixels } else { 0.0 }

    return [pscustomobject]@{
        Bins = $hist
        BinsR = $histR
        BinsG = $histG
        BinsB = $histB
        MaxCount = $maxCount
        Mean = $mean
        ShadowPct = if ($totalPixels -gt 0) { (100.0 * $shadowPixels / $totalPixels) } else { 0.0 }
        HighlightPct = if ($totalPixels -gt 0) { (100.0 * $highlightPixels / $totalPixels) } else { 0.0 }
    }
}

function Render-Histogram {
    param(
        [Parameter(Mandatory)]
        [object]$Histogram
    )

    if (-not $script:HistogramCanvas) {
        return
    }

    $script:HistogramCanvas.Children.Clear()
    [System.Windows.Media.RenderOptions]::SetEdgeMode($script:HistogramCanvas, [System.Windows.Media.EdgeMode]::Aliased)
    [System.Windows.Media.TextOptions]::SetTextFormattingMode($script:HistogramCanvas, [System.Windows.Media.TextFormattingMode]::Display)
    [System.Windows.Media.TextOptions]::SetTextRenderingMode($script:HistogramCanvas, [System.Windows.Media.TextRenderingMode]::Aliased)

    $canvasWidth = if ($script:HistogramCanvas.ActualWidth -gt 0) { [double]$script:HistogramCanvas.ActualWidth } else { 420.0 }
    $canvasHeight = if ($script:HistogramCanvas.ActualHeight -gt 0) { [double]$script:HistogramCanvas.ActualHeight } else { 90.0 }
    $dpi = [System.Windows.Media.VisualTreeHelper]::GetDpi($script:HistogramCanvas)
    $dpiScaleX = [Math]::Max(0.0001, [double]$dpi.DpiScaleX)
    $dpiScaleY = [Math]::Max(0.0001, [double]$dpi.DpiScaleY)

    $pixelCanvasWidth = [Math]::Max(1, [int][Math]::Floor($canvasWidth * $dpiScaleX))
    $pixelCanvasHeight = [Math]::Max(1, [int][Math]::Floor($canvasHeight * $dpiScaleY))

    $script:HistogramCanvas.SnapsToDevicePixels = $true
    $script:HistogramCanvas.UseLayoutRounding = $true

    # Render in physical-pixel columns and convert back to DIPs.
    $barWidthDip = 1.0 / $dpiScaleX
    $drawHeight = [Math]::Min(100.0, $canvasHeight)

    $maxCount = [Math]::Max(1, [int]$Histogram.MaxCount)
    $binCount = $Histogram.Bins.Count
    $barWidth = $canvasWidth / [double]$binCount
    $binStride = [Math]::Max(1, [int][Math]::Ceiling(1.0 / [Math]::Max(0.0001, $barWidth * $dpiScaleX)))

    # Keep the bars inside the visible pixel height.
    $pixelDrawHeight = [Math]::Max(1, [int][Math]::Floor([Math]::Min($drawHeight, $canvasHeight) * $dpiScaleY))
    $drawHeight = $pixelDrawHeight / $dpiScaleY

    for ($i = 0; $i -lt $binCount; $i += $binStride) {
        $count = 0
        $countR = 0
        $countG = 0
        $countB = 0
        $end = [Math]::Min($binCount - 1, $i + $binStride - 1)
        for ($j = $i; $j -le $end; $j++) {
            $count += [int]$Histogram.Bins[$j]
            $countR += [int]$Histogram.BinsR[$j]
            $countG += [int]$Histogram.BinsG[$j]
            $countB += [int]$Histogram.BinsB[$j]
        }
        $drawIndex = [int]($i / $binStride)
        $pixelLeft = [Math]::Min($pixelCanvasWidth - 1, [Math]::Max(0, $drawIndex))
        $left = $pixelLeft / $dpiScaleX

        $series = @(
            @{ Count = $count;  Brush = "#B0FFFFFF" },
            @{ Count = $countR; Brush = "#95FF5A5A" },
            @{ Count = $countG; Brush = "#9568FF68" },
            @{ Count = $countB; Brush = "#956FA8FF" }
        )

        foreach ($entry in $series) {
            if ($entry.Count -le 0) { continue }

            $h = [Math]::Max(1.0 / $dpiScaleY, ($entry.Count / [double]$maxCount) * $drawHeight)
            $pixelHeight = [Math]::Max(1, [int][Math]::Round($h * $dpiScaleY))
            $h = $pixelHeight / $dpiScaleY
            $pixelTop = [Math]::Max(0, $pixelDrawHeight - $pixelHeight)
            $top = $pixelTop / $dpiScaleY

            $rect = New-Object System.Windows.Shapes.Rectangle
            $rect.SnapsToDevicePixels = $true
            $rect.UseLayoutRounding = $true
            $rect.Width = $barWidthDip
            $rect.Height = $h
            $rect.Fill = [System.Windows.Media.BrushConverter]::new().ConvertFromString($entry.Brush)
            [System.Windows.Controls.Canvas]::SetLeft($rect, $left)
            [System.Windows.Controls.Canvas]::SetTop($rect, $top)
            [void]$script:HistogramCanvas.Children.Add($rect)
        }
    }

    if ($script:HistogramStatusText) {
        $script:HistogramStatusText.Visibility = "Collapsed"
    }
    if ($script:HistogramSummaryText) {
        $midtonesPct = [Math]::Max(0.0, 100.0 - $Histogram.ShadowPct - $Histogram.HighlightPct)
        $script:HistogramSummaryText.Text = @(
            ("{0,-11}{1,7}" -f "Means:", ("{0:N1}" -f $Histogram.Mean))
            ("{0,-11}{1,7}" -f "Shadows:", ("{0:N1}%" -f $Histogram.ShadowPct))
            ("{0,-11}{1,7}" -f "Midtones:", ("{0:N1}%" -f $midtonesPct))
            ("{0,-11}{1,7}" -f "Highlights:", ("{0:N1}%" -f $Histogram.HighlightPct))
        ) -join "`n"
    }
}

function Get-PreferredPathForMode {
    param(
        [Parameter(Mandatory)]
        [object]$Item,
        [Parameter(Mandatory)]
        [string]$Mode
    )

    switch ($Mode) {
        "RAW" { return $(if ($Item.RawPath) { $Item.RawPath } else { $Item.PreviewPath }) }
        default { return $(if ($Item.PreviewPath) { $Item.PreviewPath } else { $Item.RawPath }) }
    }
}

function Apply-PreviewSourceModeToItems {
    if (-not $script:Items) {
        return
    }

    foreach ($item in $script:Items) {
        $item.PreferredPath = Get-PreferredPathForMode -Item $item -Mode $script:PreviewSourceMode
    }
}

function Get-FailedGroupSet {
    param(
        [Parameter(Mandatory)]
        [string]$FolderPath
    )

    if ($null -eq $script:FailedMarksByFolder) {
        $script:FailedMarksByFolder = @{}
    }

    $normalizedFolder = [System.IO.Path]::GetFullPath($FolderPath)
    if (-not $script:FailedMarksByFolder.ContainsKey($normalizedFolder)) {
        $script:FailedMarksByFolder[$normalizedFolder] = @{}
    }

    return $script:FailedMarksByFolder[$normalizedFolder]
}

function Apply-FailedMarksToItems {
    if (-not $script:CurrentFolder -or -not $script:Items) {
        return
    }

    $failedGroups = Get-FailedGroupSet -FolderPath $script:CurrentFolder
    foreach ($item in $script:Items) {
        $item.IsFailed = $failedGroups.ContainsKey([string]$item.GroupKey)
    }
}

function Refresh-FileListItems {
    if (-not $script:FileList) {
        return
    }

    if ($null -ne $script:FileList.ItemsSource -and [object]::ReferenceEquals($script:FileList.ItemsSource, $script:Items)) {
        try {
            $script:FileList.Items.Refresh()
            if ($script:Window -and $script:Window.Dispatcher) {
                $script:Window.Dispatcher.BeginInvoke([action]{
                    Register-FileListScrollBarViewportUpdater
                    Update-FileListScrollBarViewport
                }, [System.Windows.Threading.DispatcherPriority]::Loaded) | Out-Null
            }
            return
        } catch {
        }
    }

    $script:FileList.ItemsSource = $null
    $script:FileList.ItemsSource = $script:Items
    if ($script:Window -and $script:Window.Dispatcher) {
        $script:Window.Dispatcher.BeginInvoke([action]{
            Register-FileListScrollBarViewportUpdater
            Update-FileListScrollBarViewport
        }, [System.Windows.Threading.DispatcherPriority]::Loaded) | Out-Null
    }
}

function Get-VisualDescendantByType {
    param(
        [Parameter(Mandatory)]
        [System.Windows.DependencyObject]$Root,
        [Parameter(Mandatory)]
        [type]$Type
    )

    $count = [System.Windows.Media.VisualTreeHelper]::GetChildrenCount($Root)
    for ($i = 0; $i -lt $count; $i++) {
        $child = [System.Windows.Media.VisualTreeHelper]::GetChild($Root, $i)
        if ($Type.IsInstanceOfType($child)) {
            return $child
        }

        $match = Get-VisualDescendantByType -Root $child -Type $Type
        if ($match) {
            return $match
        }
    }

    return $null
}

function Update-FileListScrollBarViewport {
    if ($script:UpdatingFileListScrollBarViewport -or -not $script:FileList) {
        return
    }

    $script:UpdatingFileListScrollBarViewport = $true
    try {
        $script:FileList.ApplyTemplate() | Out-Null
        $scrollViewer = $script:FileList.Template.FindName("PART_ScrollViewer", $script:FileList)
        if (-not $scrollViewer) {
            $scrollViewer = Get-VisualDescendantByType -Root $script:FileList -Type ([System.Windows.Controls.ScrollViewer])
        }
        if (-not $scrollViewer) {
            return
        }

        $scrollViewer.ApplyTemplate() | Out-Null
        $scrollBar = $scrollViewer.Template.FindName("PART_VerticalScrollBar", $scrollViewer)
        if (-not $scrollBar) {
            $scrollBar = Get-VisualDescendantByType -Root $scrollViewer -Type ([System.Windows.Controls.Primitives.ScrollBar])
        }
        if (-not $scrollBar -or $scrollBar.Orientation -ne [System.Windows.Controls.Orientation]::Vertical) {
            return
        }

        $scrollBar.ApplyTemplate() | Out-Null
        $track = $scrollBar.Template.FindName("PART_Track", $scrollBar)
        $trackHeight = if ($track -and $track.ActualHeight -gt 0) { [double]$track.ActualHeight } else { [double]$scrollBar.ActualHeight }
        $scrollableHeight = [double]$scrollViewer.ScrollableHeight
        $actualViewportHeight = [double]$scrollViewer.ViewportHeight
        $minimumThumbHeight = 96.0

        if ($scrollableHeight -le 0 -or $trackHeight -le ($minimumThumbHeight + 1.0)) {
            $scrollBar.ViewportSize = $actualViewportHeight
            return
        }

        $minimumViewportForThumb = ($minimumThumbHeight * $scrollableHeight) / ($trackHeight - $minimumThumbHeight)
        $scrollBar.ViewportSize = [Math]::Max($actualViewportHeight, $minimumViewportForThumb)
    } finally {
        $script:UpdatingFileListScrollBarViewport = $false
    }
}

function Register-FileListScrollBarViewportUpdater {
    if ($script:FileListScrollViewerHooked -or -not $script:FileList) {
        return
    }

    $script:FileList.ApplyTemplate() | Out-Null
    $scrollViewer = $script:FileList.Template.FindName("PART_ScrollViewer", $script:FileList)
    if (-not $scrollViewer) {
        $scrollViewer = Get-VisualDescendantByType -Root $script:FileList -Type ([System.Windows.Controls.ScrollViewer])
    }
    if (-not $scrollViewer) {
        return
    }

    $script:FileListScrollViewerHooked = $true
    $scrollViewer.Add_ScrollChanged({
        Update-FileListScrollBarViewport
    })
    $scrollViewer.Add_SizeChanged({
        Update-FileListScrollBarViewport
    })
}

function Get-PendingFailedItems {
    if (-not $script:CurrentFolder) {
        return @()
    }

    $failedGroups = Get-FailedGroupSet -FolderPath $script:CurrentFolder
    $hasFailedGroupMarks = ($failedGroups -and $failedGroups.Count -gt 0)

    if ($script:Items) {
        if ($hasFailedGroupMarks) {
            return @($script:Items | Where-Object {
                $groupKey = [string]$_.GroupKey
                $failedGroups.ContainsKey($groupKey) -and $_.FilePaths -and $_.FilePaths.Count -gt 0
            })
        }

        return @($script:Items | Where-Object { $_.IsFailed -and $_.FilePaths -and $_.FilePaths.Count -gt 0 })
    }

    if ($hasFailedGroupMarks) {
        return @($failedGroups.Keys | ForEach-Object {
            [pscustomobject]@{ GroupKey = [string]$_; IsFailed = $true; FilePaths = @("__pending__") }
        })
    }

    return @()
}

function Show-PendingFailedItemsDialog {
    param(
        [Parameter(Mandatory)]
        [int]$FailedCount
    )

    [xml]$dialogXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Pending Rejected Images"
        Width="460"
        Height="220"
        WindowStartupLocation="CenterOwner"
        ResizeMode="NoResize"
        ShowInTaskbar="False"
        Background="#FF232323"
        Foreground="#FFF5F5F5">
  <Grid Margin="18">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto" />
      <RowDefinition Height="*" />
      <RowDefinition Height="Auto" />
    </Grid.RowDefinitions>

    <TextBlock FontSize="18"
               FontWeight="SemiBold"
               Text="There are rejected images not moved yet." />

    <TextBlock x:Name="MessageText"
               Grid.Row="1"
               Margin="0,14,0,0"
               TextWrapping="Wrap"
               FontSize="14"
               Foreground="#FFD8D8D8" />

    <StackPanel Grid.Row="2"
                Orientation="Horizontal"
                HorizontalAlignment="Right"
                Margin="0,18,0,0">
      <Button x:Name="CloseAnywayButton"
              Content="Close anyway"
              MinWidth="110"
              Padding="12,8"
              Margin="0,0,8,0"
              Background="#FF444444"
              BorderBrush="#FF444444"
              Foreground="White" />
      <Button x:Name="MoveCloseButton"
              Content="Move &amp; close"
              MinWidth="110"
              Padding="12,8"
              Margin="0,0,8,0"
              Background="#FF7A2E2E"
              BorderBrush="#FF7A2E2E"
              Foreground="White" />
      <Button x:Name="CancelCloseButton"
              Content="Cancel"
              MinWidth="90"
              Padding="12,8"
              IsCancel="True"
              Background="#FF2B5D8A"
              BorderBrush="#FF2B5D8A"
              Foreground="White" />
    </StackPanel>
  </Grid>
</Window>
"@

    $reader = New-Object System.Xml.XmlNodeReader $dialogXaml
    $dialog = [Windows.Markup.XamlReader]::Load($reader)
    if ($script:Window -and $script:Window.IsLoaded) {
        $dialog.Owner = $script:Window
    }

    $dialog.Tag = "cancel"
    $messageText = $dialog.FindName("MessageText")
    $closeAnywayButton = $dialog.FindName("CloseAnywayButton")
    $moveCloseButton = $dialog.FindName("MoveCloseButton")
    $cancelCloseButton = $dialog.FindName("CancelCloseButton")

    if ($messageText) {
        $messageText.Text = "There are $FailedCount rejected image group(s) still marked. Choose whether to close without moving, move them first, or cancel and keep the app open."
    }

    if ($closeAnywayButton) {
        $closeAnywayButton.Add_Click({
            $dialog.Tag = "close_anyway"
            $dialog.DialogResult = $true
            $dialog.Close()
        })
    }

    if ($moveCloseButton) {
        $moveCloseButton.Add_Click({
            $dialog.Tag = "move_and_close"
            $dialog.DialogResult = $true
            $dialog.Close()
        })
    }

    if ($cancelCloseButton) {
        $cancelCloseButton.Add_Click({
            $dialog.Tag = "cancel"
            $dialog.DialogResult = $false
            $dialog.Close()
        })
    }

    [void]$dialog.ShowDialog()
    return [string]$dialog.Tag
}

function Apply-QualityPreset {
    param(
        [Parameter(Mandatory)]
        [string]$PresetName
    )

    if (-not $script:QualityPresets.Contains($PresetName)) {
        return
    }

    $script:QualityPresetName = $PresetName
    $script:MainPreviewDecodeWidth = [int]$script:QualityPresets[$PresetName]
    $script:MainRawPreviewSize = [int]$script:QualityPresets[$PresetName]
    $script:StableRawPreviewSize = [int]$script:QualityPresets[$PresetName]
    Save-AppSettings
}

function Update-QualityButtons {
    $isRawMode = ($script:PreviewSourceMode -eq "RAW")

    if ($script:QualityLabelText) {
        $script:QualityLabelText.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString(
            $(if ($isRawMode) { "#FF7A7A7A" } else { "#FFE0E0E0" })
        )
    }

    if (-not $script:QualityPresetComboBox) {
        return
    }

    $script:QualityPresetComboBox.IsEnabled = (-not $isRawMode)
    $script:QualityPresetComboBox.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString(
        $(if ($isRawMode) { "#FF8E8E8E" } else { "#FFFFFFFF" })
    )
    $script:QualityPresetComboBox.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString(
        $(if ($isRawMode) { "#FF1E1E1E" } else { "#FF101010" })
    )
    $script:QualityPresetComboBox.BorderBrush = [System.Windows.Media.BrushConverter]::new().ConvertFromString(
        $(if ($isRawMode) { "#FF303030" } else { "#FF404040" })
    )

    foreach ($item in $script:QualityPresetComboBox.Items) {
        if ($item -and [string]$item.Content -eq $script:QualityPresetName) {
            $script:QualityPresetComboBox.SelectedItem = $item
            break
        }
    }
}

function Update-InfoPanelVisibility {
    if (-not $script:InfoPanelBorder) {
        return
    }

    $script:InfoPanelBorder.Visibility = if ([bool]$script:ShowInfoPanel) { [System.Windows.Visibility]::Visible } else { [System.Windows.Visibility]::Collapsed }

    if ($script:ShowInfoCheckBox) {
        $script:ShowInfoCheckBox.IsChecked = $script:ShowInfoPanel
    }
}

function Set-InfoPopupDefaultPosition {
    return
}

function Get-FilePropertiesText {
    param(
        [string]$Path
    )

    return ""
}

function Get-PreviewCacheKey {
    param(
        [string]$Path,
        [double]$TargetWidth,
        [double]$TargetHeight,
        [string]$Mode
    )

    $bucketWidth = [int]([Math]::Max(1, [Math]::Ceiling($TargetWidth / 128)))
    $bucketHeight = [int]([Math]::Max(1, [Math]::Ceiling($TargetHeight / 128)))
    return "{0}|{1}|{2}|{3}" -f $Mode, $Path.ToLowerInvariant(), $bucketWidth, $bucketHeight
}

function Get-PreviewBitmap {
    param(
        [Parameter(Mandatory)]
        [object]$Item,
        [double]$TargetWidth,
        [double]$TargetHeight
    )

    $preferredPath = if ($Item.PreferredPath) { [string]$Item.PreferredPath } else { $null }
    $isPreferredRaw = $false
    if ($preferredPath -and $Item.RawPath -and ($preferredPath -ieq [string]$Item.RawPath)) {
        $isPreferredRaw = $true
    }

    if ($preferredPath -and -not $isPreferredRaw) {
        $cacheKey = "{0}|{1}|{2}" -f "preview-main", $preferredPath.ToLowerInvariant(), $script:MainPreviewDecodeWidth
        if (-not $script:PreviewCache.ContainsKey($cacheKey)) {
            $script:PreviewCache[$cacheKey] = Get-OrientedBitmapImage -Path $preferredPath -TargetWidth $script:MainPreviewDecodeWidth -MinimumDecodeWidth 256
        }
        return [pscustomobject]@{
            Bitmap = $script:PreviewCache[$cacheKey]
            Source = "Preview JPG/PNG: {0}" -f $preferredPath
            CacheKey = $cacheKey
        }
    }

    if ($preferredPath) {
        $cacheKey = "{0}|{1}|{2}" -f "raw-direct-main", $preferredPath.ToLowerInvariant(), $script:StableRawPreviewSize
        if (-not $script:PreviewCache.ContainsKey($cacheKey)) {
            try {
                $script:PreviewCache[$cacheKey] = Get-OrientedBitmapImage -Path $preferredPath -TargetWidth $script:StableRawPreviewSize -MinimumDecodeWidth 256
            } catch {
                $fallbackKey = "{0}|{1}|{2}" -f "raw-shell-main", $preferredPath.ToLowerInvariant(), $script:StableRawPreviewSize
                if (-not $script:PreviewCache.ContainsKey($fallbackKey)) {
                    $script:PreviewCache[$fallbackKey] = New-ShellThumbnailImage -Path $preferredPath -TargetWidth $script:StableRawPreviewSize -TargetHeight $script:StableRawPreviewSize
                }
                return [pscustomobject]@{
                    Bitmap = $script:PreviewCache[$fallbackKey]
                    Source = "Preview RAW thumbnail fallback: {0}" -f $preferredPath
                    CacheKey = $fallbackKey
                }
            }
        }
        return [pscustomobject]@{
            Bitmap = $script:PreviewCache[$cacheKey]
            Source = "Preview RAW direct: {0}" -f $preferredPath
            CacheKey = $cacheKey
        }
    }

    throw "No preview source available."
}

function Get-AnalysisBitmap {
    param(
        [Parameter(Mandatory)]
        [object]$Item
    )

    if ($Item.PreviewPath) {
        $cacheKey = "{0}|{1}|{2}" -f "analysis-preview", $Item.PreviewPath.ToLowerInvariant(), $script:AnalysisPreviewSize
        if (-not $script:PreviewCache.ContainsKey($cacheKey)) {
            $script:PreviewCache[$cacheKey] = New-BitmapImage -Path $Item.PreviewPath -TargetWidth $script:AnalysisPreviewSize -MinimumDecodeWidth 256
        }

        return [pscustomobject]@{
            Bitmap = $script:PreviewCache[$cacheKey]
            CacheKey = $cacheKey
        }
    }

    if ($Item.RawPath) {
        $cacheKey = "{0}|{1}|{2}" -f "analysis-raw", $Item.RawPath.ToLowerInvariant(), $script:AnalysisPreviewSize
        if (-not $script:PreviewCache.ContainsKey($cacheKey)) {
            try {
                $script:PreviewCache[$cacheKey] = New-ShellThumbnailImage -Path $Item.RawPath -TargetWidth $script:AnalysisPreviewSize -TargetHeight $script:AnalysisPreviewSize
            } catch {
                $fallbackKey = "{0}|{1}|{2}" -f "analysis-raw-direct", $Item.RawPath.ToLowerInvariant(), $script:AnalysisPreviewSize
                if (-not $script:PreviewCache.ContainsKey($fallbackKey)) {
                    $script:PreviewCache[$fallbackKey] = New-BitmapImage -Path $Item.RawPath -TargetWidth $script:AnalysisPreviewSize -MinimumDecodeWidth 256
                }

                return [pscustomobject]@{
                    Bitmap = $script:PreviewCache[$fallbackKey]
                    CacheKey = $fallbackKey
                }
            }
        }

        return [pscustomobject]@{
            Bitmap = $script:PreviewCache[$cacheKey]
            CacheKey = $cacheKey
        }
    }

    return $null
}

function Get-100PercentPreviewBitmap {
    param(
        [Parameter(Mandatory)]
        [object]$Item
    )

    $preferredPath = if ($Item.PreferredPath) { [string]$Item.PreferredPath } else { $null }
    $isPreferredRaw = $false
    if ($preferredPath -and $Item.RawPath -and ($preferredPath -ieq [string]$Item.RawPath)) {
        $isPreferredRaw = $true
    }

    if ($preferredPath -and -not $isPreferredRaw) {
        $cacheKey = "{0}|{1}" -f "preview-native-full", $preferredPath.ToLowerInvariant()
        if (-not $script:PreviewCache.ContainsKey($cacheKey)) {
            $script:PreviewCache[$cacheKey] = Get-OrientedFullBitmapImage -Path $preferredPath
        }

        return [pscustomobject]@{
            Bitmap = $script:PreviewCache[$cacheKey]
            Source = "Preview JPG/PNG 100%: {0}" -f $preferredPath
            CacheKey = $cacheKey
        }
    }

    if ($preferredPath) {
        $cacheKey = "{0}|{1}" -f "raw-native-full", $preferredPath.ToLowerInvariant()
        if (-not $script:PreviewCache.ContainsKey($cacheKey)) {
            $script:PreviewCache[$cacheKey] = Get-OrientedFullBitmapImage -Path $preferredPath
        }

        return [pscustomobject]@{
            Bitmap = $script:PreviewCache[$cacheKey]
            Source = "Preview RAW 100%: {0}" -f $preferredPath
            CacheKey = $cacheKey
        }
    }

    return Get-PreviewBitmap -Item $Item -TargetWidth $script:MainPreviewDecodeWidth -TargetHeight $script:MainRawPreviewSize
}

function Get-ThumbnailBitmap {
    param(
        [Parameter(Mandatory)]
        [object]$Item
    )

    if (-not $Item.HasQuickPreview -or -not $Item.PreviewPath) {
        return $null
    }

    $targetWidth = 120
    $targetHeight = 84

    if ($Item.PreviewPath) {
        $cacheKey = Get-PreviewCacheKey -Path $Item.PreviewPath -TargetWidth $targetWidth -TargetHeight $targetHeight -Mode "thumb-preview-shell"
        if (-not $script:PreviewCache.ContainsKey($cacheKey)) {
            try {
                $script:PreviewCache[$cacheKey] = New-ShellThumbnailImage -Path $Item.PreviewPath -TargetWidth $targetWidth -TargetHeight $targetHeight
            } catch {
                $fallbackKey = Get-PreviewCacheKey -Path $Item.PreviewPath -TargetWidth $targetWidth -TargetHeight $targetHeight -Mode "thumb-preview-direct"
                if (-not $script:PreviewCache.ContainsKey($fallbackKey)) {
                    $script:PreviewCache[$fallbackKey] = New-BitmapImage -Path $Item.PreviewPath -TargetWidth $targetWidth -MinimumDecodeWidth 128
                }
                return $script:PreviewCache[$fallbackKey]
            }
        }
        return $script:PreviewCache[$cacheKey]
    }

    return $null
}

function Stop-ThumbnailLoading {
    if ($script:ThumbnailLoadTimer) {
        $script:ThumbnailLoadTimer.Stop()
    }
}

function Update-ThumbnailLoadingStatus {
    $total = if ($script:Items) { $script:Items.Count } else { 0 }
    if ($total -le 0) {
        Set-AppStatus -Status "Ready"
        return
    }

    $completed = [Math]::Min($script:ThumbnailLoadIndex, $total)
    if ($completed -ge $total) {
        Set-AppStatus -Status "Ready"
    } else {
        Set-AppStatus -Status ("Building thumbnails... {0}/{1}" -f $completed, $total)
    }
}

function Initialize-ThumbnailState {
    if (-not $script:Items) {
        return
    }

    foreach ($item in $script:Items) {
        $item | Add-Member -NotePropertyName ThumbnailBitmap -NotePropertyValue $null -Force
    }
}

function Load-NextThumbnailBatch {
    if (-not $script:Items) {
        Stop-ThumbnailLoading
        return
    }

    $total = $script:Items.Count
    if ($script:ThumbnailLoadIndex -ge $total) {
        Stop-ThumbnailLoading
        Update-ThumbnailLoadingStatus
        return
    }

    $batchEnd = [Math]::Min($script:ThumbnailLoadIndex + [Math]::Max(1, $script:ThumbnailBatchSize), $total)
    for ($index = $script:ThumbnailLoadIndex; $index -lt $batchEnd; $index++) {
        $item = $script:Items[$index]
        if ($null -eq $item.ThumbnailBitmap) {
            $thumbnail = Get-ThumbnailBitmap -Item $item
            $item | Add-Member -NotePropertyName ThumbnailBitmap -NotePropertyValue $thumbnail -Force
        }
    }

    $script:ThumbnailLoadIndex = $batchEnd
    if ($script:FileList) {
        $script:FileList.Items.Refresh()
    }

    Update-ThumbnailLoadingStatus

    if ($script:ThumbnailLoadIndex -ge $total) {
        Stop-ThumbnailLoading
    }
}

function Start-ThumbnailLoading {
    Stop-ThumbnailLoading
    Initialize-ThumbnailState
    $script:ThumbnailLoadIndex = 0

    if (-not $script:Items -or $script:Items.Count -le 0) {
        Set-AppStatus -Status "Ready"
        return
    }

    if (-not $script:ThumbnailLoadTimer) {
        $script:ThumbnailLoadTimer = New-Object System.Windows.Threading.DispatcherTimer
        $script:ThumbnailLoadTimer.Interval = [TimeSpan]::FromMilliseconds(8)
        $script:ThumbnailLoadTimer.Add_Tick({
            Load-NextThumbnailBatch
        })
    }

    Update-ThumbnailLoadingStatus
    Load-NextThumbnailBatch
    if ($script:ThumbnailLoadIndex -lt $script:Items.Count) {
        $script:ThumbnailLoadTimer.Start()
    }
}

function Stop-PreviewPreloading {
    if ($script:PreviewPreloadTimer) {
        $script:PreviewPreloadTimer.Stop()
    }
    $script:PreviewPreloadQueue = @()
}

function Warm-PreviewCacheForItem {
    param(
        [Parameter(Mandatory)]
        [object]$Item
    )

    if (-not $Item -or -not $Item.HasQuickPreview -or -not $Item.PreferredPath) {
        return
    }

    try {
        [void](Get-PreviewBitmap -Item $Item -TargetWidth $script:MainPreviewDecodeWidth -TargetHeight $script:MainRawPreviewSize)
    } catch {
    }
}

function Warm-100PercentPreviewCacheForItem {
    param(
        [Parameter(Mandatory)]
        [object]$Item
    )

    if (-not $Item -or -not $Item.HasQuickPreview -or -not $Item.PreferredPath) {
        return
    }

    try {
        [void](Get-100PercentPreviewBitmap -Item $Item)
    } catch {
    }
}

function Queue-100PercentPreviewPreloadForSelection {
    param(
        [Parameter(Mandatory)]
        [object]$Item
    )

    if (-not $Item -or -not $Item.HasQuickPreview -or -not $Item.PreferredPath) {
        return
    }

    # Stop any previous timer
    if ($script:FullResolutionPreviewTimer) {
        $script:FullResolutionPreviewTimer.Stop()
    }
    
    # Store paths in script scope for the event handler to use.
    # Prefer the currently selected source (PreferredPath) so RAW mode starts
    # loading full-size RAW immediately on selection.
    $preferredPath = if ($Item.PreferredPath) { [string]$Item.PreferredPath } else { $null }
    $isPreferredRaw = $false
    if ($preferredPath -and $Item.RawPath -and ($preferredPath -ieq [string]$Item.RawPath)) {
        $isPreferredRaw = $true
    }

    if ($isPreferredRaw) {
        $script:FullResolutionPreviewPath = $null
        $script:FullResolutionRawPath = $preferredPath
    } else {
        $script:FullResolutionPreviewPath = if ($preferredPath) { $preferredPath } else { $Item.PreviewPath }
        $script:FullResolutionRawPath = $Item.RawPath
    }
    
    # Use a dispatcher timer to load the 100% preview asynchronously
    $script:FullResolutionPreviewTimer = New-Object System.Windows.Threading.DispatcherTimer
    $script:FullResolutionPreviewTimer.Interval = [TimeSpan]::FromMilliseconds(50)
    $script:FullResolutionPreviewTimer.Add_Tick({
        $script:FullResolutionPreviewTimer.Stop()
        
        try {
            # Load the preview directly using script-scoped paths
            if ($script:FullResolutionPreviewPath) {
                $cacheKey = "{0}|{1}" -f "preview-native-full", $script:FullResolutionPreviewPath.ToLowerInvariant()
                if (-not $script:PreviewCache.ContainsKey($cacheKey)) {
                    $script:PreviewCache[$cacheKey] = Get-OrientedFullBitmapImage -Path $script:FullResolutionPreviewPath
                }
            } elseif ($script:FullResolutionRawPath) {
                $cacheKey = "{0}|{1}" -f "raw-native-full", $script:FullResolutionRawPath.ToLowerInvariant()
                if (-not $script:PreviewCache.ContainsKey($cacheKey)) {
                    $script:PreviewCache[$cacheKey] = Get-OrientedFullBitmapImage -Path $script:FullResolutionRawPath
                }
            }
        } catch {
        }
    })
    $script:FullResolutionPreviewTimer.Start()
}

function Queue-PreviewPreloadForSelection {
    if (-not $script:Items -or -not $script:FileList) {
        return
    }

    Stop-PreviewPreloading

    $selectedIndex = $script:FileList.SelectedIndex
    if ($selectedIndex -lt 0 -or $selectedIndex -ge $script:Items.Count) {
        return
    }

    $queuedIndexes = New-Object 'System.Collections.Generic.HashSet[int]'
    $queue = New-Object System.Collections.Generic.List[object]

    for ($offset = 1; $offset -le [Math]::Max(1, $script:PreviewPreloadRadius); $offset++) {
        $nextIndex = $selectedIndex + $offset
        if ($nextIndex -lt $script:Items.Count -and $queuedIndexes.Add($nextIndex)) {
            [void]$queue.Add($script:Items[$nextIndex])
        }

        $prevIndex = $selectedIndex - $offset
        if ($prevIndex -ge 0 -and $queuedIndexes.Add($prevIndex)) {
            [void]$queue.Add($script:Items[$prevIndex])
        }
    }

    $script:PreviewPreloadQueue = @($queue.ToArray())
    if ($script:PreviewPreloadQueue.Count -le 0) {
        return
    }

    if (-not $script:PreviewPreloadTimer) {
        $script:PreviewPreloadTimer = New-Object System.Windows.Threading.DispatcherTimer
        $script:PreviewPreloadTimer.Interval = [TimeSpan]::FromMilliseconds(35)
        $script:PreviewPreloadTimer.Add_Tick({
            if (-not $script:PreviewPreloadQueue -or $script:PreviewPreloadQueue.Count -le 0) {
                Stop-PreviewPreloading
                return
            }

            $nextItem = $script:PreviewPreloadQueue[0]
            if ($script:PreviewPreloadQueue.Count -gt 1) {
                $script:PreviewPreloadQueue = @($script:PreviewPreloadQueue[1..($script:PreviewPreloadQueue.Count - 1)])
            } else {
                $script:PreviewPreloadQueue = @()
            }

            Warm-PreviewCacheForItem -Item $nextItem

            if (-not $script:PreviewPreloadQueue -or $script:PreviewPreloadQueue.Count -le 0) {
                Stop-PreviewPreloading
            }
        })
    }

    $script:PreviewPreloadTimer.Start()
}

function Get-ShellExtendedProperty {
    param(
        [Parameter(Mandatory)]
        [string]$Path,
        [Parameter(Mandatory)]
        [string]$PropertyName
    )

    $directoryPath = Split-Path -Parent $Path
    $fileName = Split-Path -Leaf $Path

    if (-not (Test-Path -LiteralPath $directoryPath -PathType Container)) {
        return $null
    }

    if (-not $script:ShellApp) {
        $script:ShellApp = New-Object -ComObject Shell.Application
    }

    $folder = $script:ShellApp.Namespace($directoryPath)
    if ($null -eq $folder) {
        return $null
    }

    $item = $folder.ParseName($fileName)
    if ($null -eq $item) {
        return $null
    }

    try {
        return $item.ExtendedProperty($PropertyName)
    } catch {
        return $null
    }
}

function Get-BitmapRotationFromMetadata {
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    $orientation = Get-ShellExtendedProperty -Path $Path -PropertyName "System.Photo.Orientation"
    if ($null -eq $orientation) {
        return 0
    }

    $orientationText = [string]$orientation
    $numericOrientation = 0
    [void][int]::TryParse($orientationText, [ref]$numericOrientation)

    if ($numericOrientation -eq 6 -or $orientationText -match "90") {
        return 90
    }

    if ($numericOrientation -eq 3 -or $orientationText -match "180") {
        return 180
    }

    if ($numericOrientation -eq 8 -or $orientationText -match "270") {
        return 270
    }

    return 0
}

function Get-OrientedFullBitmapImage {
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    $bitmap = New-FullBitmapImage -Path $Path
    $rotation = Get-BitmapRotationFromMetadata -Path $Path
    if ($rotation -eq 0) {
        return $bitmap
    }

    $rotateTransform = New-Object System.Windows.Media.RotateTransform($rotation)
    $rotatedBitmap = New-Object System.Windows.Media.Imaging.TransformedBitmap($bitmap, $rotateTransform)
    $rotatedBitmap.Freeze()
    return $rotatedBitmap
}

function Format-ExposureTime {
    param([object]$Value)

    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) {
        return $null
    }

    $seconds = 0.0
    if (-not [double]::TryParse([string]$Value, [ref]$seconds)) {
        return [string]$Value
    }

    if ($seconds -le 0) {
        return $null
    }

    if ($seconds -ge 1) {
        return "{0:0.###} s" -f $seconds
    }

    $denominator = [Math]::Round(1.0 / $seconds)
    if ($denominator -gt 0) {
        return "1/{0} s" -f $denominator
    }

    return "{0:0.####} s" -f $seconds
}

function Format-ExifNumber {
    param(
        [object]$Value,
        [string]$Prefix = "",
        [string]$Suffix = ""
    )

    if ($null -eq $Value) {
        return $null
    }

    $text = [string]$Value
    if ([string]::IsNullOrWhiteSpace($text) -or $text -eq "0" -or $text -eq "----") {
        return $null
    }

    return "{0}{1}{2}" -f $Prefix, $text, $Suffix
}

function Add-ExifSummaryLine {
    param(
        [Parameter(Mandatory)]
        [System.Collections.Generic.List[string]]$Lines,
        [Parameter(Mandatory)]
        [string]$Label,
        [object]$Value
    )

    if ($null -eq $Value) {
        return
    }

    $text = [string]$Value
    if ([string]::IsNullOrWhiteSpace($text) -or $text -eq "----") {
        return
    }

    $Lines.Add(("{0}: {1}" -f $Label, $text))
}

function Get-ExifDetails {
    param(
        [Parameter(Mandatory)]
        [object]$Item
    )

    $metadataSource = if ($Item.PreviewPath) { $Item.PreviewPath } else { $Item.RawPath }
    if ([string]::IsNullOrWhiteSpace($metadataSource)) {
        return [pscustomobject]@{
            Summary = ""
            Overlay = ""
        }
    }

    $cacheKey = $metadataSource.ToLowerInvariant()
    if ($script:MetadataCache.ContainsKey($cacheKey)) {
        return $script:MetadataCache[$cacheKey]
    }

    $camera = Get-ShellExtendedProperty -Path $metadataSource -PropertyName "System.Photo.CameraModel"
    $cameraMake = Get-ShellExtendedProperty -Path $metadataSource -PropertyName "System.Photo.CameraManufacturer"
    $dateTaken = Get-ShellExtendedProperty -Path $metadataSource -PropertyName "System.Photo.DateTaken"
    $exposure = Get-ShellExtendedProperty -Path $metadataSource -PropertyName "System.Photo.ExposureTime"
    $iso = Get-ShellExtendedProperty -Path $metadataSource -PropertyName "System.Photo.ISOSpeed"
    $fNumber = Get-ShellExtendedProperty -Path $metadataSource -PropertyName "System.Photo.FNumber"
    $focalLength = Get-ShellExtendedProperty -Path $metadataSource -PropertyName "System.Photo.FocalLength"
    $lens = Get-ShellExtendedProperty -Path $metadataSource -PropertyName "System.Photo.LensModel"
    $flash = Get-ShellExtendedProperty -Path $metadataSource -PropertyName "System.Photo.Flash"
    $whiteBalance = Get-ShellExtendedProperty -Path $metadataSource -PropertyName "System.Photo.WhiteBalance"
    $meteringMode = Get-ShellExtendedProperty -Path $metadataSource -PropertyName "System.Photo.MeteringMode"
    $programMode = Get-ShellExtendedProperty -Path $metadataSource -PropertyName "System.Photo.ProgramMode"
    $exposureBias = Get-ShellExtendedProperty -Path $metadataSource -PropertyName "System.Photo.ExposureBias"
    $maxAperture = Get-ShellExtendedProperty -Path $metadataSource -PropertyName "System.Photo.MaxAperture"
    $orientation = Get-ShellExtendedProperty -Path $metadataSource -PropertyName "System.Photo.Orientation"
    $software = Get-ShellExtendedProperty -Path $metadataSource -PropertyName "System.ApplicationName"
    $width = Get-ShellExtendedProperty -Path $metadataSource -PropertyName "System.Image.HorizontalSize"
    $height = Get-ShellExtendedProperty -Path $metadataSource -PropertyName "System.Image.VerticalSize"

    $summaryLines = New-Object System.Collections.Generic.List[string]
    $overlayLines = New-Object System.Collections.Generic.List[string]
    $sourceLabel = if ($Item.PreviewPath) { "JPG sidecar" } else { "RAW" }
    $sourceName = Split-Path -Leaf $metadataSource
    $summaryLines.Add(("EXIF source: {0} ({1})" -f $sourceLabel, $sourceName))

    if ($dateTaken) {
        try {
            $summaryLines.Add(("Taken: {0:yyyy-MM-dd HH:mm:ss}" -f [datetime]$dateTaken))
        } catch {
            $summaryLines.Add("Taken: $dateTaken")
        }
    }

    Add-ExifSummaryLine -Lines $summaryLines -Label "Camera" -Value $camera
    Add-ExifSummaryLine -Lines $summaryLines -Label "Maker" -Value $cameraMake

    if ($dateTaken) {
        try {
            $overlayLines.Add(("Taken: {0:yyyy-MM-dd HH:mm:ss}" -f [datetime]$dateTaken))
        } catch {
            $overlayLines.Add("Taken: $dateTaken")
        }
    }

    $cameraText = if (-not [string]::IsNullOrWhiteSpace([string]$camera)) { [string]$camera } else { $null }
    $lensText = if (-not [string]::IsNullOrWhiteSpace([string]$lens)) { [string]$lens } else { $null }
    if ($cameraText -and $lensText) {
        $overlayLines.Add(("Camera: {0} | Lens: {1}" -f $cameraText, $lensText))
    } elseif ($cameraText) {
        $overlayLines.Add(("Camera: {0}" -f $cameraText))
    } elseif ($lensText) {
        $overlayLines.Add(("Lens: {0}" -f $lensText))
    }

    $settings = @()
    $formattedExposure = Format-ExposureTime -Value $exposure
    if ($formattedExposure) { $settings += $formattedExposure }

    $formattedFNumber = Format-ExifNumber -Value $fNumber -Prefix "f/"
    if ($formattedFNumber) { $settings += $formattedFNumber }

    $formattedIso = Format-ExifNumber -Value $iso -Prefix "ISO "
    if ($formattedIso) { $settings += $formattedIso }

    $formattedFocalLength = Format-ExifNumber -Value $focalLength -Suffix " mm"
    if ($formattedFocalLength) { $settings += $formattedFocalLength }

    if ($settings.Count -gt 0) {
        $overlayLines.Add(("Settings: {0}" -f ($settings -join " | ")))
        $summaryLines.Add(("Settings: {0}" -f ($settings -join " | ")))
    }

    if ($width -and $height) {
        $resolutionText = "{0} x {1}" -f $width, $height
        $overlayLines.Add(("Size: {0}" -f $resolutionText))
    }

    Add-ExifSummaryLine -Lines $summaryLines -Label "Lens" -Value $lens
    Add-ExifSummaryLine -Lines $summaryLines -Label "Exposure" -Value $formattedExposure
    Add-ExifSummaryLine -Lines $summaryLines -Label "Aperture" -Value $formattedFNumber
    Add-ExifSummaryLine -Lines $summaryLines -Label "ISO" -Value $formattedIso
    Add-ExifSummaryLine -Lines $summaryLines -Label "Focal length" -Value $formattedFocalLength
    Add-ExifSummaryLine -Lines $summaryLines -Label "Flash" -Value $flash
    Add-ExifSummaryLine -Lines $summaryLines -Label "White balance" -Value $whiteBalance
    Add-ExifSummaryLine -Lines $summaryLines -Label "Metering" -Value $meteringMode
    Add-ExifSummaryLine -Lines $summaryLines -Label "Program mode" -Value $programMode
    Add-ExifSummaryLine -Lines $summaryLines -Label "Exposure bias" -Value $exposureBias
    Add-ExifSummaryLine -Lines $summaryLines -Label "Max aperture" -Value $maxAperture
    Add-ExifSummaryLine -Lines $summaryLines -Label "Orientation" -Value $orientation

    if ($width -and $height) {
        $summaryLines.Add(("Size: {0} x {1}" -f $width, $height))
    }

    Add-ExifSummaryLine -Lines $summaryLines -Label "Software" -Value $software

    if ($overlayLines.Count -eq 0 -and $summaryLines.Count -le 1) {
        $summaryLines.Add("No EXIF metadata found.")
    }

    $details = [pscustomobject]@{
        Summary = ($summaryLines -join "`n")
        Overlay = ($overlayLines -join "`n")
    }
    $script:MetadataCache[$cacheKey] = $details
    return $details
}

function Get-ExifSummary {
    param(
        [Parameter(Mandatory)]
        [object]$Item
    )

    return (Get-ExifDetails -Item $Item).Summary
}

function Get-ExifOverlayText {
    param(
        [Parameter(Mandatory)]
        [object]$Item
    )

    return (Get-ExifDetails -Item $Item).Overlay
}

function Set-StatusText {
    param(
        [string]$Details,
        [string]$Source,
        [string]$Exif = "",
        [string]$Analysis = "",
        [string]$Overlay = ""
    )

    $script:DetailsText.Text = $Details
    $script:SourceText.Text = $Source
    $script:ExifText.Text = $Exif
    if ($script:PreviewOverlayText) {
        $script:PreviewOverlayText.Inlines.Clear()

        if (-not [string]::IsNullOrWhiteSpace($Overlay)) {
            $lines = $Overlay -split "`r?`n"
            for ($i = 0; $i -lt $lines.Count; $i++) {
                $line = [string]$lines[$i]
                if ([string]::IsNullOrWhiteSpace($line)) {
                    $script:PreviewOverlayText.Inlines.Add([System.Windows.Documents.Run]::new("")) | Out-Null
                } else {
                    $separatorIndex = $line.IndexOf(":")
                    if ($separatorIndex -gt 0) {
                        if ($line -match "\|\s*Lens:") {
                            $lensMatch = [regex]::Match($line, "\|\s*Lens:")
                            $cameraMatch = [regex]::Match($line, "^\s*Camera:")
                            $cameraLabelPart = $cameraMatch.Value
                            $cameraValuePart = $line.Substring($cameraMatch.Length, $lensMatch.Index + 1 - $cameraMatch.Length)
                            $lensLabelPart = $lensMatch.Value.Substring(1)
                            $lensValuePart = $line.Substring($lensMatch.Index + $lensMatch.Length)

                            $cameraLabelRun = [System.Windows.Documents.Run]::new($cameraLabelPart)
                            $cameraLabelRun.Foreground = [System.Windows.Media.Brushes]::Gold
                            $script:PreviewOverlayText.Inlines.Add($cameraLabelRun) | Out-Null

                            $cameraValueRun = [System.Windows.Documents.Run]::new($cameraValuePart)
                            $cameraValueRun.Foreground = [System.Windows.Media.Brushes]::White
                            $script:PreviewOverlayText.Inlines.Add($cameraValueRun) | Out-Null

                            $lensLabelRun = [System.Windows.Documents.Run]::new($lensLabelPart)
                            $lensLabelRun.Foreground = [System.Windows.Media.Brushes]::Gold
                            $script:PreviewOverlayText.Inlines.Add($lensLabelRun) | Out-Null

                            $lensValueRun = [System.Windows.Documents.Run]::new($lensValuePart)
                            $lensValueRun.Foreground = [System.Windows.Media.Brushes]::White
                            $script:PreviewOverlayText.Inlines.Add($lensValueRun) | Out-Null
                        } else {
                            $labelPart = $line.Substring(0, $separatorIndex + 1)
                            $valuePart = $line.Substring($separatorIndex + 1)

                            $labelRun = [System.Windows.Documents.Run]::new($labelPart)
                            $labelRun.Foreground = [System.Windows.Media.Brushes]::Gold
                            $script:PreviewOverlayText.Inlines.Add($labelRun) | Out-Null

                            $valueRun = [System.Windows.Documents.Run]::new($valuePart)
                            $valueRun.Foreground = [System.Windows.Media.Brushes]::White
                            $script:PreviewOverlayText.Inlines.Add($valueRun) | Out-Null
                        }
                    } else {
                        $textRun = [System.Windows.Documents.Run]::new($line)
                        $textRun.Foreground = [System.Windows.Media.Brushes]::White
                        $script:PreviewOverlayText.Inlines.Add($textRun) | Out-Null
                    }
                }

                if ($i -lt ($lines.Count - 1)) {
                    $script:PreviewOverlayText.Inlines.Add([System.Windows.Documents.LineBreak]::new()) | Out-Null
                }
            }
        }
    }
    Update-PreviewOverlayVisibility
    if ($script:AnalysisText) {
        $script:AnalysisText.Text = $Analysis
        $script:AnalysisText.Visibility = if ([string]::IsNullOrWhiteSpace($Analysis)) { "Collapsed" } else { "Visible" }
    }

    if ($script:InfoPropertiesText) {
        $script:InfoPropertiesText.Text = ""
        $script:InfoPropertiesText.Visibility = "Collapsed"
    }
}

function Set-AppStatus {
    param(
        [string]$Status
    )

    if ($script:AppStatusText) {
        $script:AppStatusText.Text = "Status: $Status"
    }
}

function Get-ImageAnalysisSummary {
    param(
        [Parameter(Mandatory)]
        [System.Windows.Media.Imaging.BitmapSource]$Bitmap,
        [Parameter(Mandatory)]
        [string]$CacheKey
    )

    if ($script:AnalysisCache.ContainsKey($CacheKey)) {
        return $script:AnalysisCache[$CacheKey]
    }

    try {
        $formatted = New-Object System.Windows.Media.Imaging.FormatConvertedBitmap
        $formatted.BeginInit()
        $formatted.Source = $Bitmap
        $formatted.DestinationFormat = [System.Windows.Media.PixelFormats]::Bgra32
        $formatted.EndInit()

        $width = [Math]::Max(1, [int]$formatted.PixelWidth)
        $height = [Math]::Max(1, [int]$formatted.PixelHeight)
        $stride = $width * 4
        $pixels = New-Object byte[] ($stride * $height)
        $formatted.CopyPixels($pixels, $stride, 0)

        $stepX = [Math]::Max(1, [int][Math]::Floor($width / 160))
        $stepY = [Math]::Max(1, [int][Math]::Floor($height / 120))
        $histogram = New-Object int[] 16
        $sampleCount = 0
        $sum = 0.0
        $sumSquares = 0.0
        $focusSum = 0.0
        $focusSamples = 0

        for ($y = 0; $y -lt $height; $y += $stepY) {
            for ($x = 0; $x -lt $width; $x += $stepX) {
                $index = ($y * $stride) + ($x * 4)
                $b = [double]$pixels[$index]
                $g = [double]$pixels[$index + 1]
                $r = [double]$pixels[$index + 2]
                $luminance = (0.114 * $b) + (0.587 * $g) + (0.299 * $r)

                $bin = [Math]::Min(15, [int][Math]::Floor($luminance / 16.0))
                $histogram[$bin] += 1
                $sum += $luminance
                $sumSquares += ($luminance * $luminance)
                $sampleCount += 1

                if (($x + $stepX) -lt $width) {
                    $rightIndex = ($y * $stride) + (($x + $stepX) * 4)
                    $rightLuma = (0.114 * [double]$pixels[$rightIndex]) + (0.587 * [double]$pixels[$rightIndex + 1]) + (0.299 * [double]$pixels[$rightIndex + 2])
                    $focusSum += [Math]::Abs($luminance - $rightLuma)
                    $focusSamples += 1
                }

                if (($y + $stepY) -lt $height) {
                    $downIndex = (($y + $stepY) * $stride) + ($x * 4)
                    $downLuma = (0.114 * [double]$pixels[$downIndex]) + (0.587 * [double]$pixels[$downIndex + 1]) + (0.299 * [double]$pixels[$downIndex + 2])
                    $focusSum += [Math]::Abs($luminance - $downLuma)
                    $focusSamples += 1
                }
            }
        }

        if ($sampleCount -le 0) {
            return ""
        }

        $mean = $sum / $sampleCount
        $variance = [Math]::Max(0.0, ($sumSquares / $sampleCount) - ($mean * $mean))
        $contrast = [Math]::Sqrt($variance)
        $focusScore = if ($focusSamples -gt 0) { $focusSum / $focusSamples } else { 0.0 }
        $focusLabel = if ($focusScore -lt 8.0) {
            "soft"
        } elseif ($focusScore -lt 16.0) {
            "ok"
        } else {
            "sharp"
        }

        $brightnessLabel = if ($mean -lt 85) {
            "dark"
        } elseif ($mean -gt 170) {
            "bright"
        } else {
            "balanced"
        }

        $toneZones = @(
            @{ Label = "Shadows";   Start = 0;  End = 2  }
            @{ Label = "Darks";     Start = 3;  End = 5  }
            @{ Label = "Midtones";  Start = 6;  End = 9  }
            @{ Label = "Brights";   Start = 10; End = 12 }
            @{ Label = "Highlights"; Start = 13; End = 15 }
        )

        $toneLines = foreach ($zone in $toneZones) {
            $zoneCount = 0
            for ($i = $zone.Start; $i -le $zone.End; $i++) {
                $zoneCount += $histogram[$i]
            }

            $ratio = if ($sampleCount -gt 0) { $zoneCount / [double]$sampleCount } else { 0.0 }
            $filled = [Math]::Max(0, [Math]::Min(10, [int][Math]::Round($ratio * 10)))
            $bar = ("#" * $filled).PadRight(10, "-")
            "{0,-10} {1,3:0}% [{2}]" -f $zone.Label, ($ratio * 100.0), $bar
        }

        $summary = @(
            ("Analysis: brightness {0:0} ({1}) | contrast {2:0.0} | focus {3:0.0} ({4})" -f $mean, $brightnessLabel, $contrast, $focusScore, $focusLabel)
            "Tone distribution:"
            ($toneLines -join "`n")
        ) -join "`n"

        $script:AnalysisCache[$CacheKey] = $summary
        return $summary
    } catch {
        return ""
    }
}

function Update-PreviewLayout {
    $targetWidth = if ($script:PreviewHost.ActualWidth -gt 0) {
        $script:PreviewHost.ActualWidth - 24
    } else {
        1400
    }
    $targetHeight = if ($script:PreviewHost.ActualHeight -gt 0) {
        $script:PreviewHost.ActualHeight - 24
    } else {
        900
    }

    $script:PreviewImage.MaxWidth = [Math]::Max(200, $targetWidth)
    $script:PreviewImage.MaxHeight = [Math]::Max(200, $targetHeight)
}

function Get-PreviewImageMetrics {
    if (-not $script:PreviewHost -or -not $script:PreviewImage -or -not $script:PreviewImage.Source) {
        return $null
    }

    $imageWidth = $script:PreviewImage.ActualWidth
    $imageHeight = $script:PreviewImage.ActualHeight
    $hostWidth = $script:PreviewHost.ActualWidth
    $hostHeight = $script:PreviewHost.ActualHeight

    if ($imageWidth -le 0 -or $imageHeight -le 0 -or $hostWidth -le 0 -or $hostHeight -le 0) {
        return $null
    }

    return [pscustomobject]@{
        HostWidth = $hostWidth
        HostHeight = $hostHeight
        ImageWidth = $imageWidth
        ImageHeight = $imageHeight
        Left = ($hostWidth - $imageWidth) / 2.0
        Top = ($hostHeight - $imageHeight) / 2.0
    }
}

function Get-NativeZoomForCurrentPreview {
    $metrics = Get-PreviewImageMetrics
    if ($null -eq $metrics -or -not $script:PreviewImage -or -not $script:PreviewImage.Source) {
        return 1.0
    }

    $sourceBitmap = $script:PreviewImage.Source -as [System.Windows.Media.Imaging.BitmapSource]
    if ($null -eq $sourceBitmap -or $sourceBitmap.PixelWidth -le 0 -or $sourceBitmap.PixelHeight -le 0) {
        return 1.0
    }

    $dpiScaleX = 1.0
    $dpiScaleY = 1.0

    try {
        $dpiInfo = [System.Windows.Media.VisualTreeHelper]::GetDpi($script:PreviewImage)
        $dpiScaleX = [double]$dpiInfo.DpiScaleX
        $dpiScaleY = [double]$dpiInfo.DpiScaleY
    } catch {
        try {
            $source = [System.Windows.PresentationSource]::FromVisual($script:PreviewImage)
            if ($source -and $source.CompositionTarget) {
                $dpiScaleX = [double]$source.CompositionTarget.TransformToDevice.M11
                $dpiScaleY = [double]$source.CompositionTarget.TransformToDevice.M22
            }
        } catch {
        }
    }

    $displayedPixelWidth = [Math]::Max(1.0, [double]$metrics.ImageWidth * $dpiScaleX)
    $displayedPixelHeight = [Math]::Max(1.0, [double]$metrics.ImageHeight * $dpiScaleY)

    $zoomX = $sourceBitmap.PixelWidth / $displayedPixelWidth
    $zoomY = $sourceBitmap.PixelHeight / $displayedPixelHeight
    return [Math]::Max(1.0, [Math]::Min($script:MaxPreviewZoom, [Math]::Max($zoomX, $zoomY)))
}

function Set-PreviewCursor {
    if ($script:IsPanningPreview) {
        $script:PreviewImage.Cursor = [System.Windows.Input.Cursors]::SizeAll
    } elseif ($script:PreviewZoom -gt 1.0) {
        $script:PreviewImage.Cursor = [System.Windows.Input.Cursors]::Hand
    } else {
        $script:PreviewImage.Cursor = [System.Windows.Input.Cursors]::Arrow
    }
}

function Clamp-PreviewTranslation {
    $metrics = Get-PreviewImageMetrics
    if ($null -eq $metrics) {
        return
    }

    $scale = [Math]::Max(1.0, $script:PreviewZoom)
    $scaledWidth = $metrics.ImageWidth * $scale
    $scaledHeight = $metrics.ImageHeight * $scale

    if ($scaledWidth -le $metrics.HostWidth) {
        $script:PreviewTranslateTransform.X = (($metrics.HostWidth - $scaledWidth) / 2.0) - $metrics.Left
    } else {
        $minX = $metrics.HostWidth - $metrics.Left - $scaledWidth
        $maxX = -$metrics.Left
        $script:PreviewTranslateTransform.X = [Math]::Max($minX, [Math]::Min($maxX, $script:PreviewTranslateTransform.X))
    }

    if ($scaledHeight -le $metrics.HostHeight) {
        $script:PreviewTranslateTransform.Y = (($metrics.HostHeight - $scaledHeight) / 2.0) - $metrics.Top
    } else {
        $minY = $metrics.HostHeight - $metrics.Top - $scaledHeight
        $maxY = -$metrics.Top
        $script:PreviewTranslateTransform.Y = [Math]::Max($minY, [Math]::Min($maxY, $script:PreviewTranslateTransform.Y))
    }
}

function Reset-PreviewZoom {
    if (-not $script:PreviewScaleTransform -or -not $script:PreviewTranslateTransform) {
        return
    }

    $script:PreviewZoom = 1.0
    $script:PreviewScaleTransform.ScaleX = 1.0
    $script:PreviewScaleTransform.ScaleY = 1.0
    $script:PreviewTranslateTransform.X = 0.0
    $script:PreviewTranslateTransform.Y = 0.0
    $script:IsPreviewMouseDown = $false
    $script:PreviewMouseDownPoint = $null
    $script:IsPanningPreview = $false
    if ($script:PreviewImage -and $script:PreviewImage.IsMouseCaptured) {
        $script:PreviewImage.ReleaseMouseCapture()
    }
    Set-PreviewCursor
    Save-AppSettings
}

function Set-PreviewZoomAtPoint {
    param(
        [Parameter(Mandatory)]
        [System.Windows.Point]$HostPoint,
        [Parameter(Mandatory)]
        [double]$TargetZoom
    )

    $metrics = Get-PreviewImageMetrics
    if ($null -eq $metrics) {
        return
    }

    $pointX = $HostPoint.X - $metrics.Left
    $pointY = $HostPoint.Y - $metrics.Top

    $pointX = [Math]::Max(0.0, [Math]::Min([double]$metrics.ImageWidth, $pointX))
    $pointY = [Math]::Max(0.0, [Math]::Min([double]$metrics.ImageHeight, $pointY))

    $newZoom = [Math]::Max(1.0, [Math]::Min($script:MaxPreviewZoom, $TargetZoom))
    $currentZoom = [Math]::Max(1.0, $script:PreviewZoom)

    if ([Math]::Abs($newZoom - 1.0) -lt 0.01) {
        Reset-PreviewZoom
        Save-AppSettings
        return
    }

    $contentX = ($pointX - $script:PreviewTranslateTransform.X) / $currentZoom
    $contentY = ($pointY - $script:PreviewTranslateTransform.Y) / $currentZoom

    $script:PreviewZoom = $newZoom
    $script:PreviewScaleTransform.ScaleX = $newZoom
    $script:PreviewScaleTransform.ScaleY = $newZoom
    $script:PreviewTranslateTransform.X = $pointX - ($contentX * $newZoom)
    $script:PreviewTranslateTransform.Y = $pointY - ($contentY * $newZoom)
    Clamp-PreviewTranslation
    Set-PreviewCursor
    Save-AppSettings
}

function Apply-PreviewZoomCentered {
    param(
        [Parameter(Mandatory)]
        [double]$TargetZoom
    )

    $metrics = Get-PreviewImageMetrics
    if ($null -eq $metrics) {
        return
    }

    $centerPoint = New-Object System.Windows.Point(($metrics.HostWidth / 2.0), ($metrics.HostHeight / 2.0))
    Set-PreviewZoomAtPoint -HostPoint $centerPoint -TargetZoom $TargetZoom
}

function Zoom-PreviewAtPoint {
    param(
        [Parameter(Mandatory)]
        [System.Windows.Point]$HostPoint
    )

    $selectedItem = if ($script:FileList) { $script:FileList.SelectedItem } else { $null }
    if ($selectedItem) {
        try {
            Set-AppStatus -Status "Loading 100% preview..."
            Flush-Ui

            $nativePreview = Get-100PercentPreviewBitmap -Item $selectedItem
            if ($nativePreview -and $script:RenderedPreviewKey -ne $nativePreview.CacheKey) {
                $script:PreviewImage.Source = $nativePreview.Bitmap
                $script:PreviewImage.Visibility = "Visible"
                $script:EmptyText.Visibility = "Collapsed"
                $script:RenderedPreviewKey = $nativePreview.CacheKey
                Update-PreviewLayout
                Flush-Ui
            }
        } catch {
        }
    }

    $nativeZoom = Get-NativeZoomForCurrentPreview
    Set-PreviewZoomAtPoint -HostPoint $HostPoint -TargetZoom $nativeZoom
    Set-AppStatus -Status "Ready"
}

function Step-PreviewZoomAtPoint {
    param(
        [Parameter(Mandatory)]
        [System.Windows.Point]$HostPoint,
        [Parameter(Mandatory)]
        [int]$WheelDelta
    )

    $stepFactor = 1.25
    $currentZoom = [Math]::Max(1.0, $script:PreviewZoom)
    $targetZoom = if ($WheelDelta -gt 0) {
        $currentZoom * $stepFactor
    } else {
        $currentZoom / $stepFactor
    }

    if ($targetZoom -lt 1.05) {
        $targetZoom = 1.0
    }

    Set-PreviewZoomAtPoint -HostPoint $HostPoint -TargetZoom $targetZoom
}

function Start-PreviewPan {
    param(
        [Parameter(Mandatory)]
        [System.Windows.Point]$HostPoint
    )

    if ($script:PreviewZoom -le 1.0) {
        return
    }

    $script:IsPanningPreview = $true
    $script:PanStartPoint = $HostPoint
    $script:PanStartTranslateX = $script:PreviewTranslateTransform.X
    $script:PanStartTranslateY = $script:PreviewTranslateTransform.Y
    $script:PreviewImage.CaptureMouse() | Out-Null
    Set-PreviewCursor
}

function Update-PreviewPan {
    param(
        [Parameter(Mandatory)]
        [System.Windows.Point]$HostPoint
    )

    if (-not $script:IsPanningPreview) {
        return
    }

    $script:PreviewTranslateTransform.X = $script:PanStartTranslateX + ($HostPoint.X - $script:PanStartPoint.X)
    $script:PreviewTranslateTransform.Y = $script:PanStartTranslateY + ($HostPoint.Y - $script:PanStartPoint.Y)
    Clamp-PreviewTranslation
}

function Stop-PreviewPan {
    $script:IsPreviewMouseDown = $false
    $script:PreviewMouseDownPoint = $null
    if (-not $script:IsPanningPreview) {
        if ($script:PreviewImage -and $script:PreviewImage.IsMouseCaptured) {
            $script:PreviewImage.ReleaseMouseCapture()
        }
        Set-PreviewCursor
        return
    }

    $script:IsPanningPreview = $false
    if ($script:PreviewImage -and $script:PreviewImage.IsMouseCaptured) {
        $script:PreviewImage.ReleaseMouseCapture()
    }
    Set-PreviewCursor
    Save-AppSettings
}

function Focus-FileList {
    if ($script:FileList -and $script:FileList.Items.Count -gt 0) {
        $script:FileList.Focus() | Out-Null
        [void][System.Windows.Input.Keyboard]::Focus($script:FileList)
    }
}

function Show-NoPreview {
    param(
        [string]$Message,
        [object]$Item
    )

    $script:PreviewImage.Source = $null
    $script:PreviewImage.Visibility = "Collapsed"
    $script:RenderedPreviewKey = $null
    Reset-PreviewZoom
    $script:EmptyText.Text = $Message
    $script:EmptyText.Visibility = "Visible"
    Set-AppStatus -Status "Ready"
    Clear-HistogramView

    if ($null -eq $Item) {
        Set-StatusText -Details "No item selected." -Source "" -Exif "" -Analysis "" -Overlay ""
        return
    }

    $details = "{0}`nFiles: {1}" -f $Item.Label, $Item.Extensions
    $source = if ($Item.RawPath) { "RAW file: $($Item.RawPath)" } else { "" }
    Set-StatusText -Details $details -Source $source -Exif (Get-ExifSummary -Item $Item) -Analysis "" -Overlay (Get-ExifOverlayText -Item $Item)
}

function Update-Preview {
    $item = $script:FileList.SelectedItem
    Stop-PreviewPreloading
    if ($item -and $item.PreferredPath) {
        $script:LastPreferredPath = $item.PreferredPath
    }
    if ($null -eq $item) {
        Show-NoPreview -Message "Select an item from the list." -Item $null
        return
    }

    if (-not $item.HasQuickPreview -or -not $item.PreferredPath) {
        Show-NoPreview -Message "No preview source available for this file." -Item $item
        return
    }

    try {
        Update-PreviewLayout
        Set-AppStatus -Status "Loading preview..."
        Flush-Ui

        $preview = Get-PreviewBitmap -Item $item -TargetWidth $script:MainPreviewDecodeWidth -TargetHeight $script:MainRawPreviewSize
        if ($script:RenderedPreviewKey -eq $preview.CacheKey) {
            Save-AppSettings
            return
        }

        $script:PreviewImage.Source = $preview.Bitmap
        $script:PreviewImage.Visibility = "Visible"
        $script:EmptyText.Visibility = "Collapsed"
        $script:RenderedPreviewKey = $preview.CacheKey
        Reset-PreviewZoom
        Flush-Ui

        $histCacheKey = $preview.CacheKey
        if (-not $script:HistogramCache) {
            $script:HistogramCache = @{}
        }
        $histData = $script:HistogramCache[$histCacheKey]
        if (-not $histData) {
            $histData = Get-HistogramData -BitmapSource $preview.Bitmap -Bins 256
            if ($histData) {
                $script:HistogramCache[$histCacheKey] = $histData
            }
        }
        if ($histData) {
            Render-Histogram -Histogram $histData
        } else {
            Clear-HistogramView
        }

        if ($script:ShouldRestoreZoomForSelection -and $script:PendingRestorePreviewZoom -gt 1.0) {
            Apply-PreviewZoomCentered -TargetZoom $script:PendingRestorePreviewZoom
            $script:ShouldRestoreZoomForSelection = $false
        }

        $details = "{0}`nFiles: {1}" -f $item.Label, $item.Extensions
        if ($item.PreviewPath) {
            $source = $preview.Source
            if ($item.RawPath) {
                $source = "{0}`nRAW: {1}" -f $source, $item.RawPath
            }
        } else {
            $source = $preview.Source
        }

        Set-AppStatus -Status "Reading EXIF..."
        Flush-Ui
        $exifDetails = Get-ExifDetails -Item $item

        Set-StatusText -Details $details -Source $source -Exif $exifDetails.Summary -Analysis "" -Overlay $exifDetails.Overlay
        Queue-PreviewPreloadForSelection
        Queue-100PercentPreviewPreloadForSelection -Item $item
        Set-AppStatus -Status "Ready"
        Save-AppSettings
    } catch {
        Show-NoPreview -Message "Could not load preview image." -Item $item
    }
}

function Refresh-Items {
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    $script:CurrentFolder = [System.IO.Path]::GetFullPath($Path)
    Stop-ThumbnailLoading
    Stop-PreviewPreloading
    Set-AppStatus -Status "Scanning folder..."
    Flush-Ui
    $script:Items = @(Get-GroupedImages -Path $script:CurrentFolder)
    Apply-PreviewSourceModeToItems
    Apply-FailedMarksToItems
    $script:PreviewCache = @{}
    $script:MetadataCache = @{}
    $script:AnalysisCache = @{}
    $script:HistogramCache = @{}
    Update-FailFolderText
    $script:CountText.Text = "{0} grouped images" -f $script:Items.Count
    $script:Window.Title = "LightRAW.R - $($script:CurrentFolder)"
    $script:RenderedPreviewKey = $null
    Refresh-FileListItems
    Start-ThumbnailLoading
    Reset-PreviewZoom

    $preferredSelection = if ($script:PendingRestoreSelectionPath) {
        $script:PendingRestoreSelectionPath
    } elseif ($script:LastPreferredPath) {
        $script:LastPreferredPath
    } else {
        $null
    }

    if ($script:Items.Count -gt 0) {
        $selectedIndex = 0
        $script:ShouldRestoreZoomForSelection = $false

        if ($preferredSelection) {
            $match = $script:Items | Where-Object { $_.PreferredPath -eq $preferredSelection } | Select-Object -First 1
            if ($match) {
                $selectedIndex = [Array]::IndexOf($script:Items, $match)
                if ($selectedIndex -lt 0) {
                    $selectedIndex = 0
                } elseif ($script:PendingRestorePreviewZoom -gt 1.0) {
                    $script:ShouldRestoreZoomForSelection = $true
                }
            }
        }

        $script:FileList.SelectedIndex = $selectedIndex
        if ($script:Window.IsLoaded) {
            $script:Window.Dispatcher.BeginInvoke([action]{ Focus-FileList }) | Out-Null
        }
        Set-AppStatus -Status "Ready"
    } else {
        Show-NoPreview -Message "No image files found in this folder." -Item $null
    }

    $script:PendingRestoreSelectionPath = $null
    Save-AppSettings
}

function Show-FolderPathDialog {
    param(
        [Parameter(Mandatory)]
        [string]$Title,
        [Parameter(Mandatory)]
        [string]$Prompt,
        [string]$InitialPath
    )

    [xml]$dialogXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="$Title"
        Width="640"
        Height="200"
        MinHeight="200"
    SizeToContent="Height"
        WindowStartupLocation="CenterOwner"
        ResizeMode="NoResize"
        ShowInTaskbar="False"
        Background="#FF232323"
        Foreground="#FFF5F5F5">
  <Grid Margin="18">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto" />
      <RowDefinition Height="Auto" />
      <RowDefinition Height="Auto" />
    </Grid.RowDefinitions>

    <TextBlock x:Name="PromptText"
               FontSize="14"
               TextWrapping="Wrap"
               Foreground="#FFD8D8D8" />

        <Grid Grid.Row="1"
            Margin="0,12,0,0">
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="*" />
        <ColumnDefinition Width="Auto" />
      </Grid.ColumnDefinitions>

    <Border Grid.Column="0"
          MinWidth="420"
          CornerRadius="8"
          Background="#FF101010"
          BorderBrush="#FF404040"
          BorderThickness="1">
      <TextBox x:Name="PathTextBox"
             Padding="8,6"
             Background="Transparent"
             BorderThickness="0"
             Foreground="#FFF5F5F5" />
    </Border>

      <Button x:Name="BrowseButton"
              Grid.Column="1"
              Margin="10,0,0,0"
              MinWidth="118"
              MinHeight="30"
              Padding="18,7"
              Background="#FF2B5D8A"
              BorderBrush="#FF2B5D8A"
              Foreground="White"
                            Content="Browse...">
                <Button.Template>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="ButtonBorder"
                                CornerRadius="6"
                                        Background="{TemplateBinding Background}"
                                        BorderBrush="{TemplateBinding BorderBrush}"
                                        BorderThickness="{TemplateBinding BorderThickness}">
                            <ContentPresenter HorizontalAlignment="Center"
                                                                VerticalAlignment="Center" />
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="ButtonBorder" Property="Opacity" Value="0.92" />
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="ButtonBorder" Property="Opacity" Value="0.82" />
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter TargetName="ButtonBorder" Property="Opacity" Value="0.5" />
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Button.Template>
            </Button>
    </Grid>

    <StackPanel Grid.Row="2"
                Orientation="Horizontal"
                HorizontalAlignment="Right"
                Margin="0,18,0,0">
      <Button x:Name="CancelButton"
              Content="Cancel"
              IsCancel="True"
              MinWidth="108"
              MinHeight="30"
              Padding="18,7"
              Margin="0,0,8,0"
              Background="#FF444444"
              BorderBrush="#FF444444"
                            Foreground="White">
                <Button.Template>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="ButtonBorder"
                                CornerRadius="6"
                                        Background="{TemplateBinding Background}"
                                        BorderBrush="{TemplateBinding BorderBrush}"
                                        BorderThickness="{TemplateBinding BorderThickness}">
                            <ContentPresenter HorizontalAlignment="Center"
                                                                VerticalAlignment="Center" />
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="ButtonBorder" Property="Opacity" Value="0.92" />
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="ButtonBorder" Property="Opacity" Value="0.82" />
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter TargetName="ButtonBorder" Property="Opacity" Value="0.5" />
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Button.Template>
            </Button>
      <Button x:Name="OkButton"
              Content="OK"
              IsDefault="True"
              MinWidth="108"
              MinHeight="30"
              Padding="18,7"
              Background="#FF2B5D8A"
              BorderBrush="#FF2B5D8A"
                            Foreground="White">
                <Button.Template>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="ButtonBorder"
                                CornerRadius="6"
                                        Background="{TemplateBinding Background}"
                                        BorderBrush="{TemplateBinding BorderBrush}"
                                        BorderThickness="{TemplateBinding BorderThickness}">
                            <ContentPresenter HorizontalAlignment="Center"
                                                                VerticalAlignment="Center" />
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="ButtonBorder" Property="Opacity" Value="0.92" />
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="ButtonBorder" Property="Opacity" Value="0.82" />
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter TargetName="ButtonBorder" Property="Opacity" Value="0.5" />
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Button.Template>
            </Button>
    </StackPanel>
  </Grid>
</Window>
"@

    $reader = New-Object System.Xml.XmlNodeReader $dialogXaml
    $dialog = [Windows.Markup.XamlReader]::Load($reader)
    if ($script:Window -and $script:Window.IsLoaded) {
        $dialog.Owner = $script:Window
    }

    $promptText = $dialog.FindName("PromptText")
    $pathTextBox = $dialog.FindName("PathTextBox")
    $browseButton = $dialog.FindName("BrowseButton")
    $cancelButton = $dialog.FindName("CancelButton")
    $okButton = $dialog.FindName("OkButton")

    if ($promptText) {
        $promptText.Text = $Prompt
    }
    if ($pathTextBox -and $InitialPath) {
        $pathTextBox.Text = $InitialPath
        $pathTextBox.SelectAll()
    }

    if ($browseButton -and $pathTextBox) {
        $browseButton.Add_Click({
            $picker = New-Object System.Windows.Forms.OpenFileDialog
            $picker.Title = "Select a file in the target folder"
            $picker.Filter = "All files (*.*)|*.*"
            $picker.FilterIndex = 1
            $picker.CheckFileExists = $true
            $picker.CheckPathExists = $true
            $picker.Multiselect = $false
            $picker.RestoreDirectory = $true
            $picker.DereferenceLinks = $true

            $candidatePath = $pathTextBox.Text
            if (-not [string]::IsNullOrWhiteSpace($candidatePath) -and (Test-Path -LiteralPath $candidatePath -PathType Container)) {
                $picker.InitialDirectory = [System.IO.Path]::GetFullPath($candidatePath)
            } elseif ($script:CurrentFolder -and (Test-Path -LiteralPath $script:CurrentFolder -PathType Container)) {
                $picker.InitialDirectory = $script:CurrentFolder
            }

            if ($picker.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                $selectedFile = $picker.FileName
                if (-not [string]::IsNullOrWhiteSpace($selectedFile)) {
                    $selectedFolder = Split-Path -Parent $selectedFile
                    if ($selectedFolder -and (Test-Path -LiteralPath $selectedFolder -PathType Container)) {
                        $pathTextBox.Text = [System.IO.Path]::GetFullPath($selectedFolder)
                        $pathTextBox.Focus() | Out-Null
                        $pathTextBox.SelectAll()
                    }
                }
            }
        })
    }

    if ($cancelButton) {
        $cancelButton.Add_Click({
            $dialog.DialogResult = $false
            $dialog.Close()
        })
    }

    if ($okButton -and $pathTextBox) {
        $okButton.Add_Click({
            $candidatePath = $pathTextBox.Text.Trim()
            if ([string]::IsNullOrWhiteSpace($candidatePath)) {
                [System.Windows.MessageBox]::Show("Please enter or browse to a folder path.", $Title) | Out-Null
                $pathTextBox.Focus() | Out-Null
                return
            }

            if (-not (Test-Path -LiteralPath $candidatePath -PathType Container)) {
                [System.Windows.MessageBox]::Show("Folder not found:`n$candidatePath", $Title) | Out-Null
                $pathTextBox.Focus() | Out-Null
                return
            }

            $dialog.Tag = [System.IO.Path]::GetFullPath($candidatePath)
            $dialog.DialogResult = $true
            $dialog.Close()
        })
    }

    if ($pathTextBox) {
        $dialog.Add_Loaded({
            $pathTextBox.Focus() | Out-Null
            if (-not [string]::IsNullOrWhiteSpace($pathTextBox.Text)) {
                $pathTextBox.SelectAll()
            }
        })
    }

    $result = $dialog.ShowDialog()
    if ($result -eq $true -and -not [string]::IsNullOrWhiteSpace([string]$dialog.Tag)) {
        return [string]$dialog.Tag
    }

    return $null
}

function Select-Folder {
    $initialPath = if ($script:CurrentFolder -and (Test-Path -LiteralPath $script:CurrentFolder -PathType Container)) {
        $script:CurrentFolder
    } else {
        $null
    }

    $selectedPath = Show-FolderPathDialog `
        -Title "Browse Folder" `
        -Prompt "Paste a folder path or browse to a folder that contains ARW/JPG pairs." `
        -InitialPath $initialPath

    if ($selectedPath) {
        Refresh-Items -Path $selectedPath
    }
}

function Select-FailFolder {
    $initialPath = Get-ActiveFailFolderPath
    if (-not ($initialPath -and (Test-Path -LiteralPath $initialPath -PathType Container)) -and $script:CurrentFolder -and (Test-Path -LiteralPath $script:CurrentFolder -PathType Container)) {
        $initialPath = $script:CurrentFolder
    }

    $selectedPath = Show-FolderPathDialog `
        -Title "Reject Folder" `
        -Prompt "Paste a folder path or browse to a folder where rejected image groups should be moved." `
        -InitialPath $initialPath

    if ($selectedPath) {
        $script:FailDestinationFolder = [System.IO.Path]::GetFullPath($selectedPath)
        Update-FailFolderText
        Save-AppSettings
        Set-AppStatus -Status ("Reject folder set: {0}" -f $script:FailDestinationFolder)
    }
}

function Show-InExplorer {
    $item = $script:FileList.SelectedItem
    if ($null -eq $item -or -not $item.PreferredPath) {
        return
    }

    Start-Process -FilePath "explorer.exe" -ArgumentList "/select,`"$($item.PreferredPath)`""
}

function Open-SelectionInExplorer {
    if (-not $script:FileList -or $script:FileList.SelectedIndex -lt 0) {
        return
    }

    Show-InExplorer
}

function Toggle-FailSelection {
    if (-not $script:FileList -or $script:FileList.SelectedIndex -lt 0) {
        return
    }

    $item = $script:FileList.SelectedItem
    if ($null -eq $item) {
        return
    }

    $failedGroups = Get-FailedGroupSet -FolderPath $script:CurrentFolder
    $groupKey = [string]$item.GroupKey

    if ($failedGroups.ContainsKey($groupKey)) {
        [void]$failedGroups.Remove($groupKey)
        $item.IsFailed = $false
        Set-AppStatus -Status ("Removed reject mark: {0}" -f $item.PrimaryName)
    } else {
        $failedGroups[$groupKey] = $true
        $item.IsFailed = $true
        Set-AppStatus -Status ("Marked reject: {0}" -f $item.PrimaryName)
    }

    Refresh-FileListItems
    $script:FileList.SelectedItem = $item
    $script:FileList.ScrollIntoView($item)
    Save-AppSettings
}

function Move-FailedSelections {
    param(
        [switch]$ThrowOnError
    )

    if (-not $script:CurrentFolder -or -not $script:Items -or $script:Items.Count -le 0) {
        return $false
    }

    $failedItems = @(Get-PendingFailedItems)
    if ($failedItems.Count -le 0) {
        Set-AppStatus -Status "No rejected items to move."
        return $false
    }

    $selectedIndex = if ($script:FileList) { $script:FileList.SelectedIndex } else { -1 }
    $nextItem = $null
    if ($selectedIndex -ge 0) {
        if ($selectedIndex + 1 -lt $script:Items.Count) {
            $nextItem = $script:Items[$selectedIndex + 1]
        } elseif ($selectedIndex -gt 0) {
            $nextItem = $script:Items[$selectedIndex - 1]
        }
    }

    $failFolder = Get-ActiveFailFolderPath

    try {
        if ([string]::IsNullOrWhiteSpace($failFolder)) {
            throw "Reject folder is not set."
        }

        if ($script:CurrentFolder -and [string]::Equals(
                [System.IO.Path]::GetFullPath($failFolder),
                [System.IO.Path]::GetFullPath($script:CurrentFolder),
                [System.StringComparison]::OrdinalIgnoreCase)) {
            throw "Reject folder cannot be the same as the current image folder."
        }

        Ensure-Directory -Path $failFolder
        Set-AppStatus -Status ("Moving rejected items... {0}" -f $failedItems.Count)
        Flush-Ui

        foreach ($failedItem in $failedItems) {
            foreach ($sourcePath in $failedItem.FilePaths) {
                if (-not (Test-Path -LiteralPath $sourcePath -PathType Leaf)) {
                    continue
                }

                $sourceParent = Split-Path -Parent $sourcePath
                if ([string]::Equals($sourceParent, $failFolder, [System.StringComparison]::OrdinalIgnoreCase)) {
                    continue
                }

                $destinationPath = Get-UniquePathInFolder -FolderPath $failFolder -FileName (Split-Path -Leaf $sourcePath)
                Move-Item -LiteralPath $sourcePath -Destination $destinationPath
            }
        }

        $script:FailedMarksByFolder[[System.IO.Path]::GetFullPath($script:CurrentFolder)] = @{}
        $script:PendingRestoreSelectionPath = if ($nextItem) { $nextItem.PreferredPath } else { $null }
        $script:PendingRestorePreviewZoom = 1.0
        $script:ShouldRestoreZoomForSelection = $false
        Refresh-Items -Path $script:CurrentFolder
        Set-AppStatus -Status "Ready"
        return $true
    } catch {
        if ($ThrowOnError) {
            throw
        }
        [System.Windows.MessageBox]::Show($_.Exception.Message, "LightRAW.R") | Out-Null
        Set-AppStatus -Status "Ready"
        return $false
    }
}

