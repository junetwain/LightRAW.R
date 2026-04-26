$reader = New-Object System.Xml.XmlNodeReader $xaml
$script:Window = [Windows.Markup.XamlReader]::Load($reader)
$script:PreviewImage = $script:Window.FindName("PreviewImage")
$script:PreviewOverlayBorder = $script:Window.FindName("PreviewOverlayBorder")
$script:PreviewOverlayText = $script:Window.FindName("PreviewOverlayText")
$script:ShowOverlayCheckBox = $script:Window.FindName("ShowOverlayCheckBox")
$script:EmptyText = $script:Window.FindName("EmptyText")
$script:DetailsText = $script:Window.FindName("DetailsText")
$script:SourceText = $script:Window.FindName("SourceText")
$script:ExifText = $script:Window.FindName("ExifText")
$script:AnalysisText = $script:Window.FindName("AnalysisText")
$script:InfoPropertiesText = $script:Window.FindName("InfoPropertiesText")
$script:HistogramCanvas = $script:Window.FindName("HistogramCanvas")
$script:HistogramStatusText = $script:Window.FindName("HistogramStatusText")
$script:HistogramSummaryText = $script:Window.FindName("HistogramSummaryText")
$script:CountText = $script:Window.FindName("CountText")
$script:FileList = $script:Window.FindName("FileList")
$script:PreviewBorder = $script:Window.FindName("PreviewBorder")
$script:PreviewHost = $script:Window.FindName("PreviewHost")
$script:AppStatusText = $script:Window.FindName("AppStatusText")
$script:QualityLabelText = $script:Window.FindName("QualityLabelText")
$script:QualityPresetComboBox = $script:Window.FindName("QualityPresetComboBox")
$script:PreviewSourceSwitch = $script:Window.FindName("PreviewSourceSwitch")
$script:FailFolderButton = $script:Window.FindName("FailFolderButton")
$script:ToggleFailButton = $script:Window.FindName("ToggleFailButton")
$script:MoveFailedButton = $script:Window.FindName("MoveFailedButton")
$script:FailFolderText = $script:Window.FindName("FailFolderText")
$script:ShowInfoCheckBox = $script:Window.FindName("ShowInfoCheckBox")
$script:PreviewBackgroundSlider = $script:Window.FindName("PreviewBackgroundSlider")
$script:PreviewBackgroundValueText = $script:Window.FindName("PreviewBackgroundValueText")
$script:InfoPanelBorder = $script:Window.FindName("InfoPanelBorder")
$script:InfoPopupScrollViewer = $script:Window.FindName("InfoPopupScrollViewer")
$browseButton = $script:Window.FindName("BrowseButton")
$refreshButton = $script:Window.FindName("RefreshButton")

$transformGroup = New-Object System.Windows.Media.TransformGroup
$script:PreviewScaleTransform = New-Object System.Windows.Media.ScaleTransform(1.0, 1.0)
$script:PreviewTranslateTransform = New-Object System.Windows.Media.TranslateTransform(0.0, 0.0)
[void]$transformGroup.Children.Add($script:PreviewScaleTransform)
[void]$transformGroup.Children.Add($script:PreviewTranslateTransform)
$script:PreviewImage.RenderTransform = $transformGroup
$script:PreviewImage.RenderTransformOrigin = New-Object System.Windows.Point(0.0, 0.0)
Set-PreviewCursor
Update-QualityButtons
Update-PreviewBackground
Update-InfoPanelVisibility
Update-FailFolderText
Update-PreviewOverlayVisibility

if ($script:PreviewSourceSwitch) {
    $isRaw = ($script:PreviewSourceMode -eq "RAW")
    $script:PreviewSourceSwitch.IsChecked = $isRaw
    $script:PreviewSourceSwitch.Content = if ($isRaw) { "RAW" } else { "JPG" }
}

if ($browseButton) {
    $browseButton.Add_Click({ Select-Folder })
}
if ($refreshButton) {
    $refreshButton.Add_Click({
        try {
            Refresh-Items -Path $script:CurrentFolder
        } catch {
            [System.Windows.MessageBox]::Show($_.Exception.Message, "LightRAW.R") | Out-Null
        }
    })
}
if ($script:FailFolderButton) {
    $script:FailFolderButton.Add_Click({ Select-FailFolder })
}
if ($script:QualityPresetComboBox) {
    $script:QualityPresetComboBox.Add_SelectionChanged({
        if (-not $script:QualityPresetComboBox.SelectedItem) {
            return
        }

        $selectedPreset = [string]$script:QualityPresetComboBox.Text
        if ([string]::IsNullOrWhiteSpace($selectedPreset) -or $selectedPreset -eq $script:QualityPresetName) {
            return
        }

        Apply-QualityPreset -PresetName $selectedPreset
        Update-QualityButtons
        $script:RenderedPreviewKey = $null
        if ($script:FileList.SelectedItem) {
            Update-Preview
        }
    })
}
if ($script:PreviewSourceSwitch) {
    $script:PreviewSourceSwitch.Add_Click({
        $newMode = if ([bool]$script:PreviewSourceSwitch.IsChecked) { "RAW" } else { "JPG" }
        $script:PreviewSourceSwitch.Content = $newMode

        if ($script:PreviewSourceMode -ne $newMode) {
            $script:PreviewSourceMode = $newMode
        }

        Update-QualityButtons

        Apply-PreviewSourceModeToItems
        $script:RenderedPreviewKey = $null
        if ($script:FileList) {
            $script:FileList.Items.Refresh()
        }
        if ($script:FileList.SelectedItem) {
            Update-Preview
        }
        Save-AppSettings
    })
}
if ($script:ToggleFailButton) {
    $script:ToggleFailButton.Add_Click({ Toggle-FailSelection })
}
if ($script:MoveFailedButton) {
    $script:MoveFailedButton.Add_Click({ Move-FailedSelections })
}
if ($script:ShowInfoCheckBox) {
    $script:ShowInfoCheckBox.Add_Click({
        $script:ShowInfoPanel = [bool]$script:ShowInfoCheckBox.IsChecked
        Update-InfoPanelVisibility
        Save-AppSettings
    })
}
if ($script:ShowOverlayCheckBox) {
    $script:ShowOverlayCheckBox.Add_Click({
        $script:ShowPreviewOverlay = [bool]$script:ShowOverlayCheckBox.IsChecked
        Update-PreviewOverlayVisibility
        Save-AppSettings
    })
}
if ($script:PreviewBackgroundSlider) {
    $script:PreviewBackgroundSlider.Value = [double]$script:PreviewBackgroundGray
    if ($script:PreviewBackgroundValueText) {
        $script:PreviewBackgroundValueText.Text = ([int][Math]::Round([double]$script:PreviewBackgroundSlider.Value)).ToString()
    }

    $script:PreviewBackgroundSlider.Add_ValueChanged({
        param($sender, $e)

        if (-not $sender) {
            return
        }

        $newGray = [int][Math]::Round([double]$sender.Value)
        $clampedGray = [Math]::Max(0, [Math]::Min(255, $newGray))
        if ($script:PreviewBackgroundGray -ne $clampedGray) {
            $script:PreviewBackgroundGray = $clampedGray
            Update-PreviewBackground
            Save-AppSettings
        } else {
            Update-PreviewBackground
        }
    })
}
if ($script:FileList) {
    $script:FileList.Add_SelectionChanged({
        if ($script:Window -and $script:Window.Dispatcher) {
            $script:Window.Dispatcher.BeginInvoke([action]{ Update-Preview }, [System.Windows.Threading.DispatcherPriority]::Background) | Out-Null
        } else {
            Update-Preview
        }
    })
    $script:FileList.Add_PreviewKeyDown({
        param($sender, $e)

        switch ($e.Key) {
            "Delete" {
                Toggle-FailSelection
                $e.Handled = $true
            }
            "Space" {
                Toggle-FailSelection
                $e.Handled = $true
            }
        }
    })
}

$script:Window.Add_Loaded({
    if ($script:FileList.Items.Count -gt 0 -and $script:FileList.SelectedIndex -lt 0) {
        $script:FileList.SelectedIndex = 0
    }
    Update-PreviewLayout
    Set-InfoPopupDefaultPosition
    Focus-FileList
    Set-AppStatus -Status "Ready"
})
$script:PreviewBorder.Add_SizeChanged({
    Update-PreviewLayout
    Set-InfoPopupDefaultPosition
    if ($script:PreviewZoom -gt 1.0) {
        Clamp-PreviewTranslation
    }
})
$script:PreviewImage.Add_MouseLeftButtonDown({
    param($sender, $e)

    if (-not $script:PreviewImage.Source) {
        return
    }

    $position = $e.GetPosition($script:PreviewHost)
    if ($e.ClickCount -ge 2) {
        Reset-PreviewZoom
        $e.Handled = $true
        return
    }

    if ($script:PreviewZoom -le 1.0) {
        Zoom-PreviewAtPoint -HostPoint $position
    } else {
        $script:IsPreviewMouseDown = $true
        $script:PreviewMouseDownPoint = $position
        $script:PreviewImage.CaptureMouse() | Out-Null
    }

    $e.Handled = $true
})
$script:PreviewImage.Add_MouseMove({
    param($sender, $e)

    if ($script:IsPreviewMouseDown -and -not $script:IsPanningPreview -and $script:PreviewZoom -gt 1.0 -and $script:PreviewMouseDownPoint) {
        $currentPoint = $e.GetPosition($script:PreviewHost)
        $deltaX = [Math]::Abs($currentPoint.X - $script:PreviewMouseDownPoint.X)
        $deltaY = [Math]::Abs($currentPoint.Y - $script:PreviewMouseDownPoint.Y)
        if ($deltaX -ge $script:PreviewPanThreshold -or $deltaY -ge $script:PreviewPanThreshold) {
            Start-PreviewPan -HostPoint $script:PreviewMouseDownPoint
            Update-PreviewPan -HostPoint $currentPoint
            $e.Handled = $true
            return
        }
    }

    if ($script:IsPanningPreview) {
        Update-PreviewPan -HostPoint ($e.GetPosition($script:PreviewHost))
        $e.Handled = $true
    }
})
$script:PreviewImage.Add_MouseLeftButtonUp({
    param($sender, $e)

    if ($script:IsPanningPreview) {
        Stop-PreviewPan
        $e.Handled = $true
        return
    }

    if ($script:IsPreviewMouseDown -and $script:PreviewZoom -gt 1.0) {
        $script:IsPreviewMouseDown = $false
        $script:PreviewMouseDownPoint = $null
        if ($script:PreviewImage -and $script:PreviewImage.IsMouseCaptured) {
            $script:PreviewImage.ReleaseMouseCapture()
        }
        Reset-PreviewZoom
        $e.Handled = $true
    }
})
$script:PreviewImage.Add_LostMouseCapture({ Stop-PreviewPan })
$script:PreviewHost.Add_MouseWheel({
    param($sender, $e)

    if (-not $script:PreviewImage.Source) {
        return
    }

    Step-PreviewZoomAtPoint -HostPoint ($e.GetPosition($script:PreviewHost)) -WheelDelta $e.Delta
    $e.Handled = $true
})
$script:Window.Add_PreviewKeyDown({
    param($sender, $e)

    switch ($e.Key) {
        "Escape" {
            Reset-PreviewZoom
            $e.Handled = $true
        }
    }
})
$script:Window.Add_Closing({
    param($sender, $e)

    if ($script:SkipPendingFailPromptOnClose -and -not $script:AllowNextCloseWithoutPrompt) {
        $script:SkipPendingFailPromptOnClose = $false
    }

    if ($script:SkipPendingFailPromptOnClose) {
        $script:AllowNextCloseWithoutPrompt = $false
        Stop-ThumbnailLoading
        Stop-PreviewPreloading
        Save-AppSettings
        return
    }

    $pendingFailedItems = Get-PendingFailedItems
    if ($pendingFailedItems.Count -le 0) {
        Stop-ThumbnailLoading
        Stop-PreviewPreloading
        Save-AppSettings
        return
    }

    $e.Cancel = $true
    $choice = Show-PendingFailedItemsDialog -FailedCount $pendingFailedItems.Count

    switch ($choice) {
        "close_anyway" {
            $script:AllowNextCloseWithoutPrompt = $true
            $script:SkipPendingFailPromptOnClose = $true
            Stop-ThumbnailLoading
            Stop-PreviewPreloading
            Save-AppSettings
            $script:Window.Dispatcher.BeginInvoke([action]{
                $script:Window.Close()
            }) | Out-Null
        }
        "move_and_close" {
            try {
                $moveSucceeded = Move-FailedSelections -ThrowOnError
                if ($moveSucceeded) {
                    $script:AllowNextCloseWithoutPrompt = $true
                    $script:SkipPendingFailPromptOnClose = $true
                    Stop-ThumbnailLoading
                    Stop-PreviewPreloading
                    Save-AppSettings
                    $script:Window.Dispatcher.BeginInvoke([action]{
                        $script:Window.Close()
                    }) | Out-Null
                }
            } catch {
                [System.Windows.MessageBox]::Show($_.Exception.Message, "LightRAW.R") | Out-Null
                Set-AppStatus -Status "Ready"
            }
        }
        default {
            $script:AllowNextCloseWithoutPrompt = $false
            $script:SkipPendingFailPromptOnClose = $false
            Set-AppStatus -Status "Close cancelled."
            Save-AppSettings
        }
    }
})

try {
    if (-not [string]::IsNullOrWhiteSpace($FolderPath) -and (Test-Path -LiteralPath $FolderPath -PathType Container)) {
        Refresh-Items -Path $FolderPath
    } else {
        $script:CurrentFolder = $null
        if ($script:CountText) {
            $script:CountText.Text = "No folder selected"
        }
        if ($script:Window) {
            $script:Window.Title = "LightRAW.R"
        }
        Update-FailFolderText
        Show-NoPreview -Message "No folder selected. Use Browse Folder to open a folder." -Item $null
        Set-AppStatus -Status "No folder selected."
    }

    [void]$script:Window.ShowDialog()
} catch {
    [System.Windows.MessageBox]::Show($_.Exception.Message, "LightRAW.R") | Out-Null
    exit 1
}
