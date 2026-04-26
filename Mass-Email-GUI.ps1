# Important: Save this file as "UTF-8 with BOM" to ensure German characters (ä, ö, ü) display correctly.
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms

# Hide the console window
$winApi = Add-Type -Name "Win32ShowWindow" -Namespace Win32Functions -MemberDefinition @"
[DllImport("user32.dll")]
public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
[DllImport("kernel32.dll")]
public static extern IntPtr GetConsoleWindow();
"@ -PassThru
$consoleHandle = $winApi::GetConsoleWindow()
if ($consoleHandle -ne [IntPtr]::Zero) {
    $winApi::ShowWindow($consoleHandle, 0) # 0 = SW_HIDE
}

# Load WebView2 - Note: Ensure the DLLs are in the script directory or installed in the GAC
$WpfDll = Join-Path $PSScriptRoot "Microsoft.Web.WebView2.Wpf.dll"
$CoreDll = Join-Path $PSScriptRoot "Microsoft.Web.WebView2.Core.dll"
$LoaderDll = Join-Path $PSScriptRoot "WebView2Loader.dll"

# ===== AUTOMATED DEPENDENCY MANAGEMENT =====
function Ensure-WebView2Dependencies {
    if ((Test-Path $WpfDll) -and (Test-Path $CoreDll) -and (Test-Path $LoaderDll)) {
        return $true
    }

    Write-Host "WebView2 SDK components missing. Downloading lightweight interop DLLs..." -ForegroundColor Cyan
    
    $PackageName = "Microsoft.Web.WebView2"
    $Version = "1.0.2592.51" # Stable version
    $Source = "https://www.nuget.org/api/v2/package/$PackageName/$Version"
    $ZipPath = Join-Path $env:TEMP "$PackageName.$Version.zip"
    $ExtractPath = Join-Path $env:TEMP "$PackageName.$Version"
    
    try {
        # 1. Download the NuGet package (it's just a ZIP)
        Invoke-WebRequest -Uri $Source -OutFile $ZipPath -UseBasicParsing
        
        # 2. Extract files
        if (Test-Path $ExtractPath) { Remove-Item $ExtractPath -Recurse -Force }
        Expand-Archive -Path $ZipPath -DestinationPath $ExtractPath -Force

        # 3. Copy required DLLs to script directory
        # We take the net462 version for WPF and the x64 native loader
        Copy-Item -Path "$ExtractPath\lib\net462\Microsoft.Web.WebView2.Wpf.dll" -Destination $WpfDll -Force
        Copy-Item -Path "$ExtractPath\lib\net462\Microsoft.Web.WebView2.Core.dll" -Destination $CoreDll -Force
        Copy-Item -Path "$ExtractPath\build\native\x64\WebView2Loader.dll" -Destination $LoaderDll -Force

        # 4. Cleanup
        Remove-Item -Path $ZipPath -Force -ErrorAction SilentlyContinue
        Remove-Item -Path $ExtractPath -Recurse -Force -ErrorAction SilentlyContinue
        
        Write-Host "Dependencies installed successfully." -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "Failed to download dependencies: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

# Attempt to bootstrap dependencies
$HasDependencies = Ensure-WebView2Dependencies

# Attempt to load assemblies
try {
    # Load the Core DLL first as Wpf DLL depends on it
    if (Test-Path $CoreDll) {
        Unblock-File -Path $CoreDll -ErrorAction SilentlyContinue
        Add-Type -Path $CoreDll -ErrorAction Stop
    }
    
    if (Test-Path $WpfDll) {
        Unblock-File -Path $WpfDll -ErrorAction SilentlyContinue
        Add-Type -Path $WpfDll -ErrorAction Stop
    }
    elseif ($null -eq ("Microsoft.Web.WebView2.Wpf.WebView2" -as [type])) {
        # Fallback to GAC only if local is missing AND type isn't already loaded
        Add-Type -AssemblyName "Microsoft.Web.WebView2.Wpf" -ErrorAction SilentlyContinue
    }
}
catch { }

[System.Windows.Forms.Application]::EnableVisualStyles()

# ===== CONFIGURATION =====
$Config = @{
    WindowTitle = "Mass Mailer"
    ThemeColors = @{
        DarkBg     = "#1B1A19"
        LightBg    = "#252423"
        AccentBg   = "#323130"
        Foreground = "#F3F2F1"
        Accent     = "#005A9E"
        Success    = "#107C10"
        Error      = "#A4262C"
        Warning    = "#FFB900"
    }
}

# Basic HTML Editor template for the WebView
$Global:EditorHtml = @"
<html><head><style>body { background-color: #252423; color: #F3F2F1; font-family: 'Calibri', 'Segoe UI', sans-serif; font-size: 11pt; margin: 10px; overflow: hidden; } #editor { width: 100%; height: 100%; outline: none; border: none; overflow-y: auto; margin: 0; padding: 0; } p, div { margin-top: 0; }</style></head><body><div id="editor" contenteditable="true" spellcheck="false"></div></body></html>
"@

# Helper function to convert hex to brush
function ConvertTo-Brush {
    param([string]$HexColor)
    $converter = New-Object System.Windows.Media.BrushConverter
    try { return $converter.ConvertFromString($HexColor) }
    catch { return $converter.ConvertFromString("#FFFFFF") }
}

# ===== GLOBAL VARIABLES =====
$Global:Recipients = New-Object System.Collections.ArrayList
$Global:AttachmentMap = @{}
$Global:EmailTemplate = ""
$Global:LogEntries = New-Object System.Collections.ArrayList
$Global:ImportedHtmlBody = ""
$Global:BodyIsWhite = $false
$Global:SendStatus = "Idle" # Idle, Sending, Paused, Stopped

# ===== FUNCTIONS =====
function New-ThemedButton {
    param([string]$Content, [string]$Width = "Auto", [string]$Height = "32", [string]$ToolTip = "")
    $btn = New-Object System.Windows.Controls.Button
    $btn.Content = $Content
    $btn.Width = $Width
    $btn.Height = $Height
    $btn.ToolTip = $ToolTip
    $btn.Foreground = ConvertTo-Brush $Config.ThemeColors.Foreground
    $btn.Background = ConvertTo-Brush $Config.ThemeColors.AccentBg
    $btn.BorderBrush = ConvertTo-Brush "#484644"
    $btn.BorderThickness = "1"
    $btn.Cursor = [System.Windows.Input.Cursors]::Hand
    $btn.FontSize = 12
    $btn.FontWeight = "Bold"
    $btn.Padding = "8,4,8,4"
    
    $btn.Add_MouseEnter({ $this.Background = ConvertTo-Brush $Config.ThemeColors.Accent; $this.Foreground = ConvertTo-Brush "#FFFFFF" })
    $btn.Add_MouseLeave({ $this.Background = ConvertTo-Brush $Config.ThemeColors.AccentBg; $this.Foreground = ConvertTo-Brush $Config.ThemeColors.Foreground })
    
    return $btn
}

function New-ThemedLabel {
    param([string]$Content, [double]$FontSize = 11, [string]$FGColor = "Foreground")
    $lbl = New-Object System.Windows.Controls.Label
    $lbl.Content = $Content
    $lbl.FontSize = $FontSize
    $lbl.Foreground = ConvertTo-Brush $Config.ThemeColors[$FGColor]
    $lbl.Padding = "0,0,0,0"
    return $lbl
}

function New-ThemedTextBox {
    param([string]$Text = "", [bool]$ReadOnly = $false, [bool]$IsMultiline = $false, [double]$Height = "24", [int]$FontSize = 12)
    $txt = New-Object System.Windows.Controls.TextBox
    $txt.Text = $Text
    $txt.IsReadOnly = $ReadOnly
    $txt.Height = $Height
    $txt.Foreground = ConvertTo-Brush $Config.ThemeColors.Foreground
    $txt.Background = ConvertTo-Brush $Config.ThemeColors.LightBg
    $txt.BorderBrush = ConvertTo-Brush $Config.ThemeColors.AccentBg
    $txt.BorderThickness = "1"
    $txt.Padding = "8,4,8,4"
    $txt.FontSize = $FontSize
    $txt.FontWeight = "Normal"
    
    if ($IsMultiline) {
        $txt.TextWrapping = "Wrap"
        $txt.VerticalScrollBarVisibility = "Auto"
        $txt.AcceptsReturn = $true
    }
    return $txt
}

function New-ThemedDataGrid {
    $grid = New-Object System.Windows.Controls.DataGrid
    $grid.Foreground = ConvertTo-Brush $Config.ThemeColors.Foreground
    $grid.Background = ConvertTo-Brush $Config.ThemeColors.LightBg
    $grid.RowBackground = ConvertTo-Brush $Config.ThemeColors.LightBg
    $grid.AlternatingRowBackground = ConvertTo-Brush $Config.ThemeColors.AccentBg
    $grid.BorderBrush = ConvertTo-Brush $Config.ThemeColors.AccentBg
    $grid.BorderThickness = "1"
    $grid.CanUserAddRows = $false
    $grid.CanUserDeleteRows = $false # We handle this manually to keep data in sync
    $grid.AutoGenerateColumns = $false
    $grid.FontSize = 13
    $grid.FontWeight = [System.Windows.FontWeights]::Normal
    $grid.HeadersVisibility = [System.Windows.Controls.DataGridHeadersVisibility]::Column

    # Set headers to normal font weight and ensure foreground visibility
    $headerStyle = New-Object System.Windows.Style -ArgumentList ([System.Windows.Controls.Primitives.DataGridColumnHeader])
    [void]$headerStyle.Setters.Add((New-Object System.Windows.Setter -Property @{
                Property = [System.Windows.Controls.Primitives.DataGridColumnHeader]::FontWeightProperty
                Value    = [System.Windows.FontWeights]::Bold
            }))
    [void]$headerStyle.Setters.Add((New-Object System.Windows.Setter -Property @{
                Property = [System.Windows.Controls.Primitives.DataGridColumnHeader]::FontSizeProperty
                Value    = 14.0
            }))
    [void]$headerStyle.Setters.Add((New-Object System.Windows.Setter -Property @{
                Property = [System.Windows.Controls.Primitives.DataGridColumnHeader]::ForegroundProperty
                Value    = (ConvertTo-Brush $Config.ThemeColors.Foreground)
            }))
    
    # Add a thin border to headers for that professional grid look
    [void]$headerStyle.Setters.Add((New-Object System.Windows.Setter -Property @{
                Property = [System.Windows.Controls.Primitives.DataGridColumnHeader]::BorderThicknessProperty
                Value    = (New-Object System.Windows.Thickness(0, 0, 1, 1))
            }))
    [void]$headerStyle.Setters.Add((New-Object System.Windows.Setter -Property @{
                Property = [System.Windows.Controls.Primitives.DataGridColumnHeader]::BorderBrushProperty
                Value    = (ConvertTo-Brush "#484644")
            }))
    # Push headers 3px to the right to align perfectly with the text inside the data cells
    [void]$headerStyle.Setters.Add((New-Object System.Windows.Setter -Property @{
                Property = [System.Windows.Controls.Primitives.DataGridColumnHeader]::PaddingProperty
                Value    = (New-Object System.Windows.Thickness(3, 0, 0, 0))
            }))

    $grid.ColumnHeaderStyle = $headerStyle

    # Align the actual data text with the 5px header padding
    $cellTextStyle = New-Object System.Windows.Style -ArgumentList ([System.Windows.Controls.TextBlock])
    [void]$cellTextStyle.Setters.Add((New-Object System.Windows.Setter -Property @{
                Property = [System.Windows.Controls.TextBlock]::MarginProperty
                Value    = (New-Object System.Windows.Thickness(5, 0, 0, 0))
            }))
    $grid.Resources.Add([System.Windows.Controls.TextBlock], $cellTextStyle)

    return $grid
}

function Log-Entry {
    param([string]$Message, [string]$Level = "Info", [string]$Email = "")
    
    $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $logMessage = "[$timestamp] [$Level] $Message"
    if ($Email) { $logMessage += " ($Email)" }
    
    $Global:LogEntries += $logMessage
    
    $color = switch ($Level) {
        "Success" { $Config.ThemeColors.Success }
        "Error" { $Config.ThemeColors.Error }
        "Warning" { $Config.ThemeColors.Warning }
        default { $Config.ThemeColors.Foreground }
    }
    
    if ($Global:LogTextBox) {
        try {
            $run = New-Object System.Windows.Documents.Run
            $run.Text = $logMessage + "`n"
            $run.Foreground = ConvertTo-Brush $color
            
            if ($Global:LogTextBox.Document.Blocks.Count -eq 0) {
                $para = New-Object System.Windows.Documents.Paragraph
                $Global:LogTextBox.Document.Blocks.Add($para)
            }
            $Global:LogTextBox.Document.Blocks[0].Inlines.Add($run)
            $Global:LogTextBox.ScrollToEnd()
        }
        catch { }
    }
}

function Get-WebViewContent {
    if ($Global:BodyWebView -and $Global:BodyWebView.CoreWebView2) {
        $task = $Global:BodyWebView.ExecuteScriptAsync("document.getElementById('editor').innerHTML")
        # Wait for the task to complete
        while (-not $task.IsCompleted) { [System.Windows.Forms.Application]::DoEvents() }
        $raw = $task.Result
        # WebView2 returns a JSON string, so we decode it
        return $raw | ConvertFrom-Json
    }
    return ""
}

function Set-WebViewContent {
    param([string]$Html)
    if ($Global:BodyWebView -and $Global:BodyWebView.CoreWebView2) {
        # Convert string to JSON to safely handle all characters (including quotes, backslashes, and large image data)
        $jsonHtml = $Html | ConvertTo-Json
        $Global:BodyWebView.ExecuteScriptAsync("document.getElementById('editor').innerHTML = $jsonHtml") | Out-Null
    }
}

function Initialize-WebView {
    # 1. Define a writable User Data Folder (UDF) in LocalAppData
    # WebView2 fails if it tries to write to System32 (where powershell.exe lives)
    $userDataFolder = Join-Path $env:LOCALAPPDATA "MassMailer_WebView2"
    if (-not (Test-Path $userDataFolder)) { New-Item -ItemType Directory -Path $userDataFolder | Out-Null }

    # 2. Create the WebView2 Environment with the custom data folder
    $envTask = [Microsoft.Web.WebView2.Core.CoreWebView2Environment]::CreateAsync($null, $userDataFolder)
    while (-not $envTask.IsCompleted) { [System.Windows.Forms.Application]::DoEvents() }
    $environment = $envTask.Result

    # 3. Initialize the control with this environment
    $Global:BodyWebView.EnsureCoreWebView2Async($environment) | Out-Null
    
    # 4. Wait for initialization to finish
    while ($null -eq $Global:BodyWebView.CoreWebView2) { [System.Windows.Forms.Application]::DoEvents() }
    
    # 5. Load the editor HTML
    $Global:BodyWebView.NavigateToString($Global:EditorHtml)
    
    # Set default text after a tiny delay to ensure navigation finished
    while ($Global:BodyWebView.IsInitialized -eq $false) { [System.Windows.Forms.Application]::DoEvents() }
    $defaultText = "Guten Tag [NAME],`n`nAnbei erhalten Sie die gew$([char]252)nschten Informationen zur Durchsicht."
    Set-WebViewContent -Html ($defaultText -replace "`r`n|`n", "<br>")
}

function Apply-EditorFormat {
    param([string]$Command, [string]$Value = $null)
    if ($Global:BodyWebView -and $Global:BodyWebView.CoreWebView2) {
        if ($Command -eq "fontSize") {
            # Use styleWithCSS to allow point sizes (pt) instead of the 1-7 scale
            $script = "document.execCommand('styleWithCSS', false, true); document.execCommand('fontSize', false, '$Value" + "pt');"
        }
        else {
            $script = "document.execCommand('$Command', false, "
            $script += if ($Value -ne $null) { "'$Value'" } else { "null" }
            $script += ");"
        }
        $Global:BodyWebView.ExecuteScriptAsync($script) | Out-Null
    }
}

function Toggle-BodyBackground {
    if ($Global:BodyWebView -and $Global:BodyWebView.CoreWebView2) {
        if ($Global:BodyIsWhite) {
            $bgColor = "#252423"
            $fgColor = "#F3F2F1"
            $Global:BodyIsWhite = $false
            $Global:BgToggleBtn.Content = "Light Preview"
        }
        else {
            $bgColor = "#FFFFFF"
            $fgColor = "#000000"
            $Global:BodyIsWhite = $true
            $Global:BgToggleBtn.Content = "Dark Preview"
        }
        $script = "document.body.style.backgroundColor = '$bgColor'; document.body.style.color = '$fgColor';"
        $Global:BodyWebView.ExecuteScriptAsync($script) | Out-Null
    }
}

function Get-EmailBody {
    param([string]$Template, [string]$RecipientName)
    $name = if ($RecipientName) { $RecipientName } else { "Kunde" }
    if ([string]::IsNullOrWhiteSpace($Template)) { return "" }
    return $Template -replace '\[NAME\]', $name
}

function Perform-AutoAssign {
    param([string]$MainFolder, [bool]$Silent = $false)
    if (-not (Test-Path $MainFolder)) { return }
    $assignedCount = 0
    foreach ($i in 0..($Global:Recipients.Count - 1)) {
        $recipient = $Global:Recipients[$i]
        $subfolder = Join-Path $MainFolder $recipient.Name
        if (Test-Path $subfolder) {
            $files = @(Get-ChildItem -Path $subfolder -File)
            if ($files.Count -gt 0) {
                $Global:AttachmentMap[$i] = @($files.FullName)
                $assignedCount++
            }
            elseif (-not $Silent) { Log-Entry "Folder found, but it is empty: $($recipient.Name)" "Warning" }
        }
        elseif (-not $Silent) { Log-Entry "Auto-Assign: No folder found matching '$($recipient.Name)' in $MainFolder" "Warning" }
    }
    Update-RecipientGrid
    if (-not $Silent -and $assignedCount -gt 0) {
        Log-Entry "Auto-assigned files for $assignedCount recipients from folder structure" "Success"
    }
}

function Import-Recipients {
    param([string]$FilePath)
    
    if (-not (Test-Path $FilePath)) {
        [System.Windows.MessageBox]::Show("File not found: $FilePath", "Error", "Ok", "Error")
        return $false
    }
    
    try {
        if ($FilePath -like "*.xlsx" -or $FilePath -like "*.xls") {
            $excel = New-Object -ComObject Excel.Application
            $excel.Visible = $false
            $workbook = $excel.Workbooks.Open($FilePath)
            $worksheet = $workbook.Sheets(1)
            
            $Global:Recipients = New-Object System.Collections.ArrayList
            $row = 2
            while ($worksheet.Cells($row, 1).Value2) {
                $email = "$($worksheet.Cells($row, 2).Value2)".Trim()
                $name = "$($worksheet.Cells($row, 1).Value2)".Trim()
                
                if ($email -match '^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$') {
                    [void]$Global:Recipients.Add(@{
                            Name       = $name
                            Email      = $email
                            Attachment = $null
                            Status     = "Pending"
                        })
                }
                $row++
            }
            
            $workbook.Close($false)
            $excel.Quit()
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
        }
        else {
            $csv = Import-Csv $FilePath
            $Global:Recipients = New-Object System.Collections.ArrayList
            foreach ($row in $csv) {
                $email = if ($row.Email) { $row.Email.Trim() } else { "" }
                # Fixed logic: PowerShell -or returns boolean. Using if/else to get string value.
                $name = if ($row.Name) { $row.Name.Trim() } 
                elseif ($row.FolderName) { $row.FolderName.Trim() } 
                elseif ($row.Client) { $row.Client.Trim() }
                else { $email }
                
                if ($email -match '^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$') {
                    [void]$Global:Recipients.Add(@{
                            Name       = $name
                            Email      = $email
                            Attachment = $null
                            Status     = "Pending"
                        })
                }
            }
        }
        
        Log-Entry "Imported $($Global:Recipients.Count) recipients" "Success"
        # Task 1: Auto-assign from file directory (Silent)
        $parentDir = Split-Path $FilePath -Parent
        Perform-AutoAssign -MainFolder $parentDir -Silent $true
        # Automatically resolve missing names from Outlook Address Book
        Lookup-RecipientNames
        return $true
    }
    catch {
        Log-Entry "Error importing recipients: $_" "Error"
        return $false
    }
}

function Lookup-RecipientNames {
    if ($Global:Recipients.Count -eq 0) { return }
    
    Log-Entry "Attempting to resolve names via Outlook Address Book..." "Info"
    try {
        $Outlook = New-Object -ComObject Outlook.Application
        $resolvedCount = 0
        
        foreach ($i in 0..($Global:Recipients.Count - 1)) {
            $recipient = $Global:Recipients[$i]
            # Only lookup if the name is the email or the default generic name
            if ($recipient.Name -eq $recipient.Email -or $recipient.Name -eq "New Recipient") {
                $dummyMail = $Outlook.CreateItem(0)
                $outRecip = $dummyMail.Recipients.Add($recipient.Email)
                if ($outRecip.Resolve()) {
                    $newName = $outRecip.AddressEntry.Name
                    if ($newName -and $newName -ne $recipient.Email) {
                        $Global:Recipients[$i].Name = $newName
                        $resolvedCount++
                    }
                }
                $dummyMail.Close(1) # olDiscard
            }
            [System.Windows.Forms.Application]::DoEvents()
        }
        Update-RecipientGrid
        Log-Entry "Resolved $resolvedCount name(s) successfully" "Success"
    }
    catch {
        Log-Entry "Error resolving names: $_" "Error"
    }
}

function Update-RecipientGrid {
    $displayList = New-Object System.Collections.ArrayList

    for ($i = 0; $i -lt $Global:Recipients.Count; $i++) {
        $recipient = $Global:Recipients[$i]
        $attachments = $Global:AttachmentMap[$i]
        $attachmentText = "None"
        $tooltipText = ""

        if ($attachments -and $attachments.Count -gt 0) {
            if ($attachments.Count -eq 1) {
                $attachmentText = Split-Path $attachments[0] -Leaf
            }
            else {
                $attachmentText = "$($attachments.Count) Files"
            }
            $tooltipText = ($attachments | ForEach-Object { Split-Path $_ -Leaf }) -join "`n"
        }
        
        $item = New-Object PSObject -Property @{
            Index             = $i
            Name              = $recipient.Name
            Email             = $recipient.Email
            Attachment        = $attachmentText
            AttachmentToolTip = $tooltipText
            Status            = $recipient.Status
        }

        [void]$displayList.Add($item)
    }
    # Force a re-bind to ensure UI refresh
    $Global:RecipientGrid.ItemsSource = $null
    $Global:RecipientGrid.ItemsSource = $displayList
    $Global:RecipientGrid.UpdateLayout() # Explicitly force layout update
}

function Add-RecipientManually {
    [void]$Global:Recipients.Add(@{
            Name       = "New Recipient"
            Email      = "example@domain.com"
            Attachment = $null
            Status     = "Pending"
        })
    Update-RecipientGrid
    Log-Entry "Manual recipient added to list." "Info"
}

function Remove-SelectedRecipients {
    # Capture selected items into a fixed array before we start modifying the source
    $selected = @($Global:RecipientGrid.SelectedItems) | Sort-Object Index -Descending
    if ($selected.Count -eq 0) { return }

    foreach ($item in $selected) {
        $idx = $item.Index
        if ($idx -lt $Global:Recipients.Count) {
            $Global:Recipients.RemoveAt($idx)
            
            # Update Attachment Map: remove the deleted index and shift subsequent ones down
            $newMap = @{}
            foreach ($key in $Global:AttachmentMap.Keys) {
                $intKey = [int]$key
                if ($intKey -lt $idx) {
                    $newMap[$intKey] = $Global:AttachmentMap[$key]
                }
                elseif ($intKey -gt $idx) {
                    $newMap[$intKey - 1] = $Global:AttachmentMap[$key]
                }
            }
            $Global:AttachmentMap = $newMap
        }
    }
    Update-RecipientGrid
    Log-Entry "Removed $($selected.Count) recipient(s) from list." "Info"
}

function Export-Recipients {
    if ($Global:Recipients.Count -eq 0) { return }
    $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveDialog.Filter = "CSV Files (*.csv)|*.csv"
    $saveDialog.FileName = "exported-recipients-$(Get-Date -Format 'yyyyMMdd')"
    if ($saveDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        # Convert the internal list to a flat CSV-friendly format
        $Global:Recipients | Select-Object Name, Email, Status | Export-Csv -Path $saveDialog.FileName -NoTypeInformation -Encoding UTF8
        Log-Entry "Recipient list exported to: $($saveDialog.FileName)" "Success"
    }
}

function Show-FilePickerDialog {
    param([bool]$MultiSelect = $false)
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.InitialDirectory = [Environment]::GetFolderPath('MyDocuments')
    $dialog.Filter = "All Files (*.*)|*.*|PDF Files (*.pdf)|*.pdf|Word Documents (*.docx)|*.docx|Excel Files (*.xlsx)|*.xlsx|ZIP Files (*.zip)|*.zip"
    $dialog.Multiselect = $MultiSelect

    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return if ($MultiSelect) { $dialog.FileNames } else { $dialog.FileName }
    }
    return $null
}

function Show-FolderPickerDialog {
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.SelectedPath = [Environment]::GetFolderPath('MyDocuments')
    
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $dialog.SelectedPath
    }
    return $null
}

function Assign-AttachmentToSelected {
    $selectedIndex = $Global:RecipientGrid.SelectedIndex
    if ($selectedIndex -ge 0) {
        $filePaths = Show-FilePickerDialog -MultiSelect $true
        if ($filePaths) {
            $Global:AttachmentMap[$selectedIndex] = @($filePaths)
            Update-RecipientGrid
            Log-Entry "Assigned file to $($Global:Recipients[$selectedIndex].Email)" "Success"
        }
    }
}

function Assign-SameAttachmentToAll {
    $filePaths = Show-FilePickerDialog -MultiSelect $true
    if ($filePaths) {
        for ($i = 0; $i -lt $Global:Recipients.Count; $i++) {
            $Global:AttachmentMap[$i] = @($filePaths)
        }
        Update-RecipientGrid
        Log-Entry "Assigned same file to all recipients" "Success"
    }
}

function Assign-AttachmentsFromFolders {
    $mainFolder = Show-FolderPickerDialog
    if ($mainFolder) {
        Perform-AutoAssign -MainFolder $mainFolder -Silent $false
    }
}

function Create-FolderStructure {
    if ($Global:Recipients.Count -eq 0) {
        [System.Windows.MessageBox]::Show("Please load recipients first.", "No Recipients", "Ok", "Warning")
        return
    }
    $basePath = Join-Path $PSScriptRoot "attachments"
    if (-not (Test-Path $basePath)) { New-Item -ItemType Directory -Path $basePath | Out-Null }
    
    $created = 0
    foreach ($recipient in $Global:Recipients) {
        # Sanitize name for filesystem
        $safeName = $recipient.Name -replace '[\\\/\:\*\?\"<>\|]', '_'
        $folderPath = Join-Path $basePath $safeName
        if (-not (Test-Path $folderPath)) {
            New-Item -ItemType Directory -Path $folderPath | Out-Null
            $created++
        }
    }
    Log-Entry "Created $created new folders in $basePath" "Success"
    [System.Windows.MessageBox]::Show("Folders created for all recipients in:`n$basePath", "Success", "Ok", "Information")
}

function Save-MailerTemplate {
    $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveDialog.Filter = "Mailer Template (*.json)|*.json"
    $saveDialog.FileName = "mailer-setup"
    if ($saveDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        try {
            # Convert AttachmentMap keys to strings (Required for JSON compatibility in PS 5.1)
            $stringKeyMap = @{}
            foreach ($key in $Global:AttachmentMap.Keys) {
                $stringKeyMap[$key.ToString()] = $Global:AttachmentMap[$key]
            }

            # Use PSCustomObject and explicit Array conversion for PS 5.1 stability
            $templateData = [PSCustomObject]@{
                Subject        = $Global:SubjectTextBox.Text
                Body           = Get-WebViewContent
                Recipients     = @($Global:Recipients)
                AttachmentMap  = $stringKeyMap
                WaitTime       = $Global:WaitTimeTextBox.Text
                PauseEnabled   = $Global:PauseAfterXMessagesCheckBox.IsChecked
                PauseThreshold = $Global:MessagesThresholdTextBox.Text
                PauseDuration  = $Global:PauseDurationMinutesTextBox.Text
            }

            $json = ConvertTo-Json -InputObject $templateData -Depth 20
            if ([string]::IsNullOrWhiteSpace($json)) { throw "JSON conversion resulted in an empty string." }

            Set-Content -Path $saveDialog.FileName -Value $json -Encoding UTF8
            Log-Entry "Setup saved to: $($saveDialog.FileName)" "Success"
        }
        catch {
            Log-Entry "Failed to save setup: $_" "Error"
            [System.Windows.MessageBox]::Show("Error saving setup: $_", "Error", "Ok", "Error")
        }
    }
}

function Export-OutlookMsg {
    $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveDialog.Filter = "Outlook Message (*.msg)|*.msg"
    $saveDialog.FileName = "email-message-$(Get-Date -Format 'yyyyMMdd-HHmmss')"
    if ($saveDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        try {
            $Outlook = New-Object -ComObject Outlook.Application
            $Mail = $Outlook.CreateItem(0) # olMailItem
            $Mail.Subject = $Global:SubjectTextBox.Text
            $Mail.HTMLBody = Get-WebViewContent
            $Mail.SaveAs($saveDialog.FileName, 3) # olMsg
            Log-Entry "Message exported to: $($saveDialog.FileName)" "Success"
        }
        catch {
            Log-Entry "Failed to export .msg: $_" "Error"
            [System.Windows.MessageBox]::Show("Error exporting .msg: $_", "Error", "Ok", "Error")
        }
    }
}

# Load-MessageTemplate function removed as per user request.
# Its functionality is now covered by Load-MailerTemplate for JSON templates
# and Import-OutlookMsg for .msg files.

# The 'Load Template' button in the UI has also been removed.


function Load-MailerTemplate {
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Filter = "Mailer Template (*.json)|*.json"
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $data = Get-Content $dialog.FileName -Raw -Encoding UTF8 | ConvertFrom-Json
        $Global:SubjectTextBox.Text = $data.Subject
        Set-WebViewContent -Html $data.Body
        
        # Load Delay and Pause settings
        if ($null -ne $data.WaitTime) { $Global:WaitTimeTextBox.Text = $data.WaitTime }
        if ($null -ne $data.PauseEnabled) { $Global:PauseAfterXMessagesCheckBox.IsChecked = $data.PauseEnabled }
        if ($null -ne $data.PauseThreshold) { $Global:MessagesThresholdTextBox.Text = $data.PauseThreshold }
        if ($null -ne $data.PauseDuration) { $Global:PauseDurationMinutesTextBox.Text = $data.PauseDuration }

        # Load Recipients
        $Global:Recipients = New-Object System.Collections.ArrayList
        if ($data.Recipients) {
            foreach ($r in $data.Recipients) { [void]$Global:Recipients.Add($r) }
        }
        
        # Load Attachment Map (JSON keys are strings, convert back to Int)
        $Global:AttachmentMap = @{}
        if ($data.AttachmentMap) {
            $data.AttachmentMap.PSObject.Properties | ForEach-Object {
                $Global:AttachmentMap[[int]$_.Name] = $_.Value
            }
        }
        
        Update-RecipientGrid
        Log-Entry "Setup loaded from: $($dialog.FileName)" "Success"
    }
}

function Clear-AllAttachments {
    $result = [System.Windows.MessageBox]::Show("Clear all attachment assignments?", "Confirm", "YesNo", "Question")
    if ($result -eq "Yes") {
        $Global:AttachmentMap = @{}
        Update-RecipientGrid
        Log-Entry "Cleared all attachments" "Info"
    }
}

function Send-MassEmails {
    if ($Global:SendStatus -eq "Sending" -or $Global:SendStatus -eq "Paused") {
        $res = [System.Windows.MessageBox]::Show("Stop the mailing process?", "Confirm Stop", "YesNo", "Warning")
        if ($res -eq "Yes") { Stop-SendingProcess }
        return
    }

    if ($Global:Recipients.Count -eq 0) {
        [System.Windows.MessageBox]::Show("No recipients loaded", "Error", "Ok", "Error")
        return
    }

    if ([System.Windows.MessageBox]::Show("Send emails to $($Global:Recipients.Count) recipient(s)?", "Confirm", "YesNo", "Question") -ne "Yes") { return }

    # Initialize global state for the timer
    $Global:SendStatus = "Sending"
    $Global:CurrentSendIndex = 0
    $Global:SuccessCount = 0
    $Global:ErrorCount = 0
    $Global:BatchSentCount = 0
    $Global:BodyTemplate = Get-WebViewContent
    $Global:OutlookApp = New-Object -ComObject Outlook.Application
    
    # UI Preparation
    $Global:SendButton.Content = "Stop Sending"
    $Global:SendButton.Background = ConvertTo-Brush $Config.ThemeColors.Error
    $Global:PauseButton.Visibility = [System.Windows.Visibility]::Visible
    $Global:ProgressContainer.Visibility = [System.Windows.Visibility]::Visible
    $Global:ProgressText.Content = "0 / $($Global:Recipients.Count)"
    Log-Entry "Starting email send process..." "Info"

    # Create and start Timer
    if ($null -eq $Global:SendTimer) {
        $Global:SendTimer = New-Object System.Windows.Threading.DispatcherTimer
        $Global:SendTimer.Add_Tick({ Send-NextEmailTick })
    }

    # Process first email immediately
    Send-NextEmailTick

    # Set the delay for the next emails and start timer
    $delay = 500
    if (-not [int]::TryParse($Global:WaitTimeTextBox.Text, [ref]$delay)) { $delay = 500 }
    $Global:SendTimer.Interval = [TimeSpan]::FromMilliseconds([Math]::Max(10, $delay))
    $Global:SendTimer.Start()
}

function Send-NextEmailTick {
    if ($Global:SendStatus -ne "Sending") { return }

    if ($Global:CurrentSendIndex -ge $Global:Recipients.Count) {
        Log-Entry "Send complete - Success: $Global:SuccessCount, Failed: $Global:ErrorCount" "Success"
        [System.Windows.MessageBox]::Show("Success: $Global:SuccessCount`nFailed: $Global:ErrorCount", "Complete", "Ok", "Information")
        Stop-SendingProcess
        return
    }

    $i = $Global:CurrentSendIndex
    $recipient = $Global:Recipients[$i]
    $total = $Global:Recipients.Count
    $currentDisplay = $i + 1
    
    # UI Updates
    $Global:ProgressText.Content = "$currentDisplay / $total"
    $Global:ProgressBar.Value = ($currentDisplay / $total) * 100
    $Global:RecipientGrid.SelectedIndex = $i
    $Global:RecipientGrid.ScrollIntoView($Global:RecipientGrid.SelectedItem)

    # Set status to "Sending (X/Y)" and refresh grid immediately
    $recipient.Status = "Sending ($currentDisplay/$total)"
    Update-RecipientGrid

    try {
        $Mail = $Global:OutlookApp.CreateItem(0)
        $Mail.Subject = $Global:SubjectTextBox.Text
        $Mail.HTMLBody = Get-EmailBody -Template $Global:BodyTemplate -RecipientName $recipient.Name
        
        $recip = $Mail.Recipients.Add($recipient.Email)
        $recip.Type = 1
        $recip.Resolve() | Out-Null
        
        $attachments = $Global:AttachmentMap[$i]
        if ($attachments) { foreach ($filePath in $attachments) { $Mail.Attachments.Add($filePath) | Out-Null } }
        
        $Mail.Send()
        $recipient.Status = "Sent"
        Log-Entry "Email $currentDisplay/$total sent" "Success" $recipient.Email
        $Global:SuccessCount++
    }
    catch {
        $recipient.Status = "Failed"
        Log-Entry "Failed to send: $_" "Error" $recipient.Email
        $Global:ErrorCount++
    }

    Update-RecipientGrid
    $Global:CurrentSendIndex++
    $Global:BatchSentCount++

    # Handle Automatic Batch Pause
    $threshold = 0
    $duration = 0
    if ($Global:PauseAfterXMessagesCheckBox.IsChecked -and 
        [int]::TryParse($Global:MessagesThresholdTextBox.Text, [ref]$threshold) -and 
        [int]::TryParse($Global:PauseDurationMinutesTextBox.Text, [ref]$duration) -and
        $threshold -gt 0 -and ($Global:BatchSentCount % $threshold -eq 0)) {
        
        $Global:SendTimer.Stop()
        Log-Entry "Batch threshold reached. Pausing for $duration minute(s)..." "Warning"
        
        # Resume after $duration minutes
        $resumeTimer = New-Object System.Windows.Threading.DispatcherTimer
        $resumeTimer.Interval = [TimeSpan]::FromMinutes($duration)
        $resumeTimer.Add_Tick({
                $this.Stop()
                if ($Global:SendStatus -eq "Sending") { 
                    Log-Entry "Resuming batch..." "Info"
                    $Global:SendTimer.Start() 
                }
            })
        $resumeTimer.Start()
    }
    else {
        # Reset interval to user defined delay for the next tick
        $delay = 500
        [int]::TryParse($Global:WaitTimeTextBox.Text, [ref]$delay) | Out-Null
        $Global:SendTimer.Interval = [TimeSpan]::FromMilliseconds([Math]::Max(10, $delay))
    }
}

function Stop-SendingProcess {
    if ($Global:SendTimer) { $Global:SendTimer.Stop() }
    $Global:SendStatus = "Idle"
    try {
        if ($Global:OutlookApp) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Global:OutlookApp) | Out-Null }
    }
    catch {}
    
    # Reset UI
    $window.Dispatcher.Invoke({
            $Global:SendStatus = "Idle"
            $Global:SendButton.Content = "Send Emails"
            $Global:SendButton.Background = ConvertTo-Brush $Config.ThemeColors.Success
            $Global:PauseButton.Visibility = [System.Windows.Visibility]::Collapsed
            $Global:PauseButton.Content = "Pause"
            $Global:ProgressContainer.Visibility = [System.Windows.Visibility]::Collapsed
            $Global:ProgressBar.Value = 0
            Update-RecipientGrid
        })
}

function Toggle-PauseSend {
    if ($Global:SendStatus -eq "Sending") {
        $Global:SendStatus = "Paused"
        $Global:SendTimer.Stop()
        $Global:PauseButton.Content = "Resume"
    }
    elseif ($Global:SendStatus -eq "Paused") {
        $Global:SendStatus = "Sending"
        $Global:SendTimer.Start()
        $Global:PauseButton.Content = "Pause"
    }
}

function Import-OutlookMsg {
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Filter = "Outlook Message (*.msg)|*.msg"
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        try {
            $Outlook = New-Object -ComObject Outlook.Application
            $namespace = $Outlook.GetNamespace("MAPI")
            $item = $namespace.OpenSharedItem($dialog.FileName)
            
            # Use raw HTML body to preserve formatting and intentional empty rows
            $html = $item.HTMLBody
            
            # Aggressive metadata stripping: Remove XML, Style, and Comment blocks that cause "ghost" space
            $html = $html -replace "(?s)<style.*?>.*?</style>", ""
            $html = $html -replace "(?s)<xml.*?>.*?</xml>", ""
            $html = $html -replace "(?s)<!--.*?-->", ""

            # Process Inline Images: Convert cid (Content-ID) references to Base64
            # This is why images were missing; WebView2 cannot access Outlook's internal attachment store.
            foreach ($attach in $item.Attachments) {
                $pa = $attach.PropertyAccessor
                try {
                    $cid = $pa.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001E")
                    if (-not [string]::IsNullOrEmpty($cid)) {
                        $tempFile = Join-Path $env:TEMP "mm_img_$($attach.FileName)"
                        $attach.SaveAsFile($tempFile)
                        $bytes = [System.IO.File]::ReadAllBytes($tempFile)
                        $base64 = [System.Convert]::ToBase64String($bytes)
                        $ext = [System.IO.Path]::GetExtension($attach.FileName).ToLower().Replace(".", "")
                        $mime = if ($ext -match "jpg|jpeg") { "image/jpeg" } else { "image/$ext" }
                        
                        $html = $html -replace "cid:$cid", "data:$mime;base64,$base64"
                        Remove-Item $tempFile -Force
                    }
                }
                catch {}
            }
            
            # Light Sanitization: Extract body content
            $html = $html -replace "(?i)^.*?<body[^>]*>", "" -replace "(?i)</body>.*$", ""
            
            # Remove the WordSection wrapper which often carries the "un-removable" top margin
            $html = $html -replace "(?i)<div[^>]*class=""?WordSection1""?[^>]*>", ""
            $html = $html -replace "(?i)</div>\s*$", ""
            
            $Global:SubjectTextBox.Text = $item.Subject
            
            # Standardize all top margins to 0 to prevent un-removable vertical gaps
            $html = $html -replace "(?i)margin-top\s*:\s*[^;\""']*", "margin-top:0"

            # Surgical gap fix: Only strip leading "forced" empty lines that Outlook often prepends,
            # while leaving the rest of the intentional styling intact.
            $sanitized = $html.Trim() -replace "^(?i)(\s*(&nbsp;|<br/?>|<p[^>]*>\s*(<o:p>)?\s*(&nbsp;)?\s*(</o:p>)?\s*</p>)\s*)*", ""
            
            # Force the first actual element to have no top margin to remove the "ghost line"
            $sanitized = $sanitized -replace "^(?i)<(p|div|h[1-6])", '<$1 style="margin-top:0 !important;"'
            
            Set-WebViewContent -Html $sanitized
            Log-Entry "Imported message and formatting from: $($dialog.FileName)" "Success"
            
            # Explicitly close and release to prevent the "stuck" app issue
            $item.Close(1) # olDiscard
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($item) | Out-Null
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($namespace) | Out-Null
        }
        catch {
            Log-Entry "Failed to import .msg: $_" "Error"
        }
        finally {
            if ($Outlook) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Outlook) | Out-Null }
            [GC]::Collect()
            [GC]::WaitForPendingFinalizers()
        }
    }
}


function Export-SendLog {
    $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveDialog.Filter = "Text Files (*.txt)|*.txt|CSV Files (*.csv)|*.csv"
    $saveDialog.FileName = "email-send-log-$(Get-Date -Format 'yyyyMMdd-HHmmss')"
    
    if ($saveDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $Global:LogEntries | Out-File -Path $saveDialog.FileName -Encoding UTF8
        Log-Entry "Log exported to: $($saveDialog.FileName)" "Success"
    }
}

# ===== BUILD UI =====
$window = New-Object System.Windows.Window
$window.Title = $Config.WindowTitle
$window.Width = 1000 # Keep width as is
$window.Height = 855
$window.MinWidth = 1000
$window.MinHeight = 855
$window.Background = ConvertTo-Brush $Config.ThemeColors.DarkBg
$window.Foreground = ConvertTo-Brush $Config.ThemeColors.Foreground
$window.WindowStartupLocation = "CenterScreen"

# Global Style to make ScrollBars sleek (narrower but usable)
$sbStyle = New-Object System.Windows.Style -ArgumentList ([System.Windows.Controls.Primitives.ScrollBar])
[void]$sbStyle.Setters.Add((New-Object System.Windows.Setter -Property @{ Property = [System.Windows.Controls.Primitives.ScrollBar]::WidthProperty; Value = 5.0 }))
[void]$sbStyle.Setters.Add((New-Object System.Windows.Setter -Property @{ Property = [System.Windows.Controls.Primitives.ScrollBar]::BackgroundProperty; Value = (ConvertTo-Brush $Config.ThemeColors.AccentBg) }))
[void]$sbStyle.Setters.Add((New-Object System.Windows.Setter -Property @{ Property = [System.Windows.Controls.Primitives.ScrollBar]::BorderThicknessProperty; Value = (New-Object System.Windows.Thickness(0)) }))

[void]$window.Resources.Add([System.Windows.Controls.Primitives.ScrollBar], $sbStyle)

$mainGrid = New-Object System.Windows.Controls.Grid
$mainGrid.Background = ConvertTo-Brush $Config.ThemeColors.DarkBg

[void]$mainGrid.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{Height = "Auto" }))
[void]$mainGrid.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{Height = "*" }))
[void]$mainGrid.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{Height = "Auto" }))

[void]$mainGrid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{Width = "*" }))

# ===== HEADER =====
$headerBorder = New-Object System.Windows.Controls.Border
$headerBorder.Background = ConvertTo-Brush $Config.ThemeColors.AccentBg
$headerBorder.Padding = "16,12,16,8" # Reduced bottom padding to bring it closer to tabs
$titleLabel = New-ThemedLabel "Mass Mailer" 24 "Accent"
$titleLabel.FontWeight = "Bold"
$headerBorder.Child = $titleLabel

[System.Windows.Controls.Grid]::SetRow($headerBorder, 0)
[System.Windows.Controls.Grid]::SetColumnSpan($headerBorder, 1)
[void]$mainGrid.Children.Add($headerBorder)

# ===== CONTENT PANEL =====
$configActionPanel = New-Object System.Windows.Controls.StackPanel # This panel is for Load/Save Config buttons
$configActionPanel.Orientation = "Horizontal"
$configActionPanel.HorizontalAlignment = "Right"
$configActionPanel.VerticalAlignment = "Top"
$configActionPanel.Margin = "0,0,10,0"

$contentStack = New-Object System.Windows.Controls.Grid # Changed from StackPanel to Grid for flexible layout
$contentStack.Margin = "8,8,8,8"

# Clean Grid RowDefinitions (8 rows total)
[void]$contentStack.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{Height = "Auto" })) # 0: Recipients Label
[void]$contentStack.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{Height = "Auto" })) # 1: Buttons Panel
[void]$contentStack.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{Height = "*"; MinHeight = 150 })) # 2: Recipient Grid (STRETCH)
[void]$contentStack.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{Height = "Auto" })) # 3: Splitter
[void]$contentStack.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{Height = "2*"; MinHeight = 250 })) # 4: Editor Container (STRETCH MORE)

# Main Section Title
$recLabelMain = New-ThemedLabel "Recipients & Attachments" 13
$recLabelMain.FontWeight = "Bold"
$recLabelMain.Margin = "0,0,0,8"
[System.Windows.Controls.Grid]::SetRow($recLabelMain, 0)
[void]$contentStack.Children.Add($recLabelMain)

$topLoadBtn = New-ThemedButton "LOAD CONFIG" 100 "22" "Load a complete project configuration (Recipients, Attachments, Subject, and Body)."
$topLoadBtn.FontSize = 9
$topLoadBtn.Margin = "0,2,4,0"
$topLoadBtn.Add_Click({ Load-MailerTemplate })
[void]$configActionPanel.Children.Add($topLoadBtn)

$topSaveBtn = New-ThemedButton "SAVE CONFIG" 100 "22" "Save the current project configuration (Recipients, Attachments, Subject, and Body)."
$topSaveBtn.FontSize = 9
$topSaveBtn.Margin = "0,2,0,0"
$topSaveBtn.Add_Click({ Save-MailerTemplate })
[void]$configActionPanel.Children.Add($topSaveBtn)

# Re-organized Button Section with Titles
$btnGrid = New-Object System.Windows.Controls.Grid
$btnGrid.Margin = "0,0,0,12"
[void]$btnGrid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{Width = (New-Object System.Windows.GridLength(1, [System.Windows.GridUnitType]::Star)) }))
[void]$btnGrid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{Width = (New-Object System.Windows.GridLength(1.2, [System.Windows.GridUnitType]::Star)) }))

# --- Recipients Section ---
$recGroup = New-Object System.Windows.Controls.StackPanel
$recTitle = New-ThemedLabel "Recipients" 13
$recTitle.FontWeight = "Bold"
$recTitle.Margin = "0,0,0,4"
[void]$recGroup.Children.Add($recTitle)

$recBtns = New-Object System.Windows.Controls.StackPanel
$recBtns.Orientation = "Horizontal"

$importBtn = New-ThemedButton "Import" 80 "28" "Import recipients from Excel/CSV."
$importBtn.Margin = "0,0,4,0"
$importBtn.Add_Click({
        $dialog = New-Object System.Windows.Forms.OpenFileDialog
        $dialog.Filter = "Supported Files (*.xlsx;*.csv)|*.xlsx;*.csv"
        if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            if (Import-Recipients $dialog.FileName) { Update-RecipientGrid }
        }
    })
[void]$recBtns.Children.Add($importBtn)

$exportRecipBtn = New-ThemedButton "Export" 80 "28" "Export recipients to CSV."
$exportRecipBtn.Margin = "0,0,4,0"
$exportRecipBtn.Add_Click({ Export-Recipients })
[void]$recBtns.Children.Add($exportRecipBtn)

$addManualBtn = New-ThemedButton "Add" 60 "28" "Add a blank row."
$addManualBtn.Margin = "0,0,4,0"
$addManualBtn.Add_Click({ Add-RecipientManually })
[void]$recBtns.Children.Add($addManualBtn)

$removeBtn = New-ThemedButton "Remove" 70 "28" "Remove selected rows."
$removeBtn.Add_Click({ Remove-SelectedRecipients })
[void]$recBtns.Children.Add($removeBtn)

[void]$recGroup.Children.Add($recBtns)
[System.Windows.Controls.Grid]::SetColumn($recGroup, 0)
[void]$btnGrid.Children.Add($recGroup)

# --- Attachments Section ---
$attGroup = New-Object System.Windows.Controls.StackPanel
$attGroup.Margin = "20,0,0,0"
$attTitle = New-ThemedLabel "Attachments" 13
$attTitle.FontWeight = "Bold"
$attTitle.Margin = "0,0,0,4"
[void]$attGroup.Children.Add($attTitle)

$attBtns = New-Object System.Windows.Controls.StackPanel
$attBtns.Orientation = "Horizontal"

$assignBtn = New-ThemedButton "Assign" 70 "28" "Pick file for selected recipient."
$assignBtn.Margin = "0,0,4,0"
$assignBtn.Add_Click({ Assign-AttachmentToSelected })
[void]$attBtns.Children.Add($assignBtn)

$assignAllBtn = New-ThemedButton "Same for All" 90 "28" "One file for everyone."
$assignAllBtn.Margin = "0,0,4,0"
$assignAllBtn.Add_Click({ Assign-SameAttachmentToAll })
[void]$attBtns.Children.Add($assignAllBtn)

$autoAssignBtn = New-ThemedButton "Auto-Match" 90 "28" "Scan a folder for subfolders matching recipient names. All files within a matching folder will be attached to that recipient."
$autoAssignBtn.Margin = "0,0,4,0"
$autoAssignBtn.Add_Click({ Assign-AttachmentsFromFolders })
[void]$attBtns.Children.Add($autoAssignBtn)

$createFoldersBtn = New-ThemedButton "Create Folders" 100 "28" "Create folders for each recipient in the 'attachments' directory."
$createFoldersBtn.Margin = "0,0,4,0"
$createFoldersBtn.Add_Click({ Create-FolderStructure })
[void]$attBtns.Children.Add($createFoldersBtn)

$clearBtn = New-ThemedButton "Clear All" 80 "28" "Remove all attachments."
$clearBtn.Add_Click({ Clear-AllAttachments })
[void]$attBtns.Children.Add($clearBtn)

[void]$attGroup.Children.Add($attBtns)
[System.Windows.Controls.Grid]::SetColumn($attGroup, 1)
[void]$btnGrid.Children.Add($attGroup)

[System.Windows.Controls.Grid]::SetRow($btnGrid, 1)
[void]$contentStack.Children.Add($btnGrid)

# Grid
$Global:RecipientGrid = New-ThemedDataGrid
$Global:RecipientGrid.MinHeight = 150
$Global:RecipientGrid.VerticalAlignment = "Stretch"

$col1 = New-Object System.Windows.Controls.DataGridTextColumn
[System.Windows.Controls.Grid]::SetRow($Global:RecipientGrid, 2)
$col1.Header = "Name"
$col1.Binding = [System.Windows.Data.Binding]"Name"
$col1.IsReadOnly = $false
$col1.Width = 150
[void]$Global:RecipientGrid.Columns.Add($col1)

$col2 = New-Object System.Windows.Controls.DataGridTextColumn
$col2.Header = "Email"
$col2.Binding = [System.Windows.Data.Binding]"Email"
$col2.IsReadOnly = $false
$col2.Width = 200
[void]$Global:RecipientGrid.Columns.Add($col2)

$col3 = New-Object System.Windows.Controls.DataGridTextColumn
$col3.Header = "Attachment"
$col3.Binding = [System.Windows.Data.Binding]"Attachment"
$col3.IsReadOnly = $true
$col3.Width = "*"
# Set ToolTip binding for the attachment column
$elementStyle = New-Object System.Windows.Style -ArgumentList ([System.Windows.Controls.TextBlock])
$setter = New-Object System.Windows.Setter
$setter.Property = [System.Windows.Controls.TextBlock]::ToolTipProperty
$setter.Value = New-Object System.Windows.Data.Binding -ArgumentList "AttachmentToolTip"
[void]$elementStyle.Setters.Add($setter)
$col3.ElementStyle = $elementStyle
[void]$Global:RecipientGrid.Columns.Add($col3)

$col4 = New-Object System.Windows.Controls.DataGridTextColumn
$col4.Header = "Status"
$col4.Binding = [System.Windows.Data.Binding]"Status"
$col4.IsReadOnly = $true
$col4.Width = 80
[void]$Global:RecipientGrid.Columns.Add($col4)

[void]$contentStack.Children.Add($Global:RecipientGrid)

# Handle double-click on Attachment column to trigger picker
$Global:RecipientGrid.Add_MouseDoubleClick({
        if ($this.CurrentColumn -and $this.CurrentColumn.Header -eq "Attachment") {
            Assign-AttachmentToSelected
        }
    })

# Handle Delete key on the grid
$Global:RecipientGrid.Add_PreviewKeyDown({
        param($s, $e)
        if ($e.Key -eq [System.Windows.Input.Key]::Delete) {
            Remove-SelectedRecipients
            $e.Handled = $true
        }
    })

# Sync grid edits back to the global recipients array
$Global:RecipientGrid.Add_CellEditEnding({
        param($s, $e)
        $row = $e.Row.Item
        if ($e.Column.Header -eq "Name") { $Global:Recipients[$row.Index].Name = $e.EditingElement.Text }
        elseif ($e.Column.Header -eq "Email") { $Global:Recipients[$row.Index].Email = $e.EditingElement.Text }
    })

# --- Editor Group Grid (Nested) ---
$editorGroupGrid = New-Object System.Windows.Controls.Grid
[System.Windows.Controls.Grid]::SetRow($editorGroupGrid, 4)
[void]$contentStack.Children.Add($editorGroupGrid)

[void]$editorGroupGrid.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{Height = "Auto" })) # 0: Subject Label
[void]$editorGroupGrid.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{Height = "Auto" })) # 1: Subject TB
[void]$editorGroupGrid.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{Height = "Auto" })) # 2: Body Label
[void]$editorGroupGrid.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{Height = "Auto" })) # 3: Toolbar
[void]$editorGroupGrid.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{Height = "*" }))    # 4: WebView

# Subject
$subjLabel = New-ThemedLabel "Subject" 13
$subjLabel.FontWeight = "Bold"
[void]$subjLabel.SetValue([System.Windows.Controls.Control]::PaddingProperty, (New-Object System.Windows.Thickness(0, 0, 0, 0)))
$subjLabel.Margin = "0,16,0,8"
[System.Windows.Controls.Grid]::SetRow($subjLabel, 0)
[void]$editorGroupGrid.Children.Add($subjLabel)

$Global:SubjectTextBox = New-ThemedTextBox "Review of your Cloud Backup configuration" $false $false 30 14
$Global:SubjectTextBox.Margin = "0,0,0,8"
[System.Windows.Controls.Grid]::SetRow($Global:SubjectTextBox, 1)
[void]$editorGroupGrid.Children.Add($Global:SubjectTextBox)

# Body
$bodyHeaderPanel = New-Object System.Windows.Controls.StackPanel
$bodyHeaderPanel.Orientation = "Horizontal"
$bodyHeaderPanel.Margin = "0,16,0,8"

$Global:bodyLabel = New-ThemedLabel "Message Body (use [NAME] for placeholder)" 13
$Global:bodyLabel.FontWeight = "Bold"
[void]$Global:bodyLabel.SetValue([System.Windows.Controls.Control]::PaddingProperty, (New-Object System.Windows.Thickness(0, 0, 0, 0)))
[void]$bodyHeaderPanel.Children.Add($Global:bodyLabel)

$Global:BgToggleBtn = New-ThemedButton "Light Preview" 120 24 "Switch between Dark and Light background for the message editor."
$Global:BgToggleBtn.Margin = "20,0,0,0"
$Global:BgToggleBtn.FontSize = 10
$Global:BgToggleBtn.Add_Click({ Toggle-BodyBackground })
[void]$bodyHeaderPanel.Children.Add($Global:BgToggleBtn)

[System.Windows.Controls.Grid]::SetRow($bodyHeaderPanel, 2)
[void]$editorGroupGrid.Children.Add($bodyHeaderPanel)

$formatToolbar = New-Object System.Windows.Controls.StackPanel
$formatToolbar.Orientation = "Horizontal"
$formatToolbar.Margin = "0,0,0,8" # Small margin below toolbar

# Font Family Selector
$fontFamilyCombo = New-Object System.Windows.Controls.ComboBox
$fontFamilyCombo.Width = 120
$fontFamilyCombo.Height = 24
$fontFamilyCombo.Margin = "0,0,8,0"
$fontFamilyCombo.ToolTip = "Font Family"
$fonts = @("Calibri", "Segoe UI", "Arial", "Times New Roman", "Courier New", "Verdana")
foreach ($f in $fonts) { [void]$fontFamilyCombo.Items.Add($f) }
$fontFamilyCombo.Add_SelectionChanged({ if ($this.SelectedItem) { Apply-EditorFormat "fontName" $this.SelectedItem } })
$fontFamilyCombo.SelectedIndex = 0
[void]$formatToolbar.Children.Add($fontFamilyCombo)

# Font Size Selector (Scale 1-7)
$fontSizeCombo = New-Object System.Windows.Controls.ComboBox
$fontSizeCombo.Width = 50
$fontSizeCombo.Height = 24
$fontSizeCombo.Margin = "0,0,8,0"
$fontSizeCombo.ToolTip = "Font Size (pt)"
$sizes = @(8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72)
foreach ($s in $sizes) { [void]$fontSizeCombo.Items.Add($s.ToString()) }
$fontSizeCombo.SelectedItem = "11"
$fontSizeCombo.Add_SelectionChanged({ if ($this.SelectedItem) { Apply-EditorFormat "fontSize" $this.SelectedItem } })
[void]$formatToolbar.Children.Add($fontSizeCombo)

$boldBtn = New-ThemedButton "B" 30 24 "Bold"
$boldBtn.FontWeight = "Bold"
$boldBtn.FontSize = 12
$boldBtn.Margin = "0,0,4,0"
$boldBtn.Add_Click({ Apply-EditorFormat "bold" })
[void]$formatToolbar.Children.Add($boldBtn)

$italicBtn = New-ThemedButton "I" 30 24 "Italic"
$italicBtn.FontStyle = [System.Windows.FontStyles]::Italic
$italicBtn.FontSize = 12
$italicBtn.Margin = "0,0,4,0"
$italicBtn.Add_Click({ Apply-EditorFormat "italic" })
[void]$formatToolbar.Children.Add($italicBtn)

$underlineBtn = New-ThemedButton "" 30 24 "Underline"
$underlineText = New-Object System.Windows.Controls.TextBlock
$underlineText.Text = "U"
$underlineText.TextDecorations = [System.Windows.TextDecorations]::Underline
$underlineBtn.Content = $underlineText
$underlineBtn.FontSize = 12
$underlineBtn.Margin = "0,0,4,0"
$underlineBtn.Add_Click({ Apply-EditorFormat "underline" })
[void]$formatToolbar.Children.Add($underlineBtn)

[System.Windows.Controls.Grid]::SetRow($formatToolbar, 3)
[void]$editorGroupGrid.Children.Add($formatToolbar)

$WebViewType = "Microsoft.Web.WebView2.Wpf.WebView2"
if ($null -ne ($WebViewType -as [type])) {
    # Check if WebView2 is available
    $Global:BodyWebView = New-Object Microsoft.Web.WebView2.Wpf.WebView2
    $Global:BodyWebView.MinHeight = 200
    $Global:BodyWebView.Margin = "0,0,0,8"
    [System.Windows.Controls.Grid]::SetRow($Global:BodyWebView, 4)
    [void]$editorGroupGrid.Children.Add($Global:BodyWebView)
}
else {
    $Global:BodyTextBox = New-ThemedTextBox "ERROR: WebView2 libraries not found.`n`n1. Ensure 'Microsoft.Web.WebView2.Wpf.dll' and 'Microsoft.Web.WebView2.Core.dll' are in: `n$PSScriptRoot`n`n2. Right-click both DLLs > Properties > Unblock.`n3. Ensure WebView2 Runtime is installed." $true $true 200 14
    [System.Windows.Controls.Grid]::SetRow($Global:BodyTextBox, 4)
    [void]$editorGroupGrid.Children.Add($Global:BodyTextBox)
}

# ===== SIDE PANEL =====
$sideBorder = New-Object System.Windows.Controls.Border
$sideBorder.Background = ConvertTo-Brush $Config.ThemeColors.AccentBg
$sideBorder.Padding = "8,8,8,8"

$sidePanel = New-Object System.Windows.Controls.StackPanel
$sidePanel.Orientation = "Vertical"

$logLabel = New-ThemedLabel "Activity Log" 11
$logLabel.FontWeight = "Bold"
$logLabel.Margin = "0,0,0,8"
[void]$sidePanel.Children.Add($logLabel)

$Global:LogTextBox = New-Object System.Windows.Controls.RichTextBox
$Global:LogTextBox.Foreground = ConvertTo-Brush $Config.ThemeColors.Foreground
$Global:LogTextBox.Background = ConvertTo-Brush $Config.ThemeColors.LightBg
$Global:LogTextBox.BorderBrush = ConvertTo-Brush $Config.ThemeColors.AccentBg
$Global:LogTextBox.BorderThickness = "1"
$Global:LogTextBox.Padding = "8,4,8,4"
$Global:LogTextBox.FontSize = 9
$Global:LogTextBox.FontFamily = "Consolas"
$Global:LogTextBox.IsReadOnly = $true
$Global:LogTextBox.Height = 550

$para = New-Object System.Windows.Documents.Paragraph
$para.Foreground = ConvertTo-Brush $Config.ThemeColors.Foreground
[void]$Global:LogTextBox.Document.Blocks.Add($para)

[void]$sidePanel.Children.Add($Global:LogTextBox)

$logBtnPanel = New-Object System.Windows.Controls.StackPanel
$logBtnPanel.Orientation = "Vertical"
$logBtnPanel.Margin = "0,8,0,0"

$exportBtn = New-ThemedButton "Export Log" 100 "32" "Save the current activity log to a text or CSV file for your records."
$exportBtn.Add_Click({ Export-SendLog })
[void]$logBtnPanel.Children.Add($exportBtn)

$clearLogBtn = New-ThemedButton "Clear Log" 100 "32" "Delete all entries from the activity log panel."
$clearLogBtn.Margin = "0,4,0,0"
$clearLogBtn.Add_Click({
        $Global:LogEntries = New-Object System.Collections.ArrayList
        $Global:LogTextBox.Document.Blocks.Clear()
        $p = New-Object System.Windows.Documents.Paragraph
        $p.Foreground = ConvertTo-Brush $Config.ThemeColors.Foreground
        [void]$Global:LogTextBox.Document.Blocks.Add($p)
    })
[void]$logBtnPanel.Children.Add($clearLogBtn)

[void]$sidePanel.Children.Add($logBtnPanel)

$sideBorder.Child = $sidePanel

# ===== TABS =====
$tabControl = New-Object System.Windows.Controls.TabControl
$tabControl.Background = ConvertTo-Brush $Config.ThemeColors.DarkBg
$tabControl.BorderThickness = 0
$tabControl.Padding = New-Object System.Windows.Thickness(0)

$tabMailing = New-Object System.Windows.Controls.TabItem
$tabMailing.Header = "Mailing"
$tabMailing.Padding = "15,0,15,0"
$tabMailing.Height = 22
$tabMailing.VerticalContentAlignment = "Stretch"
$tabMailing.Margin = New-Object System.Windows.Thickness(0)
$tabMailing.FontSize = 11
$tabMailing.FontWeight = "Normal"
$tabMailing.BorderThickness = 0
$tabMailing.Content = $contentStack

$tabLog = New-Object System.Windows.Controls.TabItem
$tabLog.Header = "Activity Log"
$tabLog.Padding = "15,0,15,0"
$tabLog.Height = 22
$tabLog.VerticalContentAlignment = "Stretch"
$tabLog.Margin = New-Object System.Windows.Thickness(0)
$tabLog.FontSize = 11
$tabLog.FontWeight = "Normal"
$tabLog.BorderThickness = 0
$tabLog.Content = $sideBorder

[void]$tabControl.Items.Add($tabMailing)
[void]$tabControl.Items.Add($tabLog)

[System.Windows.Controls.Grid]::SetRow($tabControl, 1)
[System.Windows.Controls.Grid]::SetRow($configActionPanel, 1)
[void]$mainGrid.Children.Add($tabControl)
[void]$mainGrid.Children.Add($configActionPanel)

# ===== PROGRESS BAR =====
$Global:ProgressContainer = New-Object System.Windows.Controls.StackPanel
$Global:ProgressContainer.Orientation = "Horizontal"
$Global:ProgressContainer.HorizontalAlignment = "Center"
$Global:ProgressContainer.Visibility = [System.Windows.Visibility]::Collapsed
$Global:ProgressContainer.Margin = "0"

$Global:ProgressText = New-ThemedLabel "0 / 0" 10
$Global:ProgressText.Margin = "0,0,10,0"
$Global:ProgressText.VerticalAlignment = "Center"

$Global:ProgressBar = New-Object System.Windows.Controls.ProgressBar
$Global:ProgressBar.Height = 2
$Global:ProgressBar.Width = 200
$Global:ProgressBar.Background = ConvertTo-Brush "#444444"
$Global:ProgressBar.Foreground = ConvertTo-Brush $Config.ThemeColors.Accent
$Global:ProgressBar.BorderThickness = 0
$Global:ProgressBar.VerticalAlignment = "Center"

[void]$Global:ProgressContainer.Children.Add($Global:ProgressText)
[void]$Global:ProgressContainer.Children.Add($Global:ProgressBar)

# ===== FOOTER =====
$footerBorder = New-Object System.Windows.Controls.Border
$footerBorder.Background = ConvertTo-Brush $Config.ThemeColors.AccentBg
$footerBorder.Padding = "8,3,8,3"

$footerMainStack = New-Object System.Windows.Controls.Grid
[void]$footerMainStack.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{Height = "Auto" }))
[void]$footerMainStack.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{Height = "*" }))

[System.Windows.Controls.Grid]::SetRow($Global:ProgressContainer, 0)
[void]$footerMainStack.Children.Add($Global:ProgressContainer)

$footerDock = New-Object System.Windows.Controls.DockPanel
$footerDock.LastChildFill = $false
$footerDock.Margin = "0"
[System.Windows.Controls.Grid]::SetRow($footerDock, 1)

$leftFooterPanel = New-Object System.Windows.Controls.StackPanel
$leftFooterPanel.Orientation = "Horizontal"

$msgBtn = New-ThemedButton "Import .msg" 110 "30" "Import an existing Outlook .msg file to use its subject line and body." # This button is kept
$msgBtn.Add_Click({ Import-OutlookMsg })
$msgBtn.Margin = "0,0,4,0"
[void]$leftFooterPanel.Children.Add($msgBtn)

$exportMsgBtn = New-ThemedButton "Export .msg" 110 "30" "Export the current subject and message body as an Outlook .msg file."
[void]$exportMsgBtn.Add_Click({ Export-OutlookMsg }) # Renamed from Save-MessageTemplate
$exportMsgBtn.Margin = "0,0,16,0" # Add right margin to separate from settings panel
[void]$leftFooterPanel.Children.Add($exportMsgBtn)

$settingsPanel = New-Object System.Windows.Controls.StackPanel
$settingsPanel.Orientation = "Vertical" # Stack delay and pause vertically
$settingsPanel.VerticalAlignment = "Center"
$settingsPanel.HorizontalAlignment = "Right"
$settingsPanel.Margin = "0,0,12,0" # Margin to separate from Send button

# Delay (ms) section
$waitPanel = New-Object System.Windows.Controls.StackPanel
$waitPanel.Orientation = "Horizontal"
$waitPanel.VerticalAlignment = "Center" # Ensure vertical alignment within its own panel
$waitPanel.HorizontalAlignment = "Right"
$waitPanel.Margin = "0,0,0,4" # Small margin below delay panel when stacked vertically

$waitLabel = New-ThemedLabel "Delay (ms):" 11
$waitLabel.Margin = "0,2,8,0" # Added 2px top margin for baseline alignment
$waitLabel.VerticalAlignment = "Center"
[void]$waitPanel.Children.Add($waitLabel)

$Global:WaitTimeTextBox = New-ThemedTextBox "500" $false $false 22 11 # Height 22 to prevent clipping
$Global:WaitTimeTextBox.Width = 60
$Global:WaitTimeTextBox.VerticalAlignment = "Center"
$Global:WaitTimeTextBox.ToolTip = "Time to wait between each email (in milliseconds)."
[void]$waitPanel.Children.Add($Global:WaitTimeTextBox)

# Pause after X messages section
$pausePanel = New-Object System.Windows.Controls.StackPanel
$pausePanel.Orientation = "Horizontal"
$pausePanel.VerticalAlignment = "Center" # Ensure vertical alignment within its own panel
$pausePanel.HorizontalAlignment = "Right"
$pausePanel.Margin = "0,0,0,0" # No top margin needed when stacked vertically

$Global:PauseAfterXMessagesCheckBox = New-Object System.Windows.Controls.CheckBox
$Global:PauseAfterXMessagesCheckBox.Content = "Pause after"
$Global:PauseAfterXMessagesCheckBox.Height = 24 # Height 24 ensures checkbox is not cut off
$Global:PauseAfterXMessagesCheckBox.VerticalContentAlignment = "Center"
$Global:PauseAfterXMessagesCheckBox.Foreground = ConvertTo-Brush $Config.ThemeColors.Foreground
$Global:PauseAfterXMessagesCheckBox.VerticalAlignment = "Center"
$Global:PauseAfterXMessagesCheckBox.Margin = "0,3,4,0" # Added 3px top margin to align text baseline
[void]$pausePanel.Children.Add($Global:PauseAfterXMessagesCheckBox)

$Global:MessagesThresholdTextBox = New-ThemedTextBox "10" $false $false 24 11
$Global:MessagesThresholdTextBox.Width = 40
$Global:MessagesThresholdTextBox.VerticalAlignment = "Center"
$Global:MessagesThresholdTextBox.ToolTip = "Number of messages after which to pause."
[void]$pausePanel.Children.Add($Global:MessagesThresholdTextBox)

$pauseLabel = New-ThemedLabel "messages for" 11
$pauseLabel.Margin = "4,2,4,0" # Added 2px top margin for alignment
$pauseLabel.VerticalAlignment = "Center"
[void]$pausePanel.Children.Add($pauseLabel)

$Global:PauseDurationMinutesTextBox = New-ThemedTextBox "5" $false $false 24 11
$Global:PauseDurationMinutesTextBox.Width = 40
$Global:PauseDurationMinutesTextBox.VerticalAlignment = "Center"
$Global:PauseDurationMinutesTextBox.ToolTip = "Duration of the pause in minutes."
[void]$pausePanel.Children.Add($Global:PauseDurationMinutesTextBox) # Add textbox first

$minutesLabel = New-ThemedLabel "minutes" 11 # New label for "minutes"
$minutesLabel.Margin = "4,2,0,0" # Added 2px top margin for alignment
$minutesLabel.VerticalAlignment = "Center"
[void]$pausePanel.Children.Add($minutesLabel) # Add minutes label

[void]$settingsPanel.Children.Add($waitPanel) # Add waitPanel to the vertical settings stack
[void]$settingsPanel.Children.Add($pausePanel) # Add pausePanel below waitPanel

$Global:SendButton = New-ThemedButton "Send Emails" 150 52 "Begin the bulk email process using Outlook. Ensure Outlook is open and you have reviewed all attachments."
$Global:SendButton.FontSize = 14
$Global:SendButton.Background = ConvertTo-Brush $Config.ThemeColors.Success
$Global:SendButton.Foreground = ConvertTo-Brush "#FFFFFF"
$Global:SendButton.Add_Click({ Send-MassEmails })

$Global:PauseButton = New-ThemedButton "Pause" 80 52 "Pause the mailing process."
$Global:PauseButton.Visibility = [System.Windows.Visibility]::Collapsed
$Global:PauseButton.Margin = "0,0,8,0"
$Global:PauseButton.Add_Click({ Toggle-PauseSend })

# Function to dynamically update hover colors based on state
$Global:SendButton.Add_MouseEnter({ 
        if ($Global:SendStatus -eq "Idle") { $this.Background = ConvertTo-Brush "#0D610D" } 
        else { $this.Background = ConvertTo-Brush "#8E2026" }
    })

$Global:SendButton.Add_MouseLeave({ 
        if ($Global:SendStatus -eq "Idle") { $this.Background = ConvertTo-Brush $Config.ThemeColors.Success } 
        else { $this.Background = ConvertTo-Brush $Config.ThemeColors.Error }
        $this.Foreground = ConvertTo-Brush "#FFFFFF"
    })

[void][System.Windows.Controls.DockPanel]::SetDock($leftFooterPanel, [System.Windows.Controls.Dock]::Left)
[void][System.Windows.Controls.DockPanel]::SetDock($Global:SendButton, [System.Windows.Controls.Dock]::Right)
[void][System.Windows.Controls.DockPanel]::SetDock($Global:PauseButton, [System.Windows.Controls.Dock]::Right)
[void][System.Windows.Controls.DockPanel]::SetDock($settingsPanel, [System.Windows.Controls.Dock]::Right) # Dock settingsPanel to the right of SendButton

[void]$footerDock.Children.Add($leftFooterPanel)
[void]$footerDock.Children.Add($Global:SendButton)
[void]$footerDock.Children.Add($Global:PauseButton)
[void]$footerDock.Children.Add($settingsPanel)

$footerBorder.Child = $footerMainStack
$footerMainStack.Children.Add($footerDock)

[System.Windows.Controls.Grid]::SetRow($footerBorder, 2)
[System.Windows.Controls.Grid]::SetColumnSpan($footerBorder, 1)
[void]$mainGrid.Children.Add($footerBorder)

$window.Content = $mainGrid

# Initialize
Log-Entry "Application initialized" "Success"

# Initialize WebView on window loaded event
$window.Add_Loaded({
        if ($Global:BodyWebView) {
            Initialize-WebView
        }
    })

# Show window
[void]$window.ShowDialog()
