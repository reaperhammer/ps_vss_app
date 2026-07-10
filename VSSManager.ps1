# VSS Manager Main Script
# This contains your actual VSS functions and the WPF application startup

# PowerShell Edition Compatibility Check
if ($PSVersionTable.PSEdition -eq 'Core' -or $PSVersionTable.PSVersion.Major -ge 6) {
	Write-Warning "This script is not compatible with PowerShell 7 (Core). Attempting to relaunch in Windows PowerShell 5.1..."
	$winPS = Join-Path $env:WINDIR "System32\WindowsPowerShell\v1.0\powershell.exe"
	if (Test-Path $winPS) {
		& $winPS -NoProfile -ExecutionPolicy Bypass -STA -File $MyInvocation.MyCommand.Path
	} else {
		Write-Error "Windows PowerShell 5.1 (powershell.exe) was not found. Please run this script in Windows PowerShell 5.1."
	}
	exit
}

# Load your existing main.ps1 functions
$scriptPath = Split-Path $MyInvocation.MyCommand.Path
$mainScript = Join-Path $scriptPath "main.ps1"

if (Test-Path $mainScript) {
	. $mainScript
} else {
	Write-Error "Could not find main.ps1 file"
	exit 1
}

# WPF Application Entry Point
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName WindowsBase

# VSS create/delete operations require elevation. Fail early with a clear GUI message
# instead of letting WMI calls fail later with vague access denied errors.
if (-not (Test-VSSAdministrator)) {
	[System.Windows.MessageBox]::Show(
		"VSS Manager must be run as Administrator to manage volume shadow copies.`n`nUse Launch.bat or restart PowerShell as Administrator.",
		"Administrator Required",
		[System.Windows.MessageBoxButton]::OK,
		[System.Windows.MessageBoxImage]::Warning
	) | Out-Null
	exit 1
}

# Ensure STA thread for WPF (under Windows PowerShell 5.1)
if ([Threading.Thread]::CurrentThread.ApartmentState -ne [Threading.ApartmentState]::STA) {
	Write-Host "Restarting in STA mode..."
	$winPS = Join-Path $env:WINDIR "System32\WindowsPowerShell\v1.0\powershell.exe"
	& $winPS -STA -NoProfile -ExecutionPolicy Bypass -File $MyInvocation.MyCommand.Path
	exit
}

# Define the XAML for our window
$xaml = @"
<Window
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="VSS Manager - Volume Shadow Copy Service" 
        Height="700" Width="1000"
        MinHeight="600" MinWidth="800"
        WindowStartupLocation="CenterScreen"
        Background="#F5F5F5"
        FontFamily="Segoe UI"
        FontSize="12"
        Icon=".\assets\png\app_icon.png">
    
    <Window.Resources>
        <!-- Modern Color Scheme -->
        <SolidColorBrush x:Key="PrimaryBrush" Color="#2196F3"/>
        <SolidColorBrush x:Key="PrimaryDarkBrush" Color="#1976D2"/>
        <SolidColorBrush x:Key="AccentBrush" Color="#FF4081"/>
        <SolidColorBrush x:Key="BackgroundBrush" Color="#F5F5F5"/>
        <SolidColorBrush x:Key="SurfaceBrush" Color="#FFFFFF"/>
        <SolidColorBrush x:Key="TextPrimaryBrush" Color="#212121"/>
        <SolidColorBrush x:Key="TextSecondaryBrush" Color="#757575"/>
        <SolidColorBrush x:Key="BorderBrush" Color="#E0E0E0"/>
        <SolidColorBrush x:Key="SuccessBrush" Color="#4CAF50"/>
        <SolidColorBrush x:Key="WarningBrush" Color="#FF9800"/>
        <SolidColorBrush x:Key="ErrorBrush" Color="#F44336"/>
        
        <!-- Modern Button Style -->
        <Style x:Key="ModernButtonStyle" TargetType="Button">
            <Setter Property="Background" Value="{StaticResource PrimaryBrush}"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding" Value="16,8"/>
            <Setter Property="Margin" Value="4"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Cursor" Value="Hand"/>
        </Style>
        
        <!-- Modern TabControl Style -->
        <Style x:Key="ModernTabControlStyle" TargetType="TabControl">
            <Setter Property="Background" Value="{StaticResource SurfaceBrush}"/>
            <Setter Property="BorderBrush" Value="{StaticResource BorderBrush}"/>
            <Setter Property="BorderThickness" Value="1"/>
        </Style>
        
        <!-- Modern TabItem Style -->
        <Style x:Key="ModernTabItemStyle" TargetType="TabItem">
            <Setter Property="Background" Value="{StaticResource BackgroundBrush}"/>
            <Setter Property="Foreground" Value="{StaticResource TextSecondaryBrush}"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Padding" Value="20,12"/>
            <Setter Property="Margin" Value="0"/>
        </Style>
        
        <!-- Modern DataGrid Style -->
        <Style x:Key="ModernDataGridStyle" TargetType="DataGrid">
            <Setter Property="Background" Value="{StaticResource SurfaceBrush}"/>
            <Setter Property="BorderBrush" Value="{StaticResource BorderBrush}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="GridLinesVisibility" Value="Horizontal"/>
            <Setter Property="HorizontalGridLinesBrush" Value="{StaticResource BorderBrush}"/>
            <Setter Property="VerticalGridLinesBrush" Value="Transparent"/>
            <Setter Property="RowBackground" Value="{StaticResource SurfaceBrush}"/>
            <Setter Property="AlternatingRowBackground" Value="#FAFAFA"/>
            <Setter Property="HeadersVisibility" Value="Column"/>
            <Setter Property="SelectionMode" Value="Extended"/>
            <Setter Property="SelectionUnit" Value="FullRow"/>
            <Setter Property="CanUserAddRows" Value="False"/>
            <Setter Property="CanUserDeleteRows" Value="False"/>
            <Setter Property="CanUserReorderColumns" Value="False"/>
            <Setter Property="CanUserResizeColumns" Value="True"/>
            <Setter Property="CanUserSortColumns" Value="True"/>
            <Setter Property="AutoGenerateColumns" Value="False"/>
            <Setter Property="FontSize" Value="11"/>
        </Style>
        
        <!-- Modern DataGridColumnHeader Style -->
        <Style x:Key="ModernDataGridColumnHeaderStyle" TargetType="DataGridColumnHeader">
            <Setter Property="Background" Value="#F8F9FA"/>
            <Setter Property="Foreground" Value="{StaticResource TextPrimaryBrush}"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Padding" Value="12,8"/>
            <Setter Property="BorderBrush" Value="{StaticResource BorderBrush}"/>
            <Setter Property="BorderThickness" Value="0,0,1,1"/>
        </Style>
    </Window.Resources>
    
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        
        <!-- Header Section -->
        <Border Grid.Row="0" Background="{StaticResource SurfaceBrush}" 
                BorderBrush="{StaticResource BorderBrush}" BorderThickness="0,0,0,1"
                Padding="20,16">
            <StackPanel>
                <TextBlock Text="Volume Shadow Copy Service Manager" 
                          FontSize="20" FontWeight="Bold" 
                          Foreground="{StaticResource TextPrimaryBrush}"/>
                <TextBlock Text="Manage volume shadow copies and system snapshots" 
                          FontSize="12" 
                          Foreground="{StaticResource TextSecondaryBrush}"
                          Margin="0,4,0,0"/>
            </StackPanel>
        </Border>
        
        <!-- Main Content -->
        <TabControl Grid.Row="1" Style="{StaticResource ModernTabControlStyle}" Margin="0">
            <!-- Volume Management Tab -->
            <TabItem Style="{StaticResource ModernTabItemStyle}">
                <TabItem.Header>
                    <StackPanel Orientation="Horizontal">
                        <Image Source=".\assets\png\volume.png" Width="16" Height="16" Margin="0,0,8,0"/>
                        <TextBlock Text="Volume Management" VerticalAlignment="Center"/>
                    </StackPanel>
                </TabItem.Header>
                <Grid Margin="20">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>
                    
                    <!-- Volume Management Toolbar -->
                    <Border Grid.Row="0" Background="{StaticResource SurfaceBrush}" 
                            BorderBrush="{StaticResource BorderBrush}" BorderThickness="1"
                            CornerRadius="4" Padding="16" Margin="0,0,0,16">
                        <StackPanel Orientation="Horizontal">
                            <Button Name="btnRefreshVolumes" Style="{StaticResource ModernButtonStyle}"
                                    Width="140" Height="36">
                                <StackPanel Orientation="Horizontal">
                                    <Image Source=".\assets\png\refresh.png" Width="16" Height="16" Margin="0,0,8,0"/>
                                    <TextBlock Text="Refresh Volumes" VerticalAlignment="Center"/>
                                </StackPanel>
                            </Button>
                            <TextBlock Text="Select a volume to view and manage its shadow copies" 
                                      VerticalAlignment="Center" 
                                      Foreground="{StaticResource TextSecondaryBrush}"
                                      Margin="16,0,0,0"/>
                        </StackPanel>
                    </Border>
                    
                    <!-- Volumes DataGrid -->
                    <DataGrid Name="dgVolumes" Grid.Row="1" 
                              Style="{StaticResource ModernDataGridStyle}">
                        <DataGrid.ColumnHeaderStyle>
                            <Style TargetType="DataGridColumnHeader" BasedOn="{StaticResource ModernDataGridColumnHeaderStyle}"/>
                        </DataGrid.ColumnHeaderStyle>
                        <DataGrid.Columns>
                            <DataGridTemplateColumn Header="" Width="50">
                                <DataGridTemplateColumn.CellTemplate>
                                    <DataTemplate>
                                         <CheckBox HorizontalAlignment="Center" VerticalAlignment="Center"
                                                   IsHitTestVisible="False" Focusable="False"
                                                   IsChecked="{Binding RelativeSource={RelativeSource AncestorType=DataGridRow}, Path=IsSelected, Mode=OneWay}"/>
                                    </DataTemplate>
                                </DataGridTemplateColumn.CellTemplate>
                            </DataGridTemplateColumn>
                            <DataGridTextColumn Header="Drive" Binding="{Binding DriveLetter}" Width="60"/>
                            <DataGridTextColumn Header="Volume Name" Binding="{Binding VolumeName}" Width="180"/>
                            <DataGridTextColumn Header="File System" Binding="{Binding FileSystem}" Width="100"/>
                            <DataGridTemplateColumn Header="Usage" Width="120">
                                <DataGridTemplateColumn.CellTemplate>
                                    <DataTemplate>
                                        <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                                            <ProgressBar Value="{Binding UsagePercent}" 
                                                        Width="60" Height="8" 
                                                        Background="#E0E0E0"
                                                        Foreground="{Binding UsageColor}"
                                                        Margin="0,0,8,0"/>
                                            <TextBlock Text="{Binding UsagePercentText}" 
                                                      VerticalAlignment="Center"
                                                      FontSize="10"
                                                      Foreground="{StaticResource TextSecondaryBrush}"/>
                                        </StackPanel>
                                    </DataTemplate>
                                </DataGridTemplateColumn.CellTemplate>
                            </DataGridTemplateColumn>
                            <DataGridTextColumn Header="Capacity (GB)" Binding="{Binding CapacityGB}" Width="100"/>
                            <DataGridTextColumn Header="Free Space (GB)" Binding="{Binding FreeSpaceGB}" Width="120"/>
                            <DataGridTextColumn Header="Used Space (GB)" Binding="{Binding UsedSpaceGB}" Width="120"/>
                        </DataGrid.Columns>
                    </DataGrid>
                </Grid>
            </TabItem>
            
            <!-- Shadow Copy Operations Tab -->
            <TabItem Style="{StaticResource ModernTabItemStyle}">
                <TabItem.Header>
                    <StackPanel Orientation="Horizontal">
                        <Image Source=".\assets\png\shadowcopy.png" Width="16" Height="16" Margin="0,0,8,0"/>
                        <TextBlock Text="Shadow Copies" VerticalAlignment="Center"/>
                    </StackPanel>
                </TabItem.Header>
                <Grid Margin="20">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>
                    
                    <!-- Shadow Copy Management Toolbar -->
                    <Border Grid.Row="0" Background="{StaticResource SurfaceBrush}" 
                            BorderBrush="{StaticResource BorderBrush}" BorderThickness="1"
                            CornerRadius="4" Padding="16" Margin="0,0,0,16">
                        <StackPanel>
                            <StackPanel Orientation="Horizontal" Margin="0,0,0,8">
                                <Button Name="btnRefreshShadowCopies" Style="{StaticResource ModernButtonStyle}"
                                        Width="100" Height="36" Margin="0,0,8,0">
                                    <StackPanel Orientation="Horizontal">
                                        <Image Source=".\assets\png\refresh.png" Width="16" Height="16" Margin="0,0,4,0"/>
                                        <TextBlock Text="Refresh" VerticalAlignment="Center"/>
                                    </StackPanel>
                                </Button>
                                <Button Name="btnCreateShadowCopy" Style="{StaticResource ModernButtonStyle}"
                                        Width="160" Height="36" Margin="0,0,8,0">
                                    <StackPanel Orientation="Horizontal">
                                        <Image Source=".\assets\png\create.png" Width="16" Height="16" Margin="0,0,4,0"/>
                                        <TextBlock Text="Create Shadow Copy" VerticalAlignment="Center"/>
                                    </StackPanel>
                                </Button>
                                <Button Name="btnDeleteShadowCopy" Style="{StaticResource ModernButtonStyle}"
                                        Width="140" Height="36">
                                    <StackPanel Orientation="Horizontal">
                                        <Image Source=".\assets\png\delete.png" Width="16" Height="16" Margin="0,0,4,0"/>
                                        <TextBlock Text="Delete Selected" VerticalAlignment="Center"/>
                                    </StackPanel>
                                </Button>
                            </StackPanel>
                            <TextBlock Name="txtSelectedVolume" Text="No volumes selected" 
                                      Foreground="{StaticResource TextSecondaryBrush}"
                                      FontStyle="Italic"/>
                        </StackPanel>
                    </Border>
                    
                    <!-- Shadow Copies DataGrid -->
                    <DataGrid Name="dgShadowCopies" Grid.Row="1" 
                              Style="{StaticResource ModernDataGridStyle}">
                        <DataGrid.ColumnHeaderStyle>
                            <Style TargetType="DataGridColumnHeader" BasedOn="{StaticResource ModernDataGridColumnHeaderStyle}"/>
                        </DataGrid.ColumnHeaderStyle>
                        <DataGrid.Columns>
                            <DataGridTextColumn Header="Volume" Binding="{Binding RequestedVolume}" Width="60"/>
                            <DataGridTextColumn Header="Shadow ID" Binding="{Binding ShadowID}" Width="250"/>
                            <DataGridTextColumn Header="Creation Time" Binding="{Binding CreationTimeText}" Width="160"/>
                            <DataGridTemplateColumn Header="State" Width="100">
                                <DataGridTemplateColumn.CellTemplate>
                                    <DataTemplate>
                                        <Border Background="{Binding StateColor}" 
                                                CornerRadius="12" 
                                                Padding="8,2" 
                                                HorizontalAlignment="Center">
                                            <TextBlock Text="{Binding State}" 
                                                      Foreground="White" 
                                                      FontSize="10" 
                                                      FontWeight="SemiBold"
                                                      HorizontalAlignment="Center"/>
                                        </Border>
                                    </DataTemplate>
                                </DataGridTemplateColumn.CellTemplate>
                            </DataGridTemplateColumn>
                            <DataGridTextColumn Header="No Writers" Binding="{Binding NoWriters}" Width="90"/>
                            <DataGridTextColumn Header="Size (MB)" Binding="{Binding SizeMB}" Width="100"/>
                            <DataGridTextColumn Header="Age" Binding="{Binding Age}" Width="80"/>
                        </DataGrid.Columns>
                    </DataGrid>
                </Grid>
            </TabItem>
        </TabControl>
        
        <!-- Status Bar -->
        <Border Grid.Row="2" Background="{StaticResource SurfaceBrush}" 
                BorderBrush="{StaticResource BorderBrush}" BorderThickness="0,1,0,0"
                Padding="20,8">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                    <TextBlock Name="txtStatus" Text="Ready" 
                              Foreground="{StaticResource TextSecondaryBrush}"
                              VerticalAlignment="Center"/>
                    <ProgressBar Name="progressBar" Width="200" Height="4" 
                                Margin="16,0,0,0" 
                                Background="{StaticResource BorderBrush}"
                                Foreground="{StaticResource PrimaryBrush}"
                                Visibility="Collapsed"/>
                </StackPanel>
                <TextBlock Grid.Column="1" Text="VSS Manager v1.0" 
                          Foreground="{StaticResource TextSecondaryBrush}"
                          VerticalAlignment="Center"/>
            </Grid>
        </Border>
    </Grid>
</Window>
"@

# Function to load XAML and create window
function Load-Xaml {
	# Get the script directory - try multiple methods to ensure we get the right path
	$scriptDir = $null
	
	# Method 1: Try MyInvocation (works when script is run directly)
	if ($MyInvocation.MyCommand.Path) {
		$scriptDir = Split-Path $MyInvocation.MyCommand.Path
		Write-Host "Using MyInvocation path: $scriptDir"
	}
	# Method 2: Try PSScriptRoot (works in most contexts)
	elseif ($PSScriptRoot) {
		$scriptDir = $PSScriptRoot
		Write-Host "Using PSScriptRoot path: $scriptDir"
	}
	# Method 3: Try to find the script in the current directory
	else {
		$currentDir = Get-Location
		$scriptPath = Join-Path $currentDir "VSSManager.ps1"
		if (Test-Path $scriptPath) {
			$scriptDir = $currentDir
			Write-Host "Using current directory path: $scriptDir"
		} else {
			Write-Warning "Could not determine script directory, using current location"
			$scriptDir = $currentDir
		}
	}
	
	$assetsPath = Join-Path $scriptDir "assets\png"
	
	Write-Host "Final script directory: $scriptDir"
	Write-Host "Assets path: $assetsPath"
	
	# Verify assets directory exists
	if (-not (Test-Path $assetsPath)) {
		Write-Warning "Assets directory not found: $assetsPath"
		Write-Host "Creating assets directory..."
		New-Item -ItemType Directory -Path $assetsPath -Force | Out-Null
	}
	
	# Check if PNG files exist and prepare safe XAML asset references.
	# Missing Image sources are non-fatal, but a missing Window.Icon causes XamlReader.Parse() to fail.
	$requiredFiles = @("refresh.png", "create.png", "delete.png", "volume.png", "shadowcopy.png", "app_icon.png")
	$xamlWithPaths = $xaml
	foreach ($file in $requiredFiles) {
		$filePath = Join-Path $assetsPath $file
		$relativeAssetPath = ".\assets\png\$file"
		$relativeAssetPattern = [regex]::Escape($relativeAssetPath)
		
		if (Test-Path $filePath) {
			Write-Host "Found PNG file: $filePath"
			$resolvedPath = $filePath
			$xamlWithPaths = [regex]::Replace($xamlWithPaths, $relativeAssetPattern, { param($match) $resolvedPath })
		}
		else {
			Write-Warning "Required PNG file not found: $filePath"
			if ($file -eq "app_icon.png") {
				$xamlWithPaths = $xamlWithPaths -replace '\s+Icon="\.\\assets\\png\\app_icon\.png"', ''
			}
			else {
				$xamlWithPaths = [regex]::Replace($xamlWithPaths, $relativeAssetPattern, { param($match) "" })
			}
		}
	}
	
	$window = [Windows.Markup.XamlReader]::Parse($xamlWithPaths)
	return $window
}

# Helper functions for progress indication
function Show-Progress {
	param([string]$Message, [bool]$Indeterminate = $true)
	$window.FindName("txtStatus").Text = $Message
	$progressBar = $window.FindName("progressBar")
	$progressBar.Visibility = "Visible"
	if ($Indeterminate) {
		$progressBar.IsIndeterminate = $true
	} else {
		$progressBar.IsIndeterminate = $false
		$progressBar.Value = 0
	}
	$window.UpdateLayout()
}

function Update-Progress {
	param([int]$Value, [string]$Message = "")
	if ($Message) {
		$window.FindName("txtStatus").Text = $Message
	}
	$progressBar = $window.FindName("progressBar")
	$progressBar.Value = $Value
	$window.UpdateLayout()
}

function Hide-Progress {
	param([string]$Message = "Ready")
	$window.FindName("txtStatus").Text = $Message
	$progressBar = $window.FindName("progressBar")
	$progressBar.Visibility = "Collapsed"
	$progressBar.IsIndeterminate = $false
	$window.UpdateLayout()
}

# Create the main window
$window = Load-Xaml

# Define data structures for binding
$volumeList = New-Object System.Collections.ObjectModel.ObservableCollection[object]
$shadowCopyList = New-Object System.Collections.ObjectModel.ObservableCollection[object]

# Set up data binding
$window.FindName("dgVolumes").ItemsSource = $volumeList
$window.FindName("dgShadowCopies").ItemsSource = $shadowCopyList

# Keep track of selected volumes across tabs (multi-select)
$script:selectedVolumes = @()
$window.FindName("dgVolumes").Add_SelectionChanged({
	param($sender, $args)
	$items = @($sender.SelectedItems)
	if ($items.Count -gt 0) {
		$script:selectedVolumes = $items
		$driveNames = ($items | ForEach-Object { $_.DriveLetter + " (" + $_.VolumeName + ")" }) -join ", "
		$window.FindName("txtSelectedVolume").Text = "Selected: " + $driveNames
		$window.FindName("txtStatus").Text = "$($items.Count) volume(s) selected: " + $driveNames
	} else {
		$script:selectedVolumes = @()
		$window.FindName("txtSelectedVolume").Text = "No volumes selected"
		$window.FindName("txtStatus").Text = "Ready"
	}
})

# Toggle row selection directly when the CheckBox column cell (DisplayIndex 0) is clicked,
# allowing multi-select without holding Ctrl, while preserving shift-select on other columns.
$window.FindName("dgVolumes").Add_PreviewMouseLeftButtonDown({
	param($sender, $e)
	
	try {
		$dep = $e.OriginalSource -as [System.Windows.DependencyObject]
		while ($null -ne $dep -and $dep -isnot [System.Windows.Controls.DataGridCell]) {
			$dep = [System.Windows.Media.VisualTreeHelper]::GetParent($dep)
		}
		
		if ($null -ne $dep) {
			$cell = $dep -as [System.Windows.Controls.DataGridCell]
			$colIdx = $sender.Columns.IndexOf($cell.Column)
			
			if ($colIdx -eq 0) {
				$rowDep = $cell
				while ($null -ne $rowDep -and $rowDep -isnot [System.Windows.Controls.DataGridRow]) {
					$rowDep = [System.Windows.Media.VisualTreeHelper]::GetParent($rowDep)
				}
				if ($null -ne $rowDep) {
					$row = $rowDep -as [System.Windows.Controls.DataGridRow]
					$row.IsSelected = -not $row.IsSelected
					$e.Handled = $true
				}
			}
		}
	} catch {
		# Silent failure
	}
})

# Event handlers for buttons
$window.FindName("btnRefreshVolumes").Add_Click({
	try {
		Show-Progress "Refreshing volumes..." $true
		
		$volumes = @(Get-VSSSupportedVolumes -ErrorAction Stop)
		$volumeList.Clear()
		
		$totalVolumes = $volumes.Count
		$currentVolume = 0
		
		foreach ($vol in $volumes) {
			$currentVolume++
			$progressPercent = if ($totalVolumes -gt 0) { [math]::Round(($currentVolume / $totalVolumes) * 100) } else { 100 }
			Update-Progress $progressPercent "Processing volume $currentVolume of $totalVolumes"
			$volumeList.Add($vol)
		}
		
		Hide-Progress "Volumes refreshed successfully - Found $($volumeList.Count) volumes"
		$window.Title = "VSS Manager - Volume Shadow Copy Service"
	}
	catch {
		Hide-Progress "Error refreshing volumes"
		[System.Windows.MessageBox]::Show("Error refreshing volumes: $($_.Exception.Message)", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
	}
})

$window.FindName("btnRefreshShadowCopies").Add_Click({
	try {
		$volumes = @($script:selectedVolumes)
		if ($volumes.Count -eq 0) {
			$sel = $window.FindName("dgVolumes").SelectedItem
			if ($sel) { $volumes = @($sel) }
		}
		if ($volumes.Count -gt 0) {
			Show-Progress "Refreshing shadow copies for $($volumes.Count) volume(s)..." $true
			$shadowCopyList.Clear()
			
			$totalVolumes = $volumes.Count
			$currentVolumeIdx = 0
			
			foreach ($vol in $volumes) {
				$currentVolumeIdx++
				Update-Progress ([math]::Round(($currentVolumeIdx / $totalVolumes) * 100)) "Querying shadow copies for $($vol.DriveLetter)..."
				$shadowCopies = @(Get-VSSShadowCopies -VolumePath $vol.DeviceID -ErrorAction Stop)
				foreach ($copy in $shadowCopies) {
					$copy | Add-Member -NotePropertyName RequestedVolume -NotePropertyValue $vol.DriveLetter -Force
					$shadowCopyList.Add($copy)
				}
			}
			
			Hide-Progress "Shadow copies refreshed - Found $($shadowCopyList.Count) copies across $($volumes.Count) volume(s)"
		} else {
			Hide-Progress "Please select at least one volume first"
			[System.Windows.MessageBox]::Show("Please select at least one volume first", "Information", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
		}
	}
	catch {
		Hide-Progress "Error refreshing shadow copies"
		[System.Windows.MessageBox]::Show("Error refreshing shadow copies: $($_.Exception.Message)", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
	}
})

$window.FindName("btnCreateShadowCopy").Add_Click({
	try {
		$volumes = @($script:selectedVolumes)
		if ($volumes.Count -eq 0) {
			$sel = $window.FindName("dgVolumes").SelectedItem
			if ($sel) { $volumes = @($sel) }
		}
		if ($volumes.Count -gt 0) {
			$driveNames = ($volumes | ForEach-Object { $_.DriveLetter }) -join ", "
			Show-Progress "Creating shadow copies for $($volumes.Count) volume(s)..." $true
			
			$totalVolumes = $volumes.Count
			$currentVolumeIdx = 0
			$successCount = 0
			$errors = @()
			
			foreach ($vol in $volumes) {
				$currentVolumeIdx++
				Update-Progress ([math]::Round(($currentVolumeIdx / $totalVolumes) * 100)) "Creating shadow copy for $($vol.DriveLetter) ($currentVolumeIdx of $totalVolumes)..."
				try {
					$created = New-VSSShadowCopy -VolumePath $vol.DeviceID -Confirm:$false -ErrorAction Stop
					if ($created -and $created.Success) {
						$successCount++
					} else {
						$errors += "$($vol.DriveLetter): Creation did not return a successful result."
					}
				} catch {
					$errors += "$($vol.DriveLetter): $($_.Exception.Message)"
				}
			}
			
			if ($errors.Count -gt 0) {
				Hide-Progress "Shadow copies created: $successCount succeeded, $($errors.Count) failed"
				[System.Windows.MessageBox]::Show("Completed with errors:`n`n" + ($errors -join "`n"), "Partial Failure", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
			} else {
				Hide-Progress "Shadow copies created successfully for $successCount volume(s)"
			}
			# Refresh list after creation
			$window.FindName("btnRefreshShadowCopies").RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent)))
		} else {
			Hide-Progress "Please select at least one volume first"
			[System.Windows.MessageBox]::Show("Please select at least one volume first", "Information", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
		}
	}
	catch {
		Hide-Progress "Error creating shadow copies"
		[System.Windows.MessageBox]::Show("Error creating shadow copies: $($_.Exception.Message)", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
	}
})

$window.FindName("btnDeleteShadowCopy").Add_Click({
	try {
		$selectedCopies = @($window.FindName("dgShadowCopies").SelectedItems)
		if ($selectedCopies.Count -gt 0) {
			$confirmMessage = ""
			if ($selectedCopies.Count -eq 1) {
				$copy = $selectedCopies[0]
				$displayVolume = if ($copy.RequestedVolume) { $copy.RequestedVolume } else { $copy.VolumePath }
				$confirmMessage = "Are you sure you want to delete this shadow copy?`n`nVolume: $displayVolume`nID: $($copy.ShadowID)"
			} else {
				$confirmMessage = "Are you sure you want to delete these $($selectedCopies.Count) selected shadow copies?"
			}
			
			$result = [System.Windows.MessageBox]::Show($confirmMessage, "Confirm Deletion", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Question)
			
			if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
				Show-Progress "Deleting shadow copies..." $true
				
				$totalCopies = $selectedCopies.Count
				$currentCopyIdx = 0
				$successCount = 0
				$errors = @()
				
				foreach ($copy in $selectedCopies) {
					$currentCopyIdx++
					Update-Progress ([math]::Round(($currentCopyIdx / $totalCopies) * 100)) "Deleting shadow copy $currentCopyIdx of $totalCopies..."
					try {
						$deleted = Remove-VSSShadowCopy -VolumePath $copy.VolumePath -ShadowCopyID $copy.ShadowID -Confirm:$false -ErrorAction Stop
						if ($deleted -and $deleted.Success) {
							$successCount++
						} else {
							$errors += "ID $($copy.ShadowID): Deletion did not return a successful result."
						}
					} catch {
						$errors += "ID $($copy.ShadowID): $($_.Exception.Message)"
					}
				}
				
				if ($errors.Count -gt 0) {
					Hide-Progress "Shadow copies deleted: $successCount succeeded, $($errors.Count) failed"
					[System.Windows.MessageBox]::Show("Completed with errors:`n`n" + ($errors -join "`n"), "Partial Failure", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
				} else {
					Hide-Progress "Deleted $successCount shadow copies successfully"
				}
				
				# Refresh list after deletion
				$window.FindName("btnRefreshShadowCopies").RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent)))
			}
		} else {
			Hide-Progress "Please select at least one shadow copy to delete"
			[System.Windows.MessageBox]::Show("Please select at least one shadow copy to delete from the list", "Information", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
		}
	}
	catch {
		Hide-Progress "Error deleting shadow copies"
		[System.Windows.MessageBox]::Show("Error deleting shadow copies: $($_.Exception.Message)", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
	}
})

# Show the window
$window.ShowDialog() | Out-Null