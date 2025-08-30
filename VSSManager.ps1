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
            <Setter Property="SelectionMode" Value="Single"/>
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
                                                  IsChecked="{Binding RelativeSource={RelativeSource AncestorType=DataGridRow}, Path=IsSelected, Mode=TwoWay}"/>
                                    </DataTemplate>
                                </DataGridTemplateColumn.CellTemplate>
                            </DataGridTemplateColumn>
                            <DataGridTextColumn Header="Drive" Binding="{Binding DriveLetter}" Width="80"/>
                            <DataGridTextColumn Header="Volume Name" Binding="{Binding VolumeName}" Width="200"/>
                            <DataGridTextColumn Header="File System" Binding="{Binding FileSystem}" Width="120"/>
                            <DataGridTextColumn Header="Capacity (GB)" Binding="{Binding CapacityGB}" Width="120"/>
                            <DataGridTextColumn Header="Free Space (GB)" Binding="{Binding FreeSpaceGB}" Width="140"/>
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
                            <TextBlock Name="txtSelectedVolume" Text="No volume selected" 
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
                            <DataGridTextColumn Header="Shadow ID" Binding="{Binding ShadowID}" Width="300"/>
                            <DataGridTextColumn Header="Creation Time" Binding="{Binding CreationTime}" Width="180"/>
                            <DataGridTextColumn Header="No Writers" Binding="{Binding NoWriters}" Width="100"/>
                            <DataGridTextColumn Header="State" Binding="{Binding State}" Width="120"/>
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
	# Get the script directory and create absolute paths for images
	$scriptDir = Split-Path $MyInvocation.MyCommand.Path
	$assetsPath = Join-Path $scriptDir "assets\png"
	
	# Replace relative paths with absolute paths in XAML
	$xamlWithPaths = $xaml -replace '\.\\assets\\png\\', "$assetsPath\"
	
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

# Keep track of selected volume across tabs
$script:selectedVolumeDeviceId = $null
$script:selectedVolumeName = $null
$window.FindName("dgVolumes").Add_SelectionChanged({
	param($sender, $args)
	$sel = $sender.SelectedItem
	if ($sel) {
		$script:selectedVolumeDeviceId = $sel.DeviceID
		$script:selectedVolumeName = $sel.DriveLetter + " (" + $sel.VolumeName + ")"
		$window.FindName("txtSelectedVolume").Text = "Selected: " + $script:selectedVolumeName
		$window.FindName("txtStatus").Text = "Volume selected: " + $script:selectedVolumeName
	} else {
		$script:selectedVolumeDeviceId = $null
		$script:selectedVolumeName = $null
		$window.FindName("txtSelectedVolume").Text = "No volume selected"
		$window.FindName("txtStatus").Text = "Ready"
	}
})

# Event handlers for buttons
$window.FindName("btnRefreshVolumes").Add_Click({
	try {
		Show-Progress "Refreshing volumes..." $true
		
		# Call your existing function to get volumes
		$volumes = Get-WmiObject -Class Win32_Volume | Where-Object { $_.DriveType -eq 3 -and $_.DriveLetter -ne $null }
		$volumeList.Clear()
		
		$totalVolumes = $volumes.Count
		$currentVolume = 0
		
		foreach ($vol in $volumes) {
			$currentVolume++
			$progressPercent = [math]::Round(($currentVolume / $totalVolumes) * 100)
			Update-Progress $progressPercent "Processing volume $currentVolume of $totalVolumes"
			
			$volumeInfo = [PSCustomObject]@{
				DriveLetter = $vol.DriveLetter
				VolumeName = if ($vol.Label) { $vol.Label } else { "Local Disk" }
				FileSystem = $vol.FileSystem
				CapacityGB = [math]::Round($vol.Capacity / 1GB, 2)
				FreeSpaceGB = [math]::Round($vol.FreeSpace / 1GB, 2)
				DeviceID = $vol.DeviceID
			}
			$volumeList.Add($volumeInfo)
		}
		
		Hide-Progress "Volumes refreshed successfully - Found $($volumeList.Count) volumes"
		$window.Title = "VSS Manager - Volume Shadow Copy Service"
	}
	catch {
		Hide-Progress "Error refreshing volumes"
		[System.Windows.MessageBox]::Show("Error refreshing volumes: $_", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
	}
})

$window.FindName("btnRefreshShadowCopies").Add_Click({
	try {
		$volId = $script:selectedVolumeDeviceId
		if (-not $volId) {
			$sel = $window.FindName("dgVolumes").SelectedItem
			if ($sel) { $volId = $sel.DeviceID }
		}
		if ($volId) {
			Show-Progress "Refreshing shadow copies..." $true
			
			# Retrieve shadow copies for the selected volume (robust matching for trailing backslash differences)
			$shadowCopies = Get-WmiObject -Class Win32_ShadowCopy | Where-Object { $_.VolumeName -like "*$volId*" }
			$shadowCopyList.Clear()
			
			$totalCopies = $shadowCopies.Count
			$currentCopy = 0
			
			foreach ($copy in $shadowCopies) {
				$currentCopy++
				$progressPercent = [math]::Round(($currentCopy / $totalCopies) * 100)
				Update-Progress $progressPercent "Processing shadow copy $currentCopy of $totalCopies"
				
				$created = try { 
					[System.Management.ManagementDateTimeConverter]::ToDateTime($copy.InstallDate).ToString("yyyy-MM-dd HH:mm:ss")
				} catch { 
					$copy.InstallDate 
				}
				$shadowInfo = [PSCustomObject]@{
					ShadowID = $copy.ID
					CreationTime = $created
					NoWriters = $copy.NoWriters
					State = $copy.State
					VolumePath = $copy.VolumeName
				}
				$shadowCopyList.Add($shadowInfo)
			}
			
			Hide-Progress "Shadow copies refreshed - Found $($shadowCopyList.Count) copies"
		} else {
			Hide-Progress "Please select a volume first"
			[System.Windows.MessageBox]::Show("Please select a volume first", "Information", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
		}
	}
	catch {
		Hide-Progress "Error refreshing shadow copies"
		[System.Windows.MessageBox]::Show("Error refreshing shadow copies: $_", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
	}
})

$window.FindName("btnCreateShadowCopy").Add_Click({
	try {
		$volId = $script:selectedVolumeDeviceId
		if (-not $volId) {
			$sel = $window.FindName("dgVolumes").SelectedItem
			if ($sel) { $volId = $sel.DeviceID }
		}
		if ($volId) {
			Show-Progress "Creating shadow copy..." $true
			
			# Simulate progress for shadow copy creation
			Update-Progress 25 "Initializing shadow copy creation..."
			Start-Sleep -Milliseconds 500
			
			Update-Progress 50 "Creating shadow copy..."
			New-VSSShadowCopy -VolumePath $volId | Out-Null
			
			Update-Progress 75 "Finalizing shadow copy..."
			Start-Sleep -Milliseconds 300
			
			Update-Progress 100 "Shadow copy created successfully"
			Start-Sleep -Milliseconds 200
			
			Hide-Progress "Shadow copy created successfully"
			# Refresh list after creation
			$window.FindName("btnRefreshShadowCopies").RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent)))
		} else {
			Hide-Progress "Please select a volume first"
			[System.Windows.MessageBox]::Show("Please select a volume first", "Information", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
		}
	}
	catch {
		Hide-Progress "Error creating shadow copy"
		[System.Windows.MessageBox]::Show("Error creating shadow copy: $_", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
	}
})

$window.FindName("btnDeleteShadowCopy").Add_Click({
	try {
		$selectedCopy = $window.FindName("dgShadowCopies").SelectedItem
		$volId = $script:selectedVolumeDeviceId
		if (-not $volId) {
			$sel = $window.FindName("dgVolumes").SelectedItem
			if ($sel) { $volId = $sel.DeviceID }
		}
		if ($selectedCopy -and $volId) {
			$result = [System.Windows.MessageBox]::Show("Are you sure you want to delete this shadow copy?", "Confirm Deletion", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Question)
			
			if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
				Show-Progress "Deleting shadow copy..." $true
				
				# Simulate progress for shadow copy deletion
				Update-Progress 25 "Preparing for deletion..."
				Start-Sleep -Milliseconds 300
				
				Update-Progress 50 "Deleting shadow copy..."
				Remove-VSSShadowCopy -VolumePath $volId -ShadowCopyID $selectedCopy.ShadowID -Confirm:$false | Out-Null
				
				Update-Progress 75 "Cleaning up..."
				Start-Sleep -Milliseconds 200
				
				Update-Progress 100 "Shadow copy deleted successfully"
				Start-Sleep -Milliseconds 200
				
				Hide-Progress "Shadow copy deleted successfully"
				# Refresh list after deletion
				$window.FindName("btnRefreshShadowCopies").RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent)))
			}
		} else {
			Hide-Progress "Please select a shadow copy to delete"
			[System.Windows.MessageBox]::Show("Please select a shadow copy to delete and ensure a volume is selected", "Information", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
		}
	}
	catch {
		Hide-Progress "Error deleting shadow copy"
		[System.Windows.MessageBox]::Show("Error deleting shadow copy: $_", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
	}
})

# Show the window
$window.ShowDialog() | Out-Null