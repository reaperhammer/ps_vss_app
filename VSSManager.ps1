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
        Title="Volume Shadow Copy Manager" Height="600" Width="800"
        WindowStartupLocation="CenterScreen">
    <Grid>
        <TabControl>
            <!-- Volume Management Tab -->
            <TabItem Header="Volume Management">
                <Grid Margin="10">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>
                    <Button Name="btnRefreshVolumes" Content="Refresh Volumes" 
                            HorizontalAlignment="Left" Width="120" Height="30" 
                            Margin="10,10,10,10"/>
                    <DataGrid Name="dgVolumes" Grid.Row="1" AutoGenerateColumns="False" 
                              CanUserAddRows="False" SelectionMode="Single">
                        <DataGrid.Columns>
                            <DataGridTemplateColumn Header="" Width="40">
                                <DataGridTemplateColumn.CellTemplate>
                                    <DataTemplate>
                                        <CheckBox HorizontalAlignment="Center" VerticalAlignment="Center"
                                                  IsChecked="{Binding RelativeSource={RelativeSource AncestorType=DataGridRow}, Path=IsSelected, Mode=TwoWay}"/>
                                    </DataTemplate>
                                </DataGridTemplateColumn.CellTemplate>
                            </DataGridTemplateColumn>
                            <DataGridTextColumn Header="Drive Letter" Binding="{Binding DriveLetter}" Width="100"/>
                            <DataGridTextColumn Header="Volume Name" Binding="{Binding VolumeName}" Width="150"/>
                            <DataGridTextColumn Header="File System" Binding="{Binding FileSystem}" Width="120"/>
                            <DataGridTextColumn Header="Capacity (GB)" Binding="{Binding CapacityGB}" Width="120"/>
                        </DataGrid.Columns>
                    </DataGrid>
                </Grid>
            </TabItem>
            
            <!-- Shadow Copy Operations Tab -->
            <TabItem Header="Shadow Copies">
                <Grid Margin="10">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>
                    <StackPanel Orientation="Horizontal" Margin="0,0,0,10">
                        <Button Name="btnRefreshShadowCopies" Content="Refresh Shadow Copies" 
                                Width="150" Height="30" Margin="5"/>
                        <Button Name="btnCreateShadowCopy" Content="Create Shadow Copy" 
                                Width="150" Height="30" Margin="5"/>
                        <Button Name="btnDeleteShadowCopy" Content="Delete Selected" 
                                Width="150" Height="30" Margin="5"/>
                    </StackPanel>
                    <DataGrid Name="dgShadowCopies" Grid.Row="2" AutoGenerateColumns="False" 
                              CanUserAddRows="False">
                        <DataGrid.Columns>
                            <DataGridTextColumn Header="Shadow ID" Binding="{Binding ShadowID}" Width="200"/>
                            <DataGridTextColumn Header="Creation Time" Binding="{Binding CreationTime}" Width="200"/>
                            <DataGridTextColumn Header="Description" Binding="{Binding Description}" Width="200"/>
                        </DataGrid.Columns>
                    </DataGrid>
                </Grid>
            </TabItem>
        </TabControl>
    </Grid>
</Window>
"@

# Function to load XAML and create window
function Load-Xaml {
	$window = [Windows.Markup.XamlReader]::Parse($xaml)
	return $window
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
$window.FindName("dgVolumes").Add_SelectionChanged({
	param($sender, $args)
	$sel = $sender.SelectedItem
	if ($sel) {
		$script:selectedVolumeDeviceId = $sel.DeviceID
	}
})

# Event handlers for buttons
$window.FindName("btnRefreshVolumes").Add_Click({
	try {
		# Call your existing function to get volumes
		$volumes = Get-WmiObject -Class Win32_Volume | Where-Object { $_.DriveType -eq 3 -and $_.DriveLetter -ne $null }
		$volumeList.Clear()
		
		foreach ($vol in $volumes) {
			$volumeInfo = [PSCustomObject]@{
				DriveLetter = $vol.DriveLetter
				VolumeName = $vol.Label
				FileSystem = $vol.FileSystem
				CapacityGB = [math]::Round($vol.Capacity / 1GB, 2)
				DeviceID = $vol.DeviceID
			}
			$volumeList.Add($volumeInfo)
		}
		
		$window.Title = "Volume Shadow Copy Manager - Volumes Refreshed"
	}
	catch {
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
			# Retrieve shadow copies for the selected volume (robust matching for trailing backslash differences)
			$shadowCopies = Get-WmiObject -Class Win32_ShadowCopy | Where-Object { $_.VolumeName -like "*$volId*" }
			$shadowCopyList.Clear()
			
			foreach ($copy in $shadowCopies) {
				$created = try { [System.Management.ManagementDateTimeConverter]::ToDateTime($copy.InstallDate) } catch { $copy.InstallDate }
				$shadowInfo = [PSCustomObject]@{
					ShadowID = $copy.ID
					CreationTime = $created
					Description = $copy.Description
					VolumePath = $copy.VolumeName
				}
				$shadowCopyList.Add($shadowInfo)
			}
			
			$window.Title = "Volume Shadow Copy Manager - Shadow Copies Refreshed"
		} else {
			[System.Windows.MessageBox]::Show("Please select a volume first", "Information", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
		}
	}
	catch {
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
			New-VSSShadowCopy -VolumePath $volId | Out-Null
			# Refresh list after creation
			$window.FindName("btnRefreshShadowCopies").RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent)))
		} else {
			[System.Windows.MessageBox]::Show("Please select a volume first", "Information", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
		}
	}
	catch {
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
				Remove-VSSShadowCopy -VolumePath $volId -ShadowCopyID $selectedCopy.ShadowID -Confirm:$false | Out-Null
				# Refresh list after deletion
				$window.FindName("btnRefreshShadowCopies").RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent)))
			}
		} else {
			[System.Windows.MessageBox]::Show("Please select a shadow copy to delete and ensure a volume is selected", "Information", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
		}
	}
	catch {
		[System.Windows.MessageBox]::Show("Error deleting shadow copy: $_", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
	}
})

# Show the window
$window.ShowDialog() | Out-Null