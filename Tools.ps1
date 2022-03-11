#Start Executing Script
$path = Get-Location
Copy-Item -Path "$path\*" -Destination "C:\Automation\img"
#GUI Design
$inputXML = @"
<Window
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:PSGui"
        mc:Ignorable="d"
        Title="Client Site Tools" Height="630" Width="550" FontSize="24" HorizontalAlignment="Center" VerticalAlignment="Center" WindowStartupLocation = "CenterScreen">
    <Grid MaxWidth="600" MaxHeight="700" MinWidth="500" MinHeight="600" Height="700" Width="600" HorizontalAlignment="Center" VerticalAlignment="Center">
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="11*"/>
            <ColumnDefinition Width="89*"/>
        </Grid.ColumnDefinitions>
        <Image HorizontalAlignment="Left" Height="150" Margin="50,62,0,0" VerticalAlignment="Top" Width="150" Source="C:\Automation\img\main.ico" Grid.ColumnSpan="2"/>
        <TextBlock HorizontalAlignment="Left" Height="130" Margin="165,62,0,0" Text="Client Site Tools used for fixing issues or optimizing device" TextWrapping="Wrap" VerticalAlignment="Top" Width="325" FontSize="32" Grid.Column="1"/>
        <TabControl HorizontalAlignment="Left" Height="400" Margin="50,232,0,0" VerticalAlignment="Top" Width="504" Grid.ColumnSpan="2">
            <TabItem Header="System Information" FontSize="24" Margin="-2,0,-23,0">
                <Grid Background="#FFE5E5E5" >
                    <Label Content="Username" HorizontalAlignment="Left" Margin="10,43,0,0" VerticalAlignment="Top" Width="200" Height="42" Padding="0,0,0,0"/>
                    <Label Content="Computer Name" HorizontalAlignment="Left" Margin="10,103,0,0" VerticalAlignment="Top" Width="200" Height="42" Padding="0,0,0,0"/>
                    <Label Content="Windows Version" HorizontalAlignment="Left" Margin="10,163,0,0" VerticalAlignment="Top" Width="200" Height="42" Padding="0,0,0,0"/>
                    <Label Content="Available Space C:" HorizontalAlignment="Left" Margin="10,223,0,0" VerticalAlignment="Top" Width="200" Height="42" Padding="0,0,0,0"/>
                    <Label Content="Serial Number" HorizontalAlignment="Left" Margin="10,283,0,0" VerticalAlignment="Top" Width="200" Height="42" Padding="0,0,0,0"/>
                    <TextBox Name="text_UN" HorizontalAlignment="Left" Height="42" Margin="255,40,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="233"/>
                    <TextBox Name="text_CN" HorizontalAlignment="Left" Height="42" Margin="255,100,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="233"/>
                    <TextBox Name="text_WV" HorizontalAlignment="Left" Height="42" Margin="255,160,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="233"/>
                    <TextBox x:Name="text_AC" HorizontalAlignment="Left" Height="42" Margin="255,220,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="233"/>
                    <TextBox x:Name="text_SN" HorizontalAlignment="Left" Height="42" Margin="255,280,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="233"/>
                </Grid>
            </TabItem>
            <TabItem Header="Tools" FontSize="24" Margin="32,0,-158,0">
                <ListView HorizontalAlignment="Left" Height="340" Margin="10,10,0,10" VerticalAlignment="Top" Width="480" Background="#FFE5E5E5">
                    <ListView.View>
                        <GridView>
                            <GridViewColumn Header="List" Width="450"  />
                        </GridView>
                    </ListView.View>
                    <Button Name="T_00" HorizontalAlignment="Center" Height="36" Margin="0,5,0,5" VerticalAlignment="Center" Width="36" BorderBrush="Black" Foreground="Red" FontSize="16" FontWeight="Bold">
                        <Button.Background>
                            <ImageBrush ImageSource="C:\Automation\img\if_0.ico" Stretch="UniformToFill"/>
                        </Button.Background>
                    </Button>
                    <Label Content="Set Date Manually" HorizontalAlignment="Center" Height="38" Margin="40,-60,0,0" VerticalAlignment="Center" Width="320" FontSize="22"/>
                    
                    <Button Name="T_01" HorizontalAlignment="Center" Height="36" Margin="0,5,0,5" VerticalAlignment="Center" Width="36" BorderBrush="Black" Foreground="Red" FontSize="16" FontWeight="Bold">
                        <Button.Background>
                            <ImageBrush ImageSource="C:\Automation\img\if_0.ico" Stretch="UniformToFill"/>
                        </Button.Background>
                    </Button>
                    <Label Content="Set Time Manually" HorizontalAlignment="Center" Height="38" Margin="40,-60,0,0" VerticalAlignment="Center" Width="320" FontSize="22"/>

                    <Button Name="T_02" HorizontalAlignment="Center" Height="36" Margin="0,5,0,5" VerticalAlignment="Center" Width="36" BorderBrush="Black" Foreground="Red" FontSize="16" FontWeight="Bold">
                        <Button.Background>
                            <ImageBrush ImageSource="C:\Automation\img\if_0.ico" Stretch="UniformToFill"/>
                        </Button.Background>
                    </Button>
                    <Label Content="ddMMyyyy Format" HorizontalAlignment="Center" Height="38" Margin="40,-60,0,0" VerticalAlignment="Center" Width="320" FontSize="22"/>

                    <Button Name="T_03" HorizontalAlignment="Center" Height="36" Margin="0,5,0,5" VerticalAlignment="Center" Width="36" BorderBrush="Black" Foreground="Red" FontSize="16" FontWeight="Bold">
                        <Button.Background>
                            <ImageBrush ImageSource="C:\Automation\img\if_0.ico" Stretch="UniformToFill"/>
                        </Button.Background>
                    </Button>
                    <Label Content="Device Management" HorizontalAlignment="Center" Height="38" Margin="40,-60,0,0" VerticalAlignment="Center" Width="320" FontSize="22"/>

                    <Button Name="T_04" HorizontalAlignment="Center" Height="36" Margin="0,5,0,5" VerticalAlignment="Center" Width="36" BorderBrush="Black" Foreground="Red" FontSize="16" FontWeight="Bold">
                        <Button.Background>
                            <ImageBrush ImageSource="C:\Automation\img\if_0.ico" Stretch="UniformToFill"/>
                        </Button.Background>
                    </Button>
                    <Label Content="Disk Management" HorizontalAlignment="Center" Height="38" Margin="40,-60,0,0" VerticalAlignment="Center" Width="320" FontSize="22"/>

                    <Button Name="T_05" HorizontalAlignment="Center" Height="36" Margin="0,5,0,5" VerticalAlignment="Center" Width="36" BorderBrush="Black" Foreground="Red" FontSize="16" FontWeight="Bold">
                        <Button.Background>
                            <ImageBrush ImageSource="C:\Automation\img\if_0.ico" Stretch="UniformToFill"/>
                        </Button.Background>
                    </Button>
                    <Label Content="Quick Assist" HorizontalAlignment="Center" Height="38" Margin="40,-60,0,0" VerticalAlignment="Center" Width="320" FontSize="22"/>

                    <Button Name="T_06" HorizontalAlignment="Center" Height="36" Margin="0,5,0,5" VerticalAlignment="Center" Width="36" BorderBrush="Black" Foreground="Red" FontSize="16" FontWeight="Bold">
                        <Button.Background>
                            <ImageBrush ImageSource="C:\Automation\img\if_0.ico" Stretch="UniformToFill"/>
                        </Button.Background>
                    </Button>
                    <Label Content="Reset Network (Restart Required)" HorizontalAlignment="Center" Height="38" Margin="40,-60,0,0" VerticalAlignment="Center" Width="330" FontSize="22"/>

                    <Button Name="T_07" HorizontalAlignment="Center" Height="36" Margin="0,5,0,5" VerticalAlignment="Center" Width="36" BorderBrush="Black" Foreground="Red" FontSize="16" FontWeight="Bold">
                        <Button.Background>
                            <ImageBrush ImageSource="C:\Automation\img\if_0.ico" Stretch="UniformToFill"/>
                        </Button.Background>
                    </Button>
                    <Label Content="Wi-Fi Enabled" HorizontalAlignment="Center" Height="38" Margin="40,-60,0,0" VerticalAlignment="Center" Width="330" FontSize="22"/>

                    <Button Name="T_08" HorizontalAlignment="Center" Height="36" Margin="0,5,0,5" VerticalAlignment="Center" Width="36" BorderBrush="Black" Foreground="Red" FontSize="16" FontWeight="Bold">
                        <Button.Background>
                            <ImageBrush ImageSource="C:\Automation\img\if_0.ico" Stretch="UniformToFill"/>
                        </Button.Background>
                    </Button>
                    <Label Content="Force Shutdown" HorizontalAlignment="Center" Height="38" Margin="40,-60,0,0" VerticalAlignment="Center" Width="330" FontSize="22"/>

                    <Button Name="T_09" HorizontalAlignment="Center" Height="36" Margin="0,5,0,5" VerticalAlignment="Center" Width="36" BorderBrush="Black" Foreground="Red" FontSize="16" FontWeight="Bold">
                        <Button.Background>
                            <ImageBrush ImageSource="C:\Automation\img\if_0.ico" Stretch="UniformToFill"/>
                        </Button.Background>
                    </Button>
                    <Label Content="Update Policy" HorizontalAlignment="Center" Height="38" Margin="40,-60,0,0" VerticalAlignment="Center" Width="330" FontSize="22"/>

                    <Button Name="T_10" HorizontalAlignment="Center" Height="36" Margin="0,5,0,5" VerticalAlignment="Center" Width="36" BorderBrush="Black" Foreground="Red" FontSize="16" FontWeight="Bold">
                        <Button.Background>
                            <ImageBrush ImageSource="C:\Automation\img\if_0.ico" Stretch="UniformToFill"/>
                        </Button.Background>
                    </Button>
                    <Label Content="Disk Clean" HorizontalAlignment="Center" Height="38" Margin="40,-60,0,0" VerticalAlignment="Center" Width="330" FontSize="22"/>

                    <Button Name="T_11" HorizontalAlignment="Center" Height="36" Margin="0,5,0,5" VerticalAlignment="Center" Width="36" BorderBrush="Black" Foreground="Red" FontSize="16" FontWeight="Bold">
                        <Button.Background>
                            <ImageBrush ImageSource="C:\Automation\img\if_0.ico" Stretch="UniformToFill"/>
                        </Button.Background>
                    </Button>
                    <Label Content="Check Disk (Restart Required)" HorizontalAlignment="Center" Height="38" Margin="40,-60,0,0" VerticalAlignment="Center" Width="330" FontSize="22"/>

                </ListView>
            </TabItem>
        </TabControl>
    </Grid>
</Window>
"@ 
 
$inputXML = $inputXML -replace 'mc:Ignorable="d"', '' -replace "x:N", 'N' -replace '^<Win.*', '<Window'
[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
[xml]$XAML = $inputXML
 
$reader = (New-Object System.Xml.XmlNodeReader $xaml)
try {
    $Form = [Windows.Markup.XamlReader]::Load( $reader )
}
catch {
    Write-Warning "Unable to parse XML, with error: $($Error[0])`n Ensure that there are NO SelectionChanged or TextChanged properties in your textboxes (PowerShell cannot process them)"
    throw
}

$xaml.SelectNodes("//*[@Name]") | % { "trying item $($_.Name)";
    try { Set-Variable -Name "WPF$($_.Name)" -Value $Form.FindName($_.Name) -ErrorAction Stop }
    catch { throw }
}

#Function Execution
$WPFtext_UN.Text = $env:UserName
$WPFtext_CN.Text = $(Get-WmiObject Win32_ComputerSystem).name

$v = (Get-WmiObject Win32_operatingSystem).BuildNumber
switch ($V) {
    14393 { $WPFtext_WV.Text = "1607"; break }
    15063 { $WPFtext_WV.Text = "1703"; break }
    16299 { $WPFtext_WV.Text = "1709"; break }
    17134 { $WPFtext_WV.Text = "1803"; break }
    17763 { $WPFtext_WV.Text = "1809"; break }
    18362 { $WPFtext_WV.Text = "1903"; break }
    18363 { $WPFtext_WV.Text = "1909"; break }
    19041 { $WPFtext_WV.Text = "2004"; break }
    19042 { $WPFtext_WV.Text = "21H2"; break }
    19043 { $WPFtext_WV.Text = "21H1"; break }
    19044 { $WPFtext_WV.Text = "21H2"; break }
}

$username = ""
$password = ""
$secstr = New-Object -TypeName System.Security.SecureString
$password.ToCharArray() | ForEach-Object { $secstr.AppendChar($_) }
$credential = new-object -typename System.Management.Automation.PSCredential -argumentlist $username, $secstr

$c = Get-CimInstance Win32_LogicalDisk -Filter "DeviceID = 'C:'"
$C = [int]($c.FreeSpace / 1GB)
$WPFtext_AC.Text = "$($C) GB"

$WPFtext_SN.Text = $(Get-WmiObject Win32_Bios).SerialNumber

$script = $PSCommandPath
Set-Location $script
<# Set Date Manually #>
$WPFT_00.Add_Click({
    Start-Process -FilePath "source\WMo08N4W\T_00.bat" 
})
<# Set Time Manually #>
$WPFT_01.Add_Click({
    Start-Process -FilePath "source\WMo08N4W\T_01.bat" 
})
<# ddMMyyyy #>
$WPFT_02.Add_Click({
    Start-Process -FilePath "source\WMo08N4W\T_02.bat"
})
<# Device Management #>
$WPFT_03.Add_Click({
    Start-Process -FilePath "source\WMo08N4W\T_03.bat" 
})
<# Disk Management #>
$WPFT_04.Add_Click({
    Start-Process -FilePath "source\WMo08N4W\T_04.bat" 
})
<# Quick Assist #>
$WPFT_05.Add_Click({
    Start-Process -FilePath "source\WMo08N4W\T_05.bat"
})
<# Reset Network #>
$WPFT_06.Add_Click({
    Start-Process -FilePath "source\WMo08N4W\T_06.bat" 
})
<# Wi-Fi Enabled #>
$WPFT_07.Add_Click({
    Start-Process -FilePath "source\WMo08N4W\T_07.bat" 
})
<# Force Shutdown #>
$WPFT_08.Add_Click({
    Start-Process -FilePath "source\WMo08N4W\T_08.bat"
})
<# Update Policy #>
$WPFT_09.Add_Click({
    Start-Process -FilePath "source\WMo08N4W\T_09.bat"
})
<# Disk Clean #>
$WPFT_10.Add_Click({
    Start-Process -FilePath "source\WMo08N4W\T_10.bat" 
})
<# Check Disk #>
$WPFT_11.Add_Click({
    Start-Process -FilePath "source\WMo08N4W\T_11.bat" 
})

#Show GUI
$Form.ShowDialog() | out-null
