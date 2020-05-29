#screen bound
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$res=[System.Windows.Forms.Screen]::PrimaryScreen.Bounds
$res.Width = ($res.Width * 2 ) / 6
$res.Height = ($res.Height * 2 ) / 6

#windows upgrade 

$main_form = New-Object System.Windows.Forms.Form
$main_form.Text ='Updating Windows'
$main_form.StartPosition = "CenterScreen"
$main_form.Size = New-Object System.Drawing.Size($res.Width, $res.Height)

# upgrade Button
$upgrade_button = New-Object System.Windows.Forms.Button
$upgrade_button_width = $res.Width / 2
$upgrade_button.Location = New-Object System.Drawing.Point(0, 0)
$upgrade_button.Size = New-Object System.Drawing.Size($upgrade_button_width, $res.Height )
$upgrade_button.Text = "UPGRADE"
$upgrade_button.Font = 'Microsoft Sans Serif, 15'
$main_form.Controls.Add($upgrade_button)

# activate Button
$activate_button = New-Object System.Windows.Forms.Button
$activate_button_width = $res.Width / 2
$activate_button_x_axis = $res.Width / 2
$activate_button.Location = New-Object System.Drawing.Point($activate_button_x_axis, 0)
$activate_button.Size = New-Object System.Drawing.Size($activate_button_width, $res.Height )
$activate_button.Text = "ACTIVATE"
$activate_button.Font = 'Microsoft Sans Serif, 15'
$main_form.Controls.Add($activate_button)


$upgrade_button.Add_Click({	
	netsh wlan delete profile name="DNCLtechzone"
	changepk.exe /ProductKey YNFDP-JK3VP-V9MHG-864GP-HH66T
})


$activate_button.Add_Click({
	netsh wlan add profile filename="D:\Wi-Fi-DNCLtechzone.xml" Interface="Wi-Fi" user=current
	netsh wlan add profile filename="E:\Wi-Fi-DNCLtechzone.xml" Interface="Wi-Fi" user=current
#connect WiFi
	netsh wlan connect ssid= "DNCLtechzone" name= "DNCLtechzone" 
	start ms-settings:
	slui.exe
	 $navOpenInBackgroundTab = 0x1000;
	$ie = new-object -com InternetExplorer.Application
	$ie.Navigate2("https://www.youtube.com/watch?v=2ZrWHtvSog4");
	$ie.Visible = $true;
})


[Void]$main_form.ShowDialog()

