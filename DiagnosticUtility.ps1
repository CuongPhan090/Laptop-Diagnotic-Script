# OWNER: CUONG PHAN
# DO NOT SHARE PUBLICLY WITHOUT OWNER'S PERMISSION


#########################################################
############# Connect Wifi and Testing ##################
#########################################################
#export Wi-Fi xml profile
#netsh wlan export profile "NETGEAR61" key=clear folder=D:\

#import the network profile
#netsh wlan add profile filename="Wi-Fi-NETGEAR61" Interface="Wi-Fi" user=current
 
#connect WiFi
#netsh wlan connect ssid= "NETGEAR61" name= "NETGEAR61" 

# ---------------------------------------------------------------------------------
#This is where all the test will take place

#open camera
#start microsoft.windows.camera:

#open testing websites
#$navOpenInBackgroundTab = 0x1000;
#$ie = new-object -com InternetExplorer.Application
#$ie.Navigate2("https://www.youtube.com/watch?v=2ZrWHtvSog4");
#$ie.Navigate2("http://deadpixelbuddy.com", $navOpenInBackgroundTab);
#$ie.Navigate2("http://en.key-test.ru", $navOpenInBackgroundTab);
#$ie.Navigate2("https://www.onlinemictest.com/webcam-test/", $navOpenInBackgroundTab);
#$ie.Visible = $true;



#########################################################
################ Retrieve Info Here #####################
#########################################################

# set variables
$h = Get-WmiObject -Class Win32_ComputerSystem
$display = Get-WmiObject win32_videocontroller

# Manufacturer
	$Manufacturer = $h.Manufacturer
	$Manufacturer = ($Manufacturer -split ' ')[0]
	Write-Output "Brand: $Manufacturer"
# Computer Model
	$Model = $h.Model
	Write-Output "Model: $Model"
# RAM Info
	$TotalRAM = ([Math]::Round(($h.TotalPhysicalMemory/1GB),0) )
     	Write-Output "Ram: $TotalRAM GB"

# CPU 
    	$CPU = (Get-WmiObject Win32_processor).Name
	$CPU = $CPU.Substring($CPU.LastIndexOf('-') - 2)
	Write-Output "CPU: $CPU"
	
# Touch / Non-Touch
	
# Screen Resolution
# need to be this format numberxnumber
	$Width = (Get-WmiObject win32_videocontroller).CurrentHorizontalResolution
	$Height = (Get-WmiObject win32_videocontroller).CurrentVerticalResolution
	$Resolution = "$Width x $Height"
	Write-Output "Screen Resolution: $Resolution"
# Storage 
#if storage is greater than 1024GB then convert to 1TB
#	$Storage = (Get-PhysicalDisk).LogicalSectorSize 
#	$MediaType = (Get-PhysicalDisk).MediaType
#	Write-Output "$Storage GB $MediaType"

#$drive_type = @(Get-CimInstance Win32_LogicalDisk).drivetype
#Foreach($i in $drive_type)
#{
#  if ($i -eq 2)
#  {
#    Write-Output $i
#  }
#}

# Graphic Card
	$GPU = $display.caption
	Write-Output "GPU: $GPU"
# FPR ?

# O.S Version
	$OS = Get-CimInstance Win32_OperatingSystem | Select-Object  Caption | ForEach{ $_.Caption }
	$OS = $OS.Substring($OS.LastIndexOf('W'))
	Write-Output "OS version: $OS"

#########################################################
######################## GUI ############################
#########################################################



Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$label_length = 315
$label_width = 50
$label_x_axis = 20
$label_y_axis = 30

$text_box_length = 1500
$text_box_width = 100
$text_box_x_axis = 400
$text_box_y_axis = 30

# Create an empty box
$main_form = New-Object System.Windows.Forms.Form
$main_form.Text ='Laptop Diagnostic Utility'
$main_form.Width = 2000
$main_form.Height = 1300
$main_form.StartPosition = "CenterScreen"
$main_form.AutoSize = $true

# Creating Label and Text Box
# Source
$Label = New-Object System.Windows.Forms.Label
$Label.Text = "SOURCE: "
$Label.Location  = New-Object System.Drawing.Point($label_x_axis,$label_y_axis)
$label.Size = New-Object System.Drawing.Point($label_length, $label_width)
$Label.Font = 'Microsoft Sans Serif,10'
$main_form.Controls.Add($Label)

$Input = New-Object System.Windows.Forms.TextBox
$Input.Location = New-Object System.Drawing.Point($text_box_x_axis, $text_box_y_axis)
$Input.Size = New-Object System.Drawing.Size($text_box_length, $text_box_width)
$Input.Font = 'Microsoft Sans Serif,29'
$Input.Text = "MASTERTRONICS"
$main_form.Controls.Add($Input)

# Invoice #
$Label1 = New-Object System.Windows.Forms.Label
$Label1.Text = "INVOICE #: "
$label_y_axis += 100
$text_box_y_axis += 100
$Label1.Location  = New-Object System.Drawing.Point($label_x_axis, $label_y_axis)
$label1.Size = New-Object System.Drawing.Point($label_length, $label_width)
$Label1.Font = 'Microsoft Sans Serif,10'
$main_form.Controls.Add($Label1)

$Input1 = New-Object System.Windows.Forms.TextBox
$Input1.Location = New-Object System.Drawing.Point($text_box_x_axis, $text_box_y_axis)
$Input1.Size = New-Object System.Drawing.Size($text_box_length, $text_box_width)
$Input1.Font = 'Microsoft Sans Serif,29'
$Input1.Text = "INV00-00000"
$main_form.Controls.Add($Input1)

# Charger
$Label2 = New-Object System.Windows.Forms.Label
$Label2.Text = "CHARGER: "
$label_y_axis += 50
$text_box_y_axis += 50
$Label2.Location  = New-Object System.Drawing.Point($label_x_axis, $label_y_axis)
$label2.Size = New-Object System.Drawing.Point($label_length, $label_width)
$Label2.Font = 'Microsoft Sans Serif,10'
$main_form.Controls.Add($Label2)

$Input2 = New-Object System.Windows.Forms.TextBox
$Input2.Location = New-Object System.Drawing.Point($text_box_x_axis, $text_box_y_axis)
$Input2.Size = New-Object System.Drawing.Size($text_box_length, $text_box_width)
$Input2.Font = 'Microsoft Sans Serif,29'
$Input2.Text = "YES"
$main_form.Controls.Add($Input2)

# ISSUES
$Label3 = New-Object System.Windows.Forms.Label
$Label3.Text = "ISSUES: "
$label_y_axis += 50
$text_box_y_axis += 50
$Label3.Location  = New-Object System.Drawing.Point($label_x_axis, $label_y_axis)
$label3.Size = New-Object System.Drawing.Point($label_length, $label_width)
$Label3.Font = 'Microsoft Sans Serif,10'
$main_form.Controls.Add($Label3)

$Input3 = New-Object System.Windows.Forms.TextBox
$Input3.Location = New-Object System.Drawing.Point($text_box_x_axis, $text_box_y_axis)
$Input3.Size = New-Object System.Drawing.Size($text_box_length, $text_box_width)
$Input3.Font = 'Microsoft Sans Serif,29'
$Input3.Text = "NO"
$main_form.Controls.Add($Input3)

# CONDITION
$Label4 = New-Object System.Windows.Forms.Label
$Label4.Text = "CONDITION:"
$label_y_axis += 100
$text_box_y_axis += 100
$Label4.Location  = New-Object System.Drawing.Point($label_x_axis, $label_y_axis)
$label4.Size = New-Object System.Drawing.Point($label_length, $label_width)
$Label4.Font = 'Microsoft Sans Serif,10'
$main_form.Controls.Add($Label4)

$Input4 = New-Object System.Windows.Forms.TextBox
$Input4.Location = New-Object System.Drawing.Point($text_box_x_axis, $text_box_y_axis)
$Input4.Size = New-Object System.Drawing.Size($text_box_length, $text_box_width)
$Input4.Font = 'Microsoft Sans Serif,29'
$Input4.Text = "9/10 - LIKE NEW"
$main_form.Controls.Add($Input4)

# S/N 

$Label5 = New-Object System.Windows.Forms.Label
$Label5.Text = "S/N"
$label_y_axis += 50
$text_box_y_axis += 50
$Label5.Location  = New-Object System.Drawing.Point($label_x_axis, $label_y_axis)
$label5.Size = New-Object System.Drawing.Point($label_length, $label_width)
$Label5.Font = 'Microsoft Sans Serif,10'
$main_form.Controls.Add($Label5)

$SN = (get-ciminstance win32_bios).serialnumber
$Input5 = New-Object System.Windows.Forms.TextBox
$Input5.Location = New-Object System.Drawing.Point($text_box_x_axis, $text_box_y_axis)
$Input5.Size = New-Object System.Drawing.Size($text_box_length, $text_box_width)
$Input5.Font = 'Microsoft Sans Serif,29'
$Input5.Text = "$SN"
$main_form.Controls.Add($Input5)

# COLOR
$Label6 = New-Object System.Windows.Forms.Label
$Label6.Text = "COLOR: "
$label_y_axis += 50
$text_box_y_axis += 50
$Label6.Location  = New-Object System.Drawing.Point($label_x_axis, $label_y_axis)
$label6.Size = New-Object System.Drawing.Point($label_length, $label_width)
$Label6.Font = 'Microsoft Sans Serif,10'
$main_form.Controls.Add($Label6)

$Input6 = New-Object System.Windows.Forms.TextBox
$Input6.Location = New-Object System.Drawing.Point($text_box_x_axis, $text_box_y_axis)
$Input6.Size = New-Object System.Drawing.Size($text_box_length, $text_box_width)
$Input6.Font = 'Microsoft Sans Serif,29'
$Input6.Text = "N\A"
$main_form.Controls.Add($Input6)

# SPECS
$Label7 = New-Object System.Windows.Forms.Label
$Label7.Text = "SPECS: "
$label_y_axis += 100
$text_box_y_axis += 100
$Label7.Location  = New-Object System.Drawing.Point($label_x_axis, $label_y_axis)
$label7.Size = New-Object System.Drawing.Point($label_length, $label_width)
$Label7.Font = 'Microsoft Sans Serif,10'
$main_form.Controls.Add($Label7)

$Input7 = New-Object System.Windows.Forms.TextBox
$Input7.Location = New-Object System.Drawing.Point($text_box_x_axis, $text_box_y_axis)
$Input7.Size = New-Object System.Drawing.Size($text_box_length, $text_box_width)
$Input7.Font = 'Microsoft Sans Serif,29'
$Input7.Multiline = $True
$Specs = "$Manufacturer $Model Inch $Resolution Touch $CPU $Ram $GPU $OS"
$Specs = $Specs.ToUpper()
$Input7.Text = "$Specs"
$main_form.Controls.Add($Input7)

# MODEL #
$Label8 = New-Object System.Windows.Forms.Label
$Label8.Text = "MODEL: "
$label_y_axis += 100
$text_box_y_axis += 100
$Label8.Location  = New-Object System.Drawing.Point($label_x_axis, $label_y_axis)
$label8.Size = New-Object System.Drawing.Point($label_length, $label_width)
$Label8.Font = 'Microsoft Sans Serif,10'
$main_form.Controls.Add($Label8)

$Input8 = New-Object System.Windows.Forms.TextBox
$Input8.Location = New-Object System.Drawing.Point($text_box_x_axis, $text_box_y_axis)
$Input8.Size = New-Object System.Drawing.Size($text_box_length, $text_box_width)
$Input8.Font = 'Microsoft Sans Serif,29'
$Input8.Text = "N/A"
$main_form.Controls.Add($Input8)

# UNPLUG TEST
$Label9 = New-Object System.Windows.Forms.Label
$Label9.Text = "UNPLUG TEST: "
$label_y_axis += 100
$text_box_y_axis += 100
$Label9.Location  = New-Object System.Drawing.Point($label_x_axis, $label_y_axis)
$label9.Size = New-Object System.Drawing.Point($label_length, $label_width)
$Label9.Font = 'Microsoft Sans Serif,10'
$main_form.Controls.Add($Label9)

$Input9 = New-Object System.Windows.Forms.TextBox
$Input9.Location = New-Object System.Drawing.Point($text_box_x_axis, $text_box_y_axis)
$Input9.Size = New-Object System.Drawing.Size($text_box_length, $text_box_width)
$Input9.Font = 'Microsoft Sans Serif,29'
$Input9.Text = "PASS"
$main_form.Controls.Add($Input9)

# ACTION PLAN
$Label10 = New-Object System.Windows.Forms.Label
$Label10.Text = "ACTION PLAN: "
$label_y_axis += 100
$text_box_y_axis += 100
$Label10.Location  = New-Object System.Drawing.Point($label_x_axis, $label_y_axis)
$label10.Size = New-Object System.Drawing.Point($label_length, $label_width)
$Label10.Font = 'Microsoft Sans Serif,10'
$main_form.Controls.Add($Label10)

$Input10 = New-Object System.Windows.Forms.TextBox
$Input10.Location = New-Object System.Drawing.Point($text_box_x_axis, $text_box_y_axis)
$Input10.Size = New-Object System.Drawing.Size($text_box_length, $text_box_width)
$Input10.Font = 'Microsoft Sans Serif,29'
$Input10.Text = "RESALE"
$main_form.Controls.Add($Input10)

# Send Button
$Button = New-Object System.Windows.Forms.Button
$button_x_axis = ($main_form.Width / 2) / 3
$button_y_axis = $text_box_y_axis + 150
$label_length *= 4
$label_width *= 3
$Button.Location = New-Object System.Drawing.Size($button_x_axis, $button_y_axis)
$Button.Size = New-Object System.Drawing.Size($label_length, $label_width)
$Button.Text = "SEND EMAIL TO DNCL.INC"
$Button.Font = 'Microsoft Sans Serif,15' 
$main_form.Controls.Add($Button)

# Send Event Listener
$Button.Add_Click( { $Input10.Text = "It works"} )



[void] $main_form.ShowDialog()
