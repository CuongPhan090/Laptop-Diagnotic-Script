#########################################################
################ Retrieve Info Here #####################
#########################################################

# set variables
$h = Get-WmiObject -Class Win32_ComputerSystem
$display = Get-WmiObject win32_videocontroller

# Manufacturer
	$Manufacturer = $h.Manufacturer
	$Manufacturer = ($Manufacturer -split ' ')[0]
	#Write-Output "Brand: $Manufacturer"
# Computer Model
	$Model = $h.Model
	#Write-Output "Model: $Model"
# RAM Info
	$Ram = ([Math]::Round(($h.TotalPhysicalMemory/1GB),0) )
	$Ram = "$Ram" + "GB" 
	#Write-Output "Ram: $Ram"

# CPU 
    	$CPU = (Get-WmiObject Win32_processor).Name
	$CPU = $CPU.Substring($CPU.LastIndexOf('-') - 2)
	#Write-Output "CPU: $CPU"
	
# Touch / Non-Touch
	
# Screen Resolution
# need to be this format numberxnumber
	$Width = (Get-WmiObject win32_videocontroller).CurrentHorizontalResolution
	$Height = (Get-WmiObject win32_videocontroller).CurrentVerticalResolution
	$Resolution = "$Width x $Height"
	#Write-Output "Screen Resolution: $Resolution"
# Storage 
	$storage_size = (Get-PhysicalDisk).size
	$storage_mediatype = (Get-PhysicalDisk).MediaType
    	$storage = ""
	$i = 0
	
	while ($i -le $storage_size.length - 1)
	{
		$storage_size[$i] = [math]::Round($storage_size[$i] / 1000000000)
		if ($storage_size[$i] -ne "16")
		{
		   if ($storage_mediatype[$i] -eq "SSD")
		   {
		      $storage = $storage + $storage_size[$i] + "GB " + $storage_mediatype[$i] + " "
		   }

		   if ($storage_mediatype[$i] -eq "HDD")
		   {
		      $storage = $storage + $storage_size[$i] + "GB " + $storage_mediatype[$i] + " "
		   }
		}
		$i++
	}
	#Write-Output "Storage: $Storage"


# Graphic Card
	$GPU = $display.caption
	#Write-Output "GPU: $GPU"
# FPR ?

# O.S Version
	$OS = Get-CimInstance Win32_OperatingSystem | Select-Object  Caption | ForEach{ $_.Caption }
	$OS = $OS.Substring($OS.LastIndexOf('W'))
	#Write-Output "OS version: $OS"

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
$text_box_width = 50
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
$Label1 = New-Object System.Windows.Forms.Label
$Label1.Text = "SOURCE: "
$Label1.Location  = New-Object System.Drawing.Point($label_x_axis,$label_y_axis)
$label1.Size = New-Object System.Drawing.Point($label_length, $label_width)
$Label1.Font = 'Microsoft Sans Serif,10'
$main_form.Controls.Add($Label1)

$Input1 = New-Object System.Windows.Forms.RichTextBox
$Input1.Location = New-Object System.Drawing.Point($text_box_x_axis, $text_box_y_axis)
$Input1.Size = New-Object System.Drawing.Size($text_box_length, $text_box_width)
$Input1.Font = 'Microsoft Sans Serif,29'
$Input1.Text = "MASTERTRONICS"
$Input1.ForeColor = "RED"
$main_form.Controls.Add($Input1)

# Invoice #
$Label2 = New-Object System.Windows.Forms.Label
$Label2.Text = "INVOICE #: "
$label_y_axis += 100
$text_box_y_axis += 100
$Label2.Location  = New-Object System.Drawing.Point($label_x_axis, $label_y_axis)
$label2.Size = New-Object System.Drawing.Point($label_length, $label_width)
$Label2.Font = 'Microsoft Sans Serif,10'
$main_form.Controls.Add($Label2)

$Input2 = New-Object System.Windows.Forms.RichTextBox
$Input2.Location = New-Object System.Drawing.Point($text_box_x_axis, $text_box_y_axis)
$Input2.Size = New-Object System.Drawing.Size($text_box_length, $text_box_width)
$Input2.Font = 'Microsoft Sans Serif,29'
$Input2.Text = "INV00-00000"
$Input2.ForeColor = "RED"
$main_form.Controls.Add($Input2)

# Charger
$Label3 = New-Object System.Windows.Forms.Label
$Label3.Text = "CHARGER: "
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
$Input3.Text = "YES"
$main_form.Controls.Add($Input3)

# ISSUES
$Label4 = New-Object System.Windows.Forms.Label
$Label4.Text = "ISSUES: "
$label_y_axis += 50
$text_box_y_axis += 50
$Label4.Location  = New-Object System.Drawing.Point($label_x_axis, $label_y_axis)
$label4.Size = New-Object System.Drawing.Point($label_length, $label_width)
$Label4.Font = 'Microsoft Sans Serif,10'
$main_form.Controls.Add($Label4)

$Input4 = New-Object System.Windows.Forms.TextBox
$Input4.Location = New-Object System.Drawing.Point($text_box_x_axis, $text_box_y_axis)
$Input4.Size = New-Object System.Drawing.Size($text_box_length, $text_box_width)
$Input4.Font = 'Microsoft Sans Serif,29'
$Input4.Text = "NO"
$main_form.Controls.Add($Input4)

# CONDITION
$Label5 = New-Object System.Windows.Forms.Label
$Label5.Text = "CONDITION:"
$label_y_axis += 100
$text_box_y_axis += 100
$Label5.Location  = New-Object System.Drawing.Point($label_x_axis, $label_y_axis)
$label5.Size = New-Object System.Drawing.Point($label_length, $label_width)
$Label5.Font = 'Microsoft Sans Serif,10'
$main_form.Controls.Add($Label5)

$Input5 = New-Object System.Windows.Forms.RichTextBox
$Input5.Location = New-Object System.Drawing.Point($text_box_x_axis, $text_box_y_axis)
$Input5.Size = New-Object System.Drawing.Size($text_box_length, $text_box_width)
$Input5.Font = 'Microsoft Sans Serif,29'
$Input5.Text = "9/10 - LIKE NEW"
$Input5.ForeColor = "RED"
$main_form.Controls.Add($Input5)

# S/N 

$Label6 = New-Object System.Windows.Forms.Label
$Label6.Text = "S/N"
$label_y_axis += 50
$text_box_y_axis += 50
$Label6.Location  = New-Object System.Drawing.Point($label_x_axis, $label_y_axis)
$label6.Size = New-Object System.Drawing.Point($label_length, $label_width)
$Label6.Font = 'Microsoft Sans Serif,10'
$main_form.Controls.Add($Label6)

$SN = (get-ciminstance win32_bios).serialnumber
$Input6 = New-Object System.Windows.Forms.TextBox
$Input6.Location = New-Object System.Drawing.Point($text_box_x_axis, $text_box_y_axis)
$Input6.Size = New-Object System.Drawing.Size($text_box_length, $text_box_width)
$Input6.Font = 'Microsoft Sans Serif,29'
$Input6.Text = "$SN"
$main_form.Controls.Add($Input6)

# COLOR
$Label7 = New-Object System.Windows.Forms.Label
$Label7.Text = "COLOR: "
$label_y_axis += 50
$text_box_y_axis += 50
$Label7.Location  = New-Object System.Drawing.Point($label_x_axis, $label_y_axis)
$label7.Size = New-Object System.Drawing.Point($label_length, $label_width)
$Label7.Font = 'Microsoft Sans Serif,10'
$main_form.Controls.Add($Label7)

$Input7 = New-Object System.Windows.Forms.RichTextBox
$Input7.Location = New-Object System.Drawing.Point($text_box_x_axis, $text_box_y_axis)
$Input7.Size = New-Object System.Drawing.Size($text_box_length, $text_box_width)
$Input7.Font = 'Microsoft Sans Serif,29'
$Input7.Text = "N/A"
$Input7.ForeColor = "RED"
$main_form.Controls.Add($Input7)

# SPECS
$Label8 = New-Object System.Windows.Forms.Label
$Label8.Text = "SPECS: "
$label_y_axis += 100
$text_box_y_axis += 100
$Label8.Location  = New-Object System.Drawing.Point($label_x_axis, $label_y_axis)
$Label8.Size = New-Object System.Drawing.Point($label_length, $label_width)
$Label8.Font = 'Microsoft Sans Serif,10'
$main_form.Controls.Add($Label8)

$Input8 = New-Object System.Windows.Forms.RichTextBox
$text_box_width *= 2
$Input8.Location = New-Object System.Drawing.Point($text_box_x_axis, $text_box_y_axis)
$Input8.Size = New-Object System.Drawing.Size($text_box_length, $text_box_width)
$Input8.Font = 'Microsoft Sans Serif,29'
$Input8.Multiline = $True
$Specs = "$Manufacturer $Model ?Inch $Resolution ?-Touch $CPU $Ram $Storage $GPU ?FPR $OS"
$Specs = $Specs.ToUpper()	
$i = 0
while ($i -lt ($Specs.split()).length)
{
	#$Input8.AppendText("$Specs.split()[$i] ")
	#Write-Output $Specs.split()[$i]
	if ($Specs.split()[$i][0] -ne "?")
	{
		$Input8.SelectionColor = 'black'
		$Input8.AppendText($Specs.split()[$i] + " " )	
	}
	else	
	{
		$Input8.SelectionColor = 'red'
		$Input8.AppendText($Specs.split()[$i] +  " " )
	}
	$i++
}
#$Input8.Text = "$Specs"
$main_form.Controls.Add($Input8)
$text_box_width /= 2

# MODEL #
$Input9 = New-Object System.Windows.Forms.Label
$Input9.Text = "MODEL: "
$label_y_axis += 100
$text_box_y_axis += 100
$Input9.Location  = New-Object System.Drawing.Point($label_x_axis, $label_y_axis)
$Input9.Size = New-Object System.Drawing.Point($label_length, $label_width)
$Input9.Font = 'Microsoft Sans Serif,10'
$main_form.Controls.Add($Input9)

$Input9 = New-Object System.Windows.Forms.RichTextBox
$Input9.Location = New-Object System.Drawing.Point($text_box_x_axis, $text_box_y_axis)
$Input9.Size = New-Object System.Drawing.Size($text_box_length, $text_box_width)
$Input9.Font = 'Microsoft Sans Serif,29'
$Input9.ForeColor = "RED"
$Input9.Text = "N/A"
$main_form.Controls.Add($Input9)

# UNPLUG TEST
$Input10 = New-Object System.Windows.Forms.Label
$Input10.Text = "UNPLUG TEST: "
$label_y_axis += 100
$text_box_y_axis += 100
$Input10.Location  = New-Object System.Drawing.Point($label_x_axis, $label_y_axis)
$Input10.Size = New-Object System.Drawing.Point($label_length, $label_width)
$Input10.Font = 'Microsoft Sans Serif,10'
$main_form.Controls.Add($Input10)

$Input10 = New-Object System.Windows.Forms.TextBox
$Input10.Location = New-Object System.Drawing.Point($text_box_x_axis, $text_box_y_axis)
$Input10.Size = New-Object System.Drawing.Size($text_box_length, $text_box_width)
$Input10.Font = 'Microsoft Sans Serif,29'
$Input10.Text = "PASS"
$main_form.Controls.Add($Input10)

# ACTION PLAN
$Label11 = New-Object System.Windows.Forms.Label
$Label11.Text = "ACTION PLAN: "
$label_y_axis += 100
$text_box_y_axis += 100
$Label11.Location  = New-Object System.Drawing.Point($label_x_axis, $label_y_axis)
$Label11.Size = New-Object System.Drawing.Point($label_length, $label_width)
$Label11.Font = 'Microsoft Sans Serif,10'
$main_form.Controls.Add($Label11)

$Input11 = New-Object System.Windows.Forms.TextBox
$Input11.Location = New-Object System.Drawing.Point($text_box_x_axis, $text_box_y_axis)
$Input11.Size = New-Object System.Drawing.Size($text_box_length, $text_box_width)
$Input11.Font = 'Microsoft Sans Serif,29'
$Input11.Text = "RESALE"
$main_form.Controls.Add($Input11)

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




#########################################################
######################## EMAIL ##########################
#########################################################

function email_content($SOURCE, $INVOICE, $CHARGER, $ISSUES, $CONDITION, $SN, $COLOR, $SPECS, $MODEL, $UNPLUG_TEST, $ACTION_PLAN)
{

	return "SOURCE: " + $SOURCE + 
	"`n`nINVOICE #: " + $INVOICE +
	"`nCHARGER: " + $CHARGER +
	"`nISSUES: " + $ISSUES +
	"`n`nCONDITION: " + $CONDITION +
	"`nS/N: " + $SN +
	"`nCOLOR: " + $COLOR +
	#DELL XPS 9370 13.3INCH 4K TOUCH I7-8550U 16 1TB SSD INTEGRATED  WIN 10 PRO (TECH UPGRADED)
	"`n`n" + $SPECS +
	"`nMODEL #: " + $MODEL +
	"`n`nUNPLUG TEST: " + $UNPLUG_TEST +
	"`n`n" + $ACTION_PLAN
	#Write-Output $Body
}

function send_email
{
	$From = "vcuong0909@gmail.com"	
	$To = "daniel.dncltechzone@gmail.com"
	#$Cc = ""
	$Subject = ("ARRIVED FROM " + $Input1.Text + ", INVOICE #: " + $Input2.Text + ", S/N: " + $Input6.Text).ToUpper()
	$SMTPServer = "smtp.gmail.com"
	$SMTPPort = "587"
	# Create Credential Object
	$username = "vcuong0909@gmail.com"
	$secureStringPwd = ConvertTo-SecureString -String "jqjofticzavmband" -AsPlainText -Force
	$creds = New-Object System.Management.Automation.PSCredential -ArgumentList $username, $secureStringPwd
	Send-MailMessage -From $From -to $To -Subject $Subject -Body $a -SmtpServer $SMTPServer -Port $SMTPPort -UseSsl -Credential $creds

 
}


#########################################################
######################## EXCEPTION ######################
#########################################################	

# Send Event Listener
$Button.Add_Click({
	$a = email_content $Input1.Text $Input2.Text $Input3.Text $Input4.Text $Input5.Text $Input6.Text $Input7.Text $Input8.Text $Input9.Text $Input10.Text $Input11.Text	
	send_email
	$wshell = New-Object -ComObject Wscript.Shell 
	$notification = "Email has sent successfully. `nSystem is shutting down in 10 seconds."
	$main_form.Close()
	$Output = $wshell.Popup($notification)
	#shutdown /s
})
[Void]$main_form.ShowDialog()
























