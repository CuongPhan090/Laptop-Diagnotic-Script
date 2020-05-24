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