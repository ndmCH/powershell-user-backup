#####################################
#Author: Nils Marti
#Revision: Rosario Carco
#
#Changelog (newest first)
#(01.03.19)Version 1.2: changed the script so instead of selecting folders of the userprofile that is running the shell (eg. admin) it opens up an explorer to select a folder to use
#(28.02.19)Version 1.1: changed error messages to german, cleaned up some messy code
#(28.02.19)Version 1.0: fully debugged code and extensively tested and script is now fully functional and ready for release
#(27.02.19)Alpha-Version 1.0: started development based on a script found on StackOverflow.com
######################################

#Script starts here

# Import Assemblies
[void][reflection.assembly]::Load("System.Windows.Forms, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")


#region | Global Variables
$date = Get-Date -Format d.MMMM.yyyy
$backupFolder = "Backup_" + "$date" 

#endregion

#Check for USB DIsk , I got that $USBDISK code from "http://stackoverflow.com/questions/10634396/how-do-i-get-the-drive-letter-of-a-usb-drive-in-powershell"
# Search for USB

$UsbDisk = Get-WmiObject win32_diskdrive | Where-Object{$_.interfacetype -eq "USB"} | 
ForEach-Object{Get-WmiObject -Query "ASSOCIATORS OF {Win32_DiskDrive.DeviceID=`"$($_.DeviceID.replace('\','\\'))`"} WHERE AssocClass = Win32_DiskDriveToDiskPartition"} | 
 ForEach-Object{Get-WmiObject -Query "ASSOCIATORS OF {Win32_DiskPartition.DeviceID=`"$($_.DeviceID)`"} WHERE AssocClass = Win32_LogicalDiskToPartition"} | 
 ForEach-Object{$_.deviceid}

# if we dont find the USB Drive, Throw errormessage

 if ( $null -eq $UsbDisk ) { #IF_1
	[void][System.Windows.Forms.MessageBox]::Show("Keine USB Disk gefunden! Bitte USB Disk anschliessen!","USB Drive nicht gefunden")
	
	} #IF_1

else { #Else_1
	
	$Usbfreedisk =  ((Get-WmiObject win32_logicalDisk  | Where-Object { $_.DeviceID -eq $usbdisk }).FreeSpace /1gb -as [INT]) #i recommend using / 1GB in integer for the usb disk for user readability
	
	# opens explorer to select a userprofile and then calculates the size of the specified folders
	Add-Type -AssemblyName System.Windows.Forms
    $browser = New-Object System.Windows.Forms.FolderBrowserDialog
    $null = $browser.ShowDialog()
    $path = $browser.SelectedPath
    
	$Desktop	 = 	Get-ChildItem -Recurse "$path\desktop" | Measure-Object -property length -sum
	$Documents 	 = 	Get-ChildItem -Recurse "$path\documents" | Measure-Object -property length -sum
	$Downloads 	 = 	Get-ChildItem -Recurse "$path\Downloads" | Measure-Object -property length -sum
	$Favorites	 = 	Get-ChildItem -Recurse "$path\Favorites" | Measure-Object -property length -sum
    	$appdata	 = 	Get-ChildItem -Recurse "$path\AppData" | Measure-Object -property length -sum
	
	$DataFilesize = ( ($Desktop.sum + $Documents.SUM + $Pictures.Sum + $Downloads.sum + $Favorites.sum + $Music.sum )   / 1MB ) #here you can use / 1GB/MB/KB to calculate to which unit you prefer
	
	# All Folder Size Calculation Done
	
	if ( $Usbfreedisk -lt $DataFilesize ) {  #IF_2
		
		[void][System.Windows.Forms.MessageBox]::Show("Ihre $usbDisk Festplatte hat nicht genügend Speicher, $DatafileSize GB werden benötigt.","Nicht genügend Speicherplatz")
		
		} #IF_2 
	
	else { #Else_2  
		
		$testfolder = Test-Path "$UsbDisk\$backupFolder" 
		
		if ( $testfolder -eq $true ) { #IF_3 
		
		[void][System.Windows.Forms.MessageBox]::Show("Backup Ordner $UsbDisk\$backupFolder existiert bereits! Bitte umbenennen oder löschen ","Backup Ordner existiert bereits!")
			
		} #IF_3 
		
		else { #Else_3
				
		mkdir "$UsbDisk\$backupFolder"
		
		# Start Copying Data
		
		Robocopy "$path\desktop"  "$UsbDisk\$backupFolder\Desktop" /mir /r:2 /w:3
		Robocopy "$path\documents"  "$UsbDisk\$backupFolder\documents" /mir /r:2 /w:3
		Robocopy "$path\Downloadss"  "$UsbDisk\$backupFolder\Downloads" /mir /r:2 /w:3
		Robocopy "$path\Favorite"  "$UsbDisk\$backupFolder\Favorite" /mir /r:2 /w:3
		Robocopy "$path\AppData"  "$UsbDisk\$backupFolder\AppData" /mir /r:2 /w:3
		
		Write-Host -ForegroundColor 'Green' "Backup Erfolgreich!"
		
		#Open backup folder 
		explorer "$UsbDisk\$backupFolder" 
			
			}#Else_3
	
		}#Else_2

	
	
	
	
}#Else_1
