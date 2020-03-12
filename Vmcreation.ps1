
Param(
    [Parameter(Mandatory=$false)]
    [ValidateNotNullOrEmpty()]
    #===========================================================
    # VM Name 
    #===========================================================
    $NewVMName = "VM-User-1",
    #===========================================================
    # Type of server Continuous Integration (CI)
    # "CI-VS17" = Windows Server 2016 and Visual Studio 2017
    # "CI-VS19" = Windows Server 2016 and Visual Studio 2019
    #===========================================================
    # Continuous Delivery (CD)
    #===========================================================
    # "CD-WIN16-SQL16" = Windows Server 2016 and SQL Server 2016
    # "CD-WIN16-SQL19" = Windows Server 2016 and SQL Server 2019
    #===========================================================
    $VMType = "CD-WIN16-SQL16",
    #$ADOUniversalPackageVersion = "0.1.0",
    #===========================================================
    # Name of virtual switch
    #===========================================================
    $VMExternalSwitch = "VM-ExternalSwitch",
    #===========================================================
    # Set VM resources(RAM, DISK, etc)
    #===========================================================
    $VMProcessorCount = 2,
    [Int64]$VMMemoryMinimumBytes = 8GB,
    [Int64]$VMMemoryStartupBytes = 8GB,
    [Int64]$VMMemoryMaximumBytes = 8GB,
    #===========================================================
    # Service accounts 
    #===========================================================
    [string]$localAdminPassword = "password",
    [string]$localAdminUsername = "administrator",
    #===========================================================
    #Location to which the VHDs have been placed
    #===========================================================
    $artifactLocation = "V:\LocalVirtualMachine\",
    [String] $DirectoryToCreate = $NewVMName
    )
#=========================================================================================================================================================================================

#===========================================================
#Creation of Partition for VM
#===========================================================
$maxsize = Get-PartitionSupportedSize -DriveLetter C
$doingmath = [math]::round(($maxsize.SizeMax)*.8)
$shrinksize = [math]::Round([float]$doingmath)
Get-Partition -DriveLetter C | Resize-Partition -Size $shrinksize
New-Partition -DiskNumber 0 -UseMaximumSize -Driveletter V |
Format-Volume -FileSystem NTFS -Confirm:$false

#===========================================================
#Checking for directories to store images
#===========================================================

    if (-not (Test-Path -LiteralPath $artifactLocation)) {
    
        try {
            New-Item -Path $artifactLocation -ItemType Directory -ErrorAction Stop | Out-Null #-Force
        }
        catch {
            Write-Error -Message "Unable to create directory '$artifactLocation'. Error was: $_" -ErrorAction Stop
        }
            Write-host "Successfully created directory '$artifactLocation'."

        }
        else {
            Write-host "Directory already existed."
    }

    set-location -Path $artifactLocation

    if (-not (Test-Path -LiteralPath $DirectoryToCreate)) {
    
        try {
            New-Item -Path $DirectoryToCreate -ItemType Directory -ErrorAction Stop | Out-Null #-Force
        }
        catch {
            Write-Error -Message "Unable to create directory '$DirectoryToCreate'. Error was: $_" -ErrorAction Stop
        }
            Write-host "Successfully created directory '$DirectoryToCreate'."

        }
        else {
            Write-host "Directory already existed."
    }

#=========================================================================================================================================================================================
# Downloading Image from Share
#=========================================================================================================================================================================================
Copy-Item -Path "C:\OP-CD-Win2016-SQL16\Virtual Hard Disks" -Destination $($artifactLocation + $DirectoryToCreate) -Recurse
Copy-Item -Path "C:\OP-CD-Win2016-SQL16\Virtual Machines" -Destination $($artifactLocation + $DirectoryToCreate) -Recurse


#===========================================================
#Creating Virutal Machine
#===========================================================
if($NewVMName.Length -lt 15){
Try{
    $path = $($artifactLocation + $DirectoryToCreate)
    if(![System.IO.File]::Exists($path)){

    Rename-Item -Path $($artifactLocation + $DirectoryToCreate + "\Virtual Hard Disks\win2016-WU.vhdx") -NewName $($NewVMName + ".vhdx")
    Rename-Item -Path $($artifactLocation + $DirectoryToCreate + "\Virtual Hard Disks\win2016-WU-0.vhdx") -NewName $($NewVMName + "_data.vhdx")
    Rename-Item -Path $($artifactLocation + $DirectoryToCreate + "\Virtual Hard Disks\win2016-WU-1.vhdx") -NewName $($NewVMName + "_log.vhdx")

    try{
    New-VMSwitch -name $VMExternalSwitch  -NetAdapterName Ethernet -AllowManagementOS $true
    }
    catch {
        Write-Host "Network adapter is already created" -ForegroundColor Yellow
        write-host "Caught an exception:" -ForegroundColor Yellow
        write-host "Exception Type: $($_.Exception.GetType().FullName)" -ForegroundColor Yellow
        write-host "Exception Message: $($_.Exception.Message)" -ForegroundColor Yellow
    }

    New-VM -Name $NewVMName -VHDPath $($artifactLocation + $DirectoryToCreate + "\Virtual Hard Disks\" + $NewVMName + ".vhdx") -Generation 2 -SwitchName $VMExternalSwitch

    Set-VM -Name $NewVMName -ProcessorCount $VMProcessorCount -DynamicMemory  -MemoryMinimumBytes $VMMemoryMinimumBytes  -MemoryStartupBytes $VMMemoryStartupBytes -MemoryMaximumBytes $VMMemoryMaximumBytes

    if($VMType -like '*CD*' ){
        Add-VMHardDiskDrive -VMName $NewVMName -path $($artifactLocation + $DirectoryToCreate  + "\Virtual Hard Disks\" +$NewVMName + "_data.vhdx") -ControllerType SCSI 
        Add-VMHardDiskDrive -VMName $NewVMName -path $($artifactLocation + $DirectoryToCreate  + "\Virtual Hard Disks\" +$NewVMName + "_log.vhdx") -ControllerType SCSI 
        }
        else{
        Add-VMHardDiskDrive -VMName $NewVMName -path $($artifactLocation + $DirectoryToCreate + "\Virtual Hard Disks\" +$NewVMName + "_data.vhdx") -ControllerType SCSI 
        }
    
      Add-VMNetworkAdapter -VMName $NewVMName -SwitchName VM-ExternalSwitch -Name Test_NetAdapter
         }
    } 
    catch{
        Write-Host "The Server name you have selected is already being used. Please select a new name. Exception is listed below" -ForegroundColor Red
        write-host "Caught an exception:" -ForegroundColor Red
        write-host "Exception Type: $($_.Exception.GetType().FullName)" -ForegroundColor Red
        write-host "Exception Message: $($_.Exception.Message)" -ForegroundColor Red
    }
}
else {
    Write-Host "Please enter a server name less then 15 characters!"
}

Start-VM -Name $NewVMName

Start-Sleep -Seconds 30


#===========================================================
#Changing name of virtual machine
#===========================================================
Try{
write-host "Attempting to rename the server."
$password = $localAdminPassword | ConvertTo-SecureString -AsPlainText -Force
$username = $localAdminUsername
$Cred = New-Object System.Management.Automation.PSCredential ($username, $password)
Invoke-Command -VMName $NewVMName -Credential $Cred -ScriptBlock {param($comp,$localAdminUsername,$localAdminPassword)  
$password = $localAdminPassword | ConvertTo-SecureString -AsPlainText -Force
$username = $localAdminUsername
$Cred = New-Object System.Management.Automation.PSCredential ($username, $password)
Rename-Computer -NewName $comp -DomainCredential $Cred -restart -Force} -ArgumentList $NewVMName, $localAdminPassword,$localAdminUsername
}
Catch{
    write-host "Caught an exception:" -ForegroundColor Red
    write-host "Exception Type: $($_.Exception.GetType().FullName)" -ForegroundColor Red
    write-host "Exception Message: $($_.Exception.Message)" -ForegroundColor Red
}

Start-Sleep -Seconds 60


if($VMType -like '*CD*' ){

#===========================================================
#Finishing install of SQL Server
#===========================================================
Try{
write-host "Finishing Install of SQL Server"
    $password = $localAdminPassword | ConvertTo-SecureString -AsPlainText -Force
    $username = $localAdminUsername
    $Cred = New-Object System.Management.Automation.PSCredential ($username, $password)
Invoke-Command -VMName $NewVMName -Credential $Cred -ScriptBlock {param($localAdminUsername, $localAdminPassword) 
    $password = $localAdminPassword | ConvertTo-SecureString -AsPlainText -Force
    $username = $localAdminUsername
    $Cred = New-Object System.Management.Automation.PSCredential ($username, $password)

C:\SQLMedia\setup.exe /IACCEPTSQLSERVERLICENSETERMS /IACCEPTROPENLICENSETERMS /Q /ACTION="CompleteImage" /SUPPRESSPRIVACYSTATEMENTNOTICE="False" /ENU="True" /QUIETSIMPLE="False" /USEMICROSOFTUPDATE="False" /HELP="False" /INDICATEPROGRESS="False" /X86="False" /INSTANCENAME="MSSQLSERVER" /INSTANCEID="MSSQLSERVER" /RSINSTALLMODE="FilesOnlyMode" /SQLTELSVCACCT="NT Service\SQLTELEMETRY" /SQLTELSVCSTARTUPTYPE="Automatic" /AGTSVCACCOUNT="NT Service\SQLSERVERAGENT" /AGTSVCSTARTUPTYPE="Manual" /SQLSVCSTARTUPTYPE="Automatic" /FILESTREAMLEVEL="0" /ENABLERANU="False" /SQLCOLLATION="SQL_Latin1_General_CP1_CI_AS" /SQLSVCACCOUNT="NT Service\MSSQLSERVER" /SQLSVCINSTANTFILEINIT="False" /SQLSYSADMINACCOUNTS="administrator" /SQLTEMPDBFILECOUNT="4" /SQLTEMPDBFILESIZE="8" /SQLTEMPDBFILEGROWTH="64" /SQLTEMPDBLOGFILESIZE="8" /SQLTEMPDBLOGFILEGROWTH="64" /ADDCURRENTUSERASSQLADMIN="False" /TCPENABLED="1" /NPENABLED="0" /BROWSERSVCSTARTUPTYPE="Disabled" /RSSVCACCOUNT="NT Service\ReportServer"
} -ArgumentList $localAdminUsername, $localAdminPassword
} 
Catch{
    write-host "Caught an exception:" -ForegroundColor Red
    write-host "Exception Type: $($_.Exception.GetType().FullName)" -ForegroundColor Red
    write-host "Exception Message: $($_.Exception.Message)" -ForegroundColor Red
}

Start-Sleep -Seconds 15

#===========================================================
#Changing drive letters
#===========================================================
Try{
write-host "Changing drive letters"
$password = $localAdminPassword  | ConvertTo-SecureString -AsPlainText -Force
$username = $localAdminUsername
$Cred = New-Object System.Management.Automation.PSCredential ($username, $password)
Invoke-Command -VMName $NewVMName -Credential $Cred -ScriptBlock{
Get-Partition -DiskNumber 1 | Set-Partition -NewDriveLetter D
Get-Partition -DiskNumber 2 | Set-Partition -NewDriveLetter E
}
}
Catch{
    write-host "Caught an exception:" -ForegroundColor Red
    write-host "Exception Type: $($_.Exception.GetType().FullName)" -ForegroundColor Red
    write-host "Exception Message: $($_.Exception.Message)" -ForegroundColor Red
}

Start-Sleep -Seconds 15

#===========================================================
#Changing SQL Data and Log Directory
#===========================================================
Try{
write-host "Attempting to change SQL Server Data(D:\) and Log(E:\) drive letters"
$password = $localAdminPassword | ConvertTo-SecureString -AsPlainText -Force
$username = $localAdminUsername
$Cred = New-Object System.Management.Automation.PSCredential ($username, $password)
Invoke-Command -VMName $NewVMName -Credential $Cred -ScriptBlock{
Invoke-Sqlcmd -ServerInstance $NewVMName  -Database Master -Query "EXEC xp_instance_regwrite N'HKEY_LOCAL_MACHINE', N'Software\Microsoft\MSSQLServer\MSSQLServer', N'DefaultData', REG_SZ, N'D:'EXEC xp_instance_regwrite N'HKEY_LOCAL_MACHINE', N'Software\Microsoft\MSSQLServer\MSSQLServer', N'DefaultLog', REG_SZ, N'E:'" -Verbose}
}
Catch{
    write-host "Caught an exception:" -ForegroundColor Red
    write-host "Exception Type: $($_.Exception.GetType().FullName)" -ForegroundColor Red
    write-host "Exception Message: $($_.Exception.Message)" -ForegroundColor Red
}
#===========================================================
# Switching SQL Authentication to Mixed Mode
#===========================================================
Try{
write-host "Switching SQL Authentication to Mixed Mode"
$password = $localAdminPassword | ConvertTo-SecureString -AsPlainText -Force
$username = $localAdminUsername
$Cred = New-Object System.Management.Automation.PSCredential ($username, $password)
Invoke-Command -VMName $NewVMName -Credential $Cred -ScriptBlock{
Invoke-Sqlcmd -ServerInstance $NewVMName  -Database Master -Query "USE [master] EXEC xp_instance_regwrite N'HKEY_LOCAL_MACHINE', N'Software\Microsoft\MSSQLServer\MSSQLServer', N'LoginMode', REG_DWORD, 2 " -Verbose}
}
Catch{
    write-host "Caught an exception:" -ForegroundColor Red
    write-host "Exception Type: $($_.Exception.GetType().FullName)" -ForegroundColor Red
    write-host "Exception Message: $($_.Exception.Message)" -ForegroundColor Red
}

#===========================================================
# Modifying SQL Server Settings 
#===========================================================
Try{
write-host "Modifying SQL Server Settings(MAX SQL SERVER MEMORYE"
$password = $localAdminPassword | ConvertTo-SecureString -AsPlainText -Force
$username = $localAdminUsername
$Cred = New-Object System.Management.Automation.PSCredential ($username, $password)

Invoke-Command -VMName $NewVMName -Credential $Cred -ScriptBlock{
# Getting the memory allocated to VM
Function Get-ComputerMemory {
    $mem = Get-WMIObject -class Win32_PhysicalMemory | Measure-Object -Property Capacity -Sum
    return ($mem.Sum / 1MB);
}

Get-ComputerMemory

#===========================================================
#Calculating how much memory can be allocated to SQL Server
#===========================================================

Function Get-SQLMaxMemory { 
    $memtotal = Get-ComputerMemory
    $min_os_mem = 2048 ;
    if ($memtotal -le $min_os_mem) {
        Return $null;
    }
    if ($memtotal -ge 8192) {
        $sql_mem = $memtotal - 2048
    } else {
        $sql_mem = $memtotal * 0.8 ;
    }
    return [int]$sql_mem ;  
}

Get-SQLMaxMemory

#===========================================================
#Setting Max SQL Server memory based on calculated value
#===========================================================
Function Set-SQLInstanceMemory {
    param (
        [string]$SQLInstanceName = ".", 
        [int]$maxMem = $sql_mem, 
        [int]$minMem = 0
    )
 
    if ($minMem -eq 0) {
        $minMem = $maxMem
    }
    [reflection.assembly]::LoadWithPartialName("Microsoft.SqlServer.Smo") | Out-Null
    $srv = New-Object Microsoft.SQLServer.Management.Smo.Server($SQLInstanceName)
    if ($srv.status) {
        Write-Host "[Running] Setting Maximum Memory to: $($srv.Configuration.MaxServerMemory.RunValue)"
        Write-Host "[Running] Setting Minimum Memory to: $($srv.Configuration.MinServerMemory.RunValue)"
 
        Write-Host "[New] Setting Maximum Memory to: $maxmem"
        Write-Host "[New] Setting Minimum Memory to: 0"
        $srv.Configuration.MaxServerMemory.ConfigValue = $maxMem
        $srv.Configuration.MinServerMemory.ConfigValue = 0   
        $srv.Configuration.Alter()

        Restart-Service -Name MSSQLSERVER -Force
        Restart-Service -Name SQLSERVERAGENT -Force
    }
}

#===========================================================
#Executing SQLInstanceMemory function
#===========================================================
$MSSQLInstance = $NewVMName
Set-SQLInstanceMemory $MSSQLInstance (Get-SQLMaxMemory)

#Restarting SQL Server 
Restart-Service -Name MSSQLSERVER -Force
Restart-Service -Name SQLSERVERAGENT -Force
}
}
catch {
    write-host "Caught an exception:" -ForegroundColor Red
    write-host "Exception Type: $($_.Exception.GetType().FullName)" -ForegroundColor Red
    write-host "Exception Message: $($_.Exception.Message)" -ForegroundColor Red
}
}

Start-Sleep -Seconds 30

if($VMType -like '*CI*'){
Try{
#===========================================================
#Changing drive letter
#===========================================================  
write-host "Changing drive letter"
$password = $localAdminPassword | ConvertTo-SecureString -AsPlainText -Force
$username = $localAdminUsername
$Cred = New-Object System.Management.Automation.PSCredential ($username, $password)
Invoke-Command -VMName $NewVMName -Credential $Cred -ScriptBlock{
Get-Partition -DiskNumber 1 | Set-Partition -NewDriveLetter D
}
}
Catch{
    write-host "Caught an exception:" -ForegroundColor Red
    write-host "Exception Type: $($_.Exception.GetType().FullName)" -ForegroundColor Red
    write-host "Exception Message: $($_.Exception.Message)" -ForegroundColor Red
}    
}

Start-Sleep -Seconds 1

Write-host "VM " $NewVMName " has been created succefully"
#===========================================================
#Clean-Up
#===========================================================  
Unregister-ScheduledTask -TaskName "DevScript"
Remove-Item -Path "C:\OP-CD-Win2016-SQL16"
Remove-Item -Path "C:\vmcreationfolder"
Stop-VM -Name VM-User-1 -TurnOff