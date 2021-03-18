<#
        .SYNOPSIS
        A PowerShell script to create Billable Time for various Work Item types in Service Manager using Outlook calendar appointment and meetings

        .DESCRIPTION
        Using Outlook to manage Billable Time in Service Manager allows analysts to utilise the graphical nature of Outlook calendaring.
        This provides a more intuitive interface for managing Billable Time and encourages good time management practises.
        
        The script is dependent on the Syliance.BillableTime.Extension (paid version) which expands Billable Time to all Work Items in Service Manager
        and not just Incidents, it also enables comments and other features that expand on the basic built in Billable Time features in Service Manager.
        
        The script can be run in a normal PowerShell window, but is designed to be used with Service Manager console tasks.
#>

# Get Service Manager installation directory from the local registry and import the Service Manager PowerShell Cmdlets
$GetInstallDirectory = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\System Center\2010\Service Manager\Setup' -Name InstallDirectory
$SMPSModule = $GetInstallDirectory.InstallDirectory + "Powershell\System.Center.Service.Manager.psd1"
Import-Module $SMPSModule

# Global variables.
$SCSMServer = "SCSMServerName"

# How far back in the Outlook calendar to process in days
$DaysToProcess = 7

# Test the connection to Service Manager before running the rest of the script
Try {
  Get-SCSMAnnouncement -ComputerName $SCSMServer
  Write-Host "Successfully Connected to $SCSMServer"
}
Catch {
  Write-Host "Could not connect to $SCSMServer. Please check your connection and try again."
Exit
}

# Extract the default Date/Time formatting from the local computer's "Culture" settings, and then create the format to use when parsing the date/time information pulled from Active Directory
Function Set-LocalDateFormat {
    Param ($ConvertDate)
    $CultureDateTimeFormat = (Get-Culture).DateTimeFormat
    $DateFormat = $CultureDateTimeFormat.ShortDatePattern
    $TimeFormat = $CultureDateTimeFormat.LongTimePattern
    $DateTimeFormat = "$DateFormat $TimeFormat"
    $DateTime = [DateTime]::ParseExact($ConvertDate,$DateTimeFormat,[System.Globalization.DateTimeFormatInfo]::InvariantInfo,[System.Globalization.DateTimeStyles]::None)
    $DateTime
}

# Information from Outlook calendar meetings and apointments will be used to create Billable Time in Service Manager accociated to the parent Work Item
Function New-BillableTimeRecord {
    param ($Subject, $Start, $Duration)

# Filter incoming Subject strings to remove additional, non-Work Item number related characters.
If ($Subject.length -in 6..8) {
    $SubjectID = $Subject -replace '(^.{0}[A-Za-z]*\d+)([^a-zA-Z0-9]+)(.*)','$1'
}

# If the subject contains a description, we can use this for the BillableTime comments field. 
Else {
    $SubjectID = $Subject -replace '(^.{0}[A-Za-z]*\d+)([^a-zA-Z0-9]+)(.*)','$1'
    $Comment = $Subject -replace '(^.{0}[A-Za-z]*\d+)([^a-zA-Z0-9]+)(.*)','$3'
}
    
# Get the Work Item type by prefix.
switch -Wildcard ($SubjectID) {
"PR*" {$WorkItem = Get-SCSMClassInstance -ComputerName $SCSMServer -Class (Get-SCSMClass -ComputerName $SCSMServer -Name System.WorkItem.Problem) | ?{$_.ID -eq $SubjectID}; Break} 
"IR*" {$WorkItem = Get-SCSMClassInstance -ComputerName $SCSMServer -Class (Get-SCSMClass -ComputerName $SCSMServer -Name System.WorkItem.Incident) | ?{$_.ID -eq $SubjectID}; Break} 
"SR*" {$WorkItem = Get-SCSMClassInstance -ComputerName $SCSMServer -Class (Get-SCSMClass -ComputerName $SCSMServer -Name System.WorkItem.ServiceRequest) | ?{$_.ID -eq $SubjectID}; Break} 
"RR*" {$WorkItem = Get-SCSMClassInstance -ComputerName $SCSMServer -Class (Get-SCSMClass -ComputerName $SCSMServer -Name System.WorkItem.ReleaseRecord) | ?{$_.ID -eq $SubjectID}; Break} 
"CR*" {$WorkItem = Get-SCSMClassInstance -ComputerName $SCSMServer -Class (Get-SCSMClass -ComputerName $SCSMServer -Name System.WorkItem.ChangeRequest) | ?{$_.ID -eq $SubjectID}; Break}
    default { Return } 
    }

# Get the BillableTime relationship.
$BillableTimeRel = Get-SCRelationship -ComputerName $SCSMServer -name System.WorkItemHasBillableTime
 
# Set the properties for the BillableTime entry in a hash table.
$BillableTime = @{
    Id = [string]$([guid]::NewGUID())
    TimeInMinutes = $Duration
    StartDate = $Start
    LastUpdated = Set-LocalDateFormat(get-date -format "dd/MM/yyyy hh:mm:ss")
    Comment = $Comment
    }

Try
{
    # Create the BillableTime entry and Relate it to the Work Item.
    $BillabletimeRelInstance = New-SCRelationshipInstance -ComputerName $SCSMServer -RelationshipClass $BillableTimeRel -Source $WorkItem -TargetClass (Get-SCSMClass -ComputerName $SCSMServer -Name Syliance.BillableTime.Extension) -TargetProperty $BillableTime -PassThru
}
Catch {
    Write-Host "There was a problem getting the details of $SubjectID. Please ensure that $SubjectID exists in Service Manager."
    Return "False"
}

# Get the BillableTime object that was just created by guid.
$BillableTimeInstance = Get-SCSMClassInstance -ComputerName $SCSMServer -Id $BillabletimeRelInstance.TargetObject.Id.Guid

# Get currently logged on user object to relate to the Billabletime.
Try {
    $WorkingUser = Get-SCSMClassInstance -ComputerName $SCSMServer -Class (Get-SCSMClass -ComputerName $SCSMServer -Name System.User) | ?{$_.UserName -eq $env:UserName}
}
Catch {
    Write-Host "There was a problem getting the details of the username property for $env:UserName. Please ensure that $env:UserName exists in Service Manager."
    Return "False"
}

# Get the 'Has Working User' Relationship.
$HasWorkingUser = Get-SCRelationship -ComputerName $SCSMServer -name System.WorkItem.BillableTimeHasWorkingUser

# Create the 'Has Working User' relationship between the BilliableTime and the User
Try {

    New-SCRelationshipInstance -ComputerName $SCSMServer -RelationshipClass $HasWorkingUser -Source $BillableTimeInstance -Target $WorkingUser
}
Catch {
    Write-Host "There was a problem getting the details of $SubjectID. Please ensure that $SubjectID exists in Service Manager."
    Return "False"
}
Return "True", $SubjectID
}
 
# Main script.
# Get work Item related appointments from the local instance of Outlook. The script will use the default profile.

Write-Host "Processing the last $DaysToProcess days of data."
Write-Host "============================================="

$StartDate = (Get-Date).AddDays(-$DaysToProcess).ToString("dd/MM/yyyy")
$EndDate = (Get-Date).AddDays(+1).ToString("dd/MM/yyyy")

# For performance reasons, use the 'Restrict' method to filter the list of appointments, rather than only using Where-Object.
$Filter = "@SQL=""urn:schemas:calendar:dtstart"" >= '${StartDate}' AND ""urn:schemas:calendar:dtstart"" <= '${EndDate}'"
 
# Create a connection to the local Outlook instance and load all the calendar items in the default calendar. 
Add-type -assembly "Microsoft.Office.Interop.Outlook" | out-null 
$OlFolders = "Microsoft.Office.Interop.Outlook.OlDefaultFolders" -as [type]  
$Outlook = new-object -comobject outlook.application 
$Namespace = $Outlook.GetNameSpace("MAPI") 
$Folder = $Namespace.getDefaultFolder($OlFolders::olFolderCalendar).Items
$Folder.Sort("[Start]", 0) 
$Folder.IncludeRecurrences = 1
$FilterResults = $Folder.Restrict($Filter)

# Use Where-Object to filter using Regex as it is much simpler to do this in Powershell than it is in SQL or VB.
$Results = $FilterResults | Where-Object { $_.Subject -match "^\p{P}{0}[A-Za-z]{1}\d+|^\p{P}{0}[A-Za-z]{2}\d+" }
$TotalResults = ($Results | Measure-Object).Count
$TimeRemaining = $TotalResults * 20
$TotalTime = New-TimeSpan -Seconds $TimeRemaining
$TotalTimeMinutes = [math]::Round($TotalTime.TotalMinutes)
Write-Host "Estimated time to complete - $TotalTimeMinutes minutes"

# Iterate through each calendar item in the array for Billable Time to be created for each item. 
# If the Billable Item is created successfully, update the subject on the associated calendar item with an apostrophe prefix to mark the item as being processed and to prevent duplication of billable time.
$i = 0
Foreach ($AppItem in $Results) {
$i++
$PercentComplete = ($i / $TotalResults) * 100
$PercentComplete = [math]::Round($PercentComplete)

$ReturnState = New-BillableTimeRecord $AppItem.Subject $AppItem.Start $AppItem.Duration
Write-Host "$PercentComplete% complete - Successfully created Billable Time for" $ReturnState[1]
If ($ReturnState[0] -eq "True") {
    $UpdateSubject = $AppItem.Subject
    $UpdateSubject = "!" + $UpdateSubject
    $AppItem.Subject = $UpdateSubject
    $AppItem.Save()
}
}
