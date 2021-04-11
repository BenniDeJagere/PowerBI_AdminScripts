# Has a hard dependency on the Power BI MGMT cmdlets, script will check for existing modules
# https://www.powershellgallery.com/packages/MicrosoftPowerBIMgmt

#Set up variables for further use
#Number of Days for which to collect ActivityLogs, interpret as From (MaxDays)..To(MinDays), in iterations of 1 day
$MaxDays = 30
$MinDays = 1

#Basepath to store output files
$Path = "C:\Temp\PBIAdmin\ssbi\"

#Change to "SPN" when using Service Principal to authenticate, make sure to fill in the variables below
$AuthMode = "Login" 

#When using "SPN" as $AuthMode, make sure these variables are filled in correctly, to authenticate unattended
$TenantID = 'InsertTenantGUIDHere' #GUID for the current tenant
$ApplicationID = 'InserApplicationGUIDHere' #GUID for the App Registration Application ID
$Secret = 'InsertSecretHere' #Client Secret for the App Registration SPN

#PsLog Function, needs to be called on every check
function PrintPsLog() { 
    Param 
    (                            
          [parameter(Mandatory = $True)] [ValidateSet ('ACTION', 'ERROR', 'PRESENT' , 'SUCCESS', 'WARNING',  'INFO' ) ][string]$LogStatus                    
        , [Parameter(Mandatory = $True)][String]$LogDescription   
    )   
      
    $DateNow = Get-Date -Format 'dd/MM/yyyy hh:mm:ss:ff'
    switch ($LogStatus) {
        "ACTION" {
            Write-Host "$DateNow > $LogStatus  > $LogDescription" -ForegroundColor Cyan                                                                 
        }
        "ERROR" {
            Write-Host "$DateNow > $LogStatus   > $LogDescription" -ForegroundColor Red                  
        }
        "PRESENT" {
            Write-Host "$DateNow > $LogStatus > $LogDescription" -ForegroundColor Gray   
        }
        "SUCCESS" {
            Write-Host "$DateNow > $LogStatus > $LogDescription" -ForegroundColor Green   
        }
        "WARNING" {
            Write-Host "$DateNow > $LogStatus > $LogDescription"  -ForegroundColor Yellow                  
        }
        "INFO" {                
            Write-Host "$DateNow > $LogStatus    > $LogDescription" -ForegroundColor Yellow   
        }            
    }
}

PrintPsLog -LogStatus "INFO" -LogDescription "Finished declaring PSLog Function"

#Check for import status on necessary modules
$modules = @("MicrosoftPowerBIMGMT" , "DataGateway")

foreach ( $m in $modules ) 
{
    PrintPsLog -LogStatus "ACTION" -LogDescription "Checking for module $m"
    if (Get-Module -ListAvailable -Name $m) {
        PrintPsLog -LogStatus "SUCCESS" -LogDescription "Module $m is already imported."
    } 
    else {
        Install-Module -Name $m -Force -Scope CurrentUser
        Import-Module $m
        PrintPsLog -LogStatus "SUCCESS" -LogDescription "Module $m is now imported."
    }
}

PrintPsLog -LogStatus "INFO" -LogDescription "Finished checking for modules"

#Establish Connection to Power BI
if ($AuthMode -eq "SPN")
{
    # Prepare variables for session connection. Needs to be replaced with Certificate authentication
    $Password = ConvertTo-SecureString $Secret -AsPlainText -Force
    $Credential = New-Object PSCredential $ApplicationID, $password

    # Connect to Power BI with credential of Power BI Service Administrator / Service Principal
    # When using an SPN you have to provide the TenantID, as defaulting to MyOrg will not work
    try {
        PrintPsLog -LogStatus "ACTION" -LogDescription "Connecting to PowerBI - Connect-PowerBIServiceAccount with SPN"
        $Connection = Connect-PowerBIServiceAccount -ServicePrincipal -Credential $Credential -Tenant $TenantID
        if ($Connection) {
            PrintPsLog -LogStatus "SUCCESS" -LogDescription "Connection to PowerBI successful - Connect-PowerBIServiceAccount with SPN"
        } else {
        PrintPsLog -LogStatus "ERROR" -LogDescription "Connection to PowerBI failed - Connect-PowerBIServiceAccount with SPN"
        }
    } catch {
    PrintPsLog -LogStatus "ERROR" -LogDescription "Connection to PowerBI failed - Connect-PowerBIServiceAccount with SPN"
    }
}
elseif ($AuthMode -eq "Login")
{
    #When not using SPN, log in through the statement below
    try {
        PrintPsLog -LogStatus "ACTION" -LogDescription "Connecting to PowerBI - Connect-PowerBIServiceAccount through dialog"
        $Connection = Connect-PowerBIServiceAccount
        if ($Connection) {
            PrintPsLog -LogStatus "SUCCESS" -LogDescription "Connection to PowerBI succesvol - Connect-PowerBIServiceAccount through dialog"
        } else {
        PrintPsLog -LogStatus "ERROR" -LogDescription "Connection to PowerBI failed - Connect-PowerBIServiceAccount through dialog"
        }
    } catch {
    PrintPsLog -LogStatus "ERROR" -LogDescription "Connectie to PowerBI failed - Connect-PowerBIServiceAccount through dialog"
    }
}
else
{
    PrintPsLog -LogStatus "ERROR" -LogDescription "Connection to PowerBI failed - Please provide SPN or Login as a parameter to the AuthMode parameter"
}

PrintPsLog -LogStatus "INFO" -LogDescription "Finished authentication logic"

#MaxDays..MinDays, loops over as an array over every object
$MaxDays..$MinDays |
ForEach-Object {
    $Date = (((Get-Date).Date).AddDays(-$_))
    
    # -Format needs to be used, API only accepts a certain datetime format
    $StartDate = (Get-Date -Date ($Date) -Format yyyy-MM-ddTHH:mm:ss)
    $EndDate = (Get-Date -Date ((($Date).AddDays(1)).AddMilliseconds(-1)) -Format yyyy-MM-ddTHH:mm:ss)
    
    PrintPsLog -LogStatus "ACTION" -LogDescription "Collecting activities for $StartDate until $EndDate"

    #For sake of completeness the single line below could replace the entire process of calling the API, and looping over the continuationUri
    #$auditlogs = Get-PowerBIActivityEvent -StartDateTime $StartDate -EndDateTime $EndDate -ResultType JsonString | ConvertFrom-Json
    
    #Create an empty array to hold all records for the current day in context
    $activities = @()        
    
    #Initiate the first call to the API, the field continuationURI will be used to use the same session for the next batch of results
    $auditlogs = Invoke-PowerBIRestMethod -Url "https://api.powerbi.com/v1.0/myorg/admin/activityevents?startDateTime='$StartDate'&endDateTime='$EndDate'" -Method Get | ConvertFrom-Json
    $activities += $auditlogs.activityEventEntities

    #When continuationUri is NULL (blank), all records for the current day in context have been retrieved
    if($auditlogs.continuationUri) {
        do {
            $auditlogs = Invoke-PowerBIRestMethod -Url $auditlogs.continuationUri -Method Get | ConvertFrom-Json
            $activities += $auditlogs.activityEventEntities
        } until(-not $auditlogs.continuationUri)    
    }
    
    #Only select the fields we currently need, to avoid breaking metadata in the steps further down the process (Compare it to why we don't do SELECT * in T-SQL)
    $selectedActivities = $activities | Select-Object Id, RecordType, CreationTime, Operation, OrganizationId, UserType, UserKey, Workload, UserId, ClientIP, UserAgent, 
            Activity, ItemName, WorkspaceName, DashboardName, DatasetName, ReportName, CapacityId, CapacityName, WorkspaceId, 
            ObjectId, DashboardId, DatasetId, ReportId, IsSuccess, ReportType, RequestId, ActivityId, DistributionMethod, 
            ConsumptionMethod, DataflowName, DataflowId, OrgAppPermission 
    
    #Create the filename for current Day in context, and export the results to a .csv file
    $fileName = "$(Get-Date -Date $Date -Format yyyyMMdd).csv"
    $filePath = "$($Path)$($fileName)"

    $selectedActivities | Export-Csv -Path $filePath -NoTypeInformation -Encoding UTF8 -Force

    #Collect the resultset count, to report back to the output
    $activitiesCount = $activities.count
    if ($activities.count -gt 0) {
    PrintPsLog -LogStatus "SUCCESS" -LogDescription "Collected $activitiesCount activities for $Date"
    } 
    else {
       PrintPsLog -LogStatus "ERROR" -LogDescription "No activities found in Activity Logs for $Date"
    }
}
PrintPsLog -LogStatus "INFO" -LogDescription "Finished collecting activities"

#Disconnect PowerBI Service Account
PrintPsLog -LogStatus "ACTION" -LogDescription "Disconnect Power BI Session"
Disconnect-PowerBIServiceAccount
PrintPsLog -LogStatus "SUCCESS" -LogDescription "Ended Power BI Session"