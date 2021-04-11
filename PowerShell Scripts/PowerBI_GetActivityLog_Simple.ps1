# Has a hard dependency on the Power BI MGMT cmdlets, install the modules if needed
# https://www.powershellgallery.com/packages/MicrosoftPowerBIMgmt

#Set up variables for further use
#Number of Days for which to collect ActivityLogs, interpret as From (MaxDays)..To(MinDays), in iterations of 1 day
$MaxDays = 30
$MinDays = 1

#Basepath to store output files
$Path = "C:\Temp\PBIAdmin\ssbi\"

#When using "SPN" as $AuthMode, make sure these variables are filled in correctly, to authenticate unattended
#$TenantID = 'InsertTenantGUIDHere' #GUID for the current tenant
#$ApplicationID = 'InserApplicationGUIDHere' #GUID for the App Registration Application ID
#$Secret = 'InsertSecretHere' #Client Secret for the App Registration SPN

# Prepare variables for session connection. Needs to be replaced with Certificate authentication
#$Password = ConvertTo-SecureString $Secret -AsPlainText -Force
#$Credential = New-Object PSCredential $ApplicationID, $password

# Connect to Power BI with credential of Power BI Service Administrator / Service Principal
# When using an SPN you have to provide the TenantID, as defaulting to MyOrg will not work
# Connect-PowerBIServiceAccount -ServicePrincipal -Credential $Credential -Tenant $TenantID

#When not using SPN, run this command to bring up the authentication dialog
Connect-PowerBIServiceAccount

#MaxDays..MinDays, loops over as an array over every object
$MaxDays..$MinDays |
ForEach-Object {
    $Date = (((Get-Date).Date).AddDays(-$_))
    
    # -Format needs to be used, API only accepts a certain datetime format
    $StartDate = (Get-Date -Date ($Date) -Format yyyy-MM-ddTHH:mm:ss)
    $EndDate = (Get-Date -Date ((($Date).AddDays(1)).AddMilliseconds(-1)) -Format yyyy-MM-ddTHH:mm:ss)
    
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
}