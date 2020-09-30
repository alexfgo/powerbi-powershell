# This sample script calls the Power BI API to programmatically exports the workspace and all
# its dashboards, reports and datasets and imports in a new workplace (in the same account or in a different one)

# Instructions:
# 1. Install PowerShell (https://msdn.microsoft.com/en-us/powershell/scripting/setup/installing-windows-powershell) 
#    and the Azure PowerShell cmdlets (Install-Module AzureRM)
# 2. Run PowerShell as an administrator
# 3. Follow the instructions below to fill in the client ID
# 4. Change PowerShell directory to where this script is saved
# 5. > ./export-import-Workspace.ps1

# Parameters - fill these in before running the script!
# ======================================================

# AAD Client ID
# To get this, go to the following page and follow the steps to provision an app
# https://dev.powerbi.com/apps
# To get the sample to work, ensure that you have the following fields:
# App Type: Native app
# Redirect URL: urn:ietf:wg:oauth:2.0:oob
#  Level of access: check all boxes

$sourceClientId = "c2c7429e-3174-4a8b-a323-00d2e6efb511" 
$targetClientId = "9845cae9-922c-431e-b83f-1b740a71ef98" 

# End Parameters =======================================

# Calls the Active Directory Authentication Library (ADAL) to authenticate against AAD
function GetAuthToken([String] $clientId) {
    if (-not (Get-Module AzureRm.Profile)) {
        Import-Module AzureRm.Profile
    }

    $redirectUri = "urn:ietf:wg:oauth:2.0:oob"

    $resourceAppIdURI = "https://analysis.windows.net/powerbi/api"

    $authority = "https://login.microsoftonline.com/common/oauth2/authorize";

    $authContext = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext" -ArgumentList $authority

    $authResult = $authContext.AcquireToken($resourceAppIdURI, $clientId, $redirectUri, "Always")

    return $authResult
}

function get_groups_path([String] $group_id) {
    if ($group_id -eq "me") {
        return "myorg"
    }
    else {
        return "myorg/groups/$group_ID"
    }
}

function itemExistsByName([String] $find, [System.Collections.Hashtable] $collection) {
    Write-Host "Param1:"$find"/"
    Write-Host "Param2:"$collection"/"
    Foreach ($item in $collection) {
        $name = $item.name
        Write-Host "Name:"$name"/"
        Write-Host "Key:"$find"/"
        if ($find -ieq $name) {
            return 1
        }
    }
    return 0
}

# PART 1: Authentication
# ==================================================================
Write-Host "Authenticate using SOURCE Account"
$sourceToken = GetAuthToken($sourceClientId)
Write-Host "Authenticate using TARGET Account"
$targetToken = GetAuthToken($targetClientId)

Add-Type -AssemblyName System.Net.Http
$temp_path_root = "$PSScriptRoot\pbi-copy-workspace-temp-storage"

# Building Rest API header with authorization token
$source_auth_header = @{
    'Content-Type'  = 'application/json'
    'Authorization' = $sourceToken.CreateAuthorizationHeader()
}

$target_auth_header = @{
    'Content-Type'  = 'application/json'
    'Authorization' = $targetToken.CreateAuthorizationHeader()
}

# Prompt for user input
# ==================================================================
# Get the list of groups that the user is a member of
$uri = "https://api.powerbi.com/v1.0/myorg/groups/"
$all_source_groups = (Invoke-RestMethod -Uri $uri –Headers $source_auth_header –Method GET).value

# Ask for the source workspace name
$source_group_ID = ""
while (!$source_group_ID) {
    $source_group_name = Read-Host -Prompt "Enter the name of the SOURCE workspace you'd like to export from"

    if ($source_group_name -eq "My Workspace") {
        $source_group_ID = "me"
        break
    }

    Foreach ($group in $all_source_groups) {
        if ($group.name -eq $source_group_name) {
            Write-Host "Encontrado!"
            Write-Host "Workspace name:" $group.name
            Write-Host "Workspace id:" $group.id
            
            # if ($group.isReadOnly -eq "True") {
            #     "Invalid choice: you must have edit access to the group"
            #     break
            # } else {
            $source_group_ID = $group.id
            #     break
            # }
        }
    }

    if (!$source_group_id) {
        "Please try again, making sure to type the exact name of the group"  
    } 
}

# Ask for the target workspace name
$target_group_ID = "" 
while (!$target_group_id) {
    try {
        $target_group_name = Read-Host -Prompt "Enter the name of the TARGET workspace you'd like to import to"
        $uri = " https://api.powerbi.com/v1.0/myorg/groups"
        $all_target_groups = (Invoke-RestMethod -Uri $uri –Headers $target_auth_header –Method GET).value

        if ($target_group_name -eq "My Workspace") {
            $target_group_ID = "me"
            break
        }
    
        Foreach ($group in $all_target_groups) {
            if ($group.name -eq $target_group_name) {

                Write-Host "Encontrado!"
                Write-Host "Workspace name:" $group.name
                Write-Host "Workspace id:" $group.id


                if ($group.isReadOnly -eq "True") {
                    "Invalid choice: you must have edit access to the group"
                    break
                }
                else {
                    $target_group_ID = $group.id
                    break
                }
            }
        }
    
        if (!$target_group_ID) {
            "Please try again, making sure to type the exact name of the target group"  
        } 
    }
    catch { 
        "Could not found a group with that name. Please try again and make sure the name already exists"
        "More details: "
        Write-Host "StatusCode:" $_.Exception.Response.StatusCode.value__ 
        Write-Host "StatusDescription:" $_.Exception.Response.StatusDescription
        continue
    }
}

$proceed = Read-Host -Prompt "Do you want to proceed with the export/import operation? (y/n)"

if ($proceed -eq "y" -or $proceed -eq "Y") {
    # PART 2: Copying reports and datasets using Export/Import PBIX APIs
    # ==================================================================
    $report_id_mapping = @{}      # mapping of old report ID to new report ID
    $dataset_id_mapping = @{}     # mapping of old model ID to new model ID
    $failure_log = @()  
    $source_group_path = get_groups_path($source_group_ID)
    $target_group_path = get_groups_path($target_group_ID)
    $processed_datasets = 0
    $processed_reports = 0

    $uri = "https://api.powerbi.com/v1.0/$source_group_path/datasets/"
    $datasets = (Invoke-RestMethod -Uri $uri –Headers $source_auth_header –Method GET).value

    $uri = "https://api.powerbi.com/v1.0/$target_group_path/datasets/"
    $target_datasets = (Invoke-RestMethod -Uri $uri –Headers $target_auth_header –Method GET).value

    Write-Host "# of datasets found: " $datasets.count
    Write-Host " "
    Read-Host -Prompt "Press ENTER to proceed..."

    # == Export/import the DATASETS as REPORTS (this step creates the datasets)
    # for each dataset, try exporting and importing the PBIX
    if (!(Test-Path $temp_path_root)) {
        Write-Host "Export directory created: " $temp_path_root
        New-Item -Path $temp_path_root -ItemType Directory 
    }
    "=== Exporting PBIX files to copy datasets..."

    $uri = "https://api.powerbi.com/v1.0/$source_group_path/reports/"
    $reports = (Invoke-RestMethod -Uri $uri –Headers $source_auth_header –Method GET).value

    Foreach ($report in $reports) {
        $report_id = $report.id
        $dataset_id = $report.datasetId
        $report_name = $report.name

        # if report not exists (it's not a dataset), so skip the processing
        $ds_exists = 0
        Foreach ($ds in $datasets) {
            $ds_name = $ds.name
            if ($report_name -ieq $ds_name) {
                $ds_exists = 1
            }
        }
        if ($ds_exists -eq 0) {
            Write-Host "Report is not a dataset, skipped until next phase of the process: " $report_name
            continue;
        }

        $temp_path = "$temp_path_root\$report_name.pbix"

        if (Test-Path $temp_path) {
            Foreach ($ds in $target_datasets) {
                $ds_name = $ds.name
                if ($ds_name -ieq $report_name) {
                    $dataset_id_mapping[$dataset_id] = $ds.id
                    Write-Host "Report was previously exported, skipped: " $report_name
                }
            }
            continue;
        }

        "== Exporting $report_name with id: $report_id to $temp_path"
        $uri = "https://api.powerbi.com/v1.0/$source_group_path/reports/$report_id/Export"
        try {
            Invoke-RestMethod -Uri $uri –Headers $source_auth_header –Method GET -OutFile "$temp_path"
        }
        catch [Exception] {
            Write-Host $_.Exception
            Write-Host "== Error: failed to export PBIX"
            Write-Host "= HTTP Status Code:" $_.Exception.Response.StatusCode.value__ 
            Write-Host "= HTTP Status Description:" $_.Exception.Response.StatusDescription
            Write-Host "= This report and dataset cannot be copied, skipping: " $report_name
            continue
        }
        
        try {
            "== Importing $report_name to target workspace"
            $uri = "https://api.powerbi.com/v1.0/$target_group_path/imports/?datasetDisplayName=$report_name.pbix&nameConflict=Abort"

            # Here we switch to HttpClient class to help POST the form data for importing PBIX
            $httpClient = New-Object System.Net.Http.Httpclient $httpClientHandler
            $httpClient.DefaultRequestHeaders.Authorization = New-Object System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", $targetToken.AccessToken);
            $packageFileStream = New-Object System.IO.FileStream @($temp_path, [System.IO.FileMode]::Open)
            
            $contentDispositionHeaderValue = New-Object System.Net.Http.Headers.ContentDispositionHeaderValue "form-data"
            $contentDispositionHeaderValue.Name = "file0"
            $contentDispositionHeaderValue.FileName = $file_name
    
            $streamContent = New-Object System.Net.Http.StreamContent $packageFileStream
            $streamContent.Headers.ContentDisposition = $contentDispositionHeaderValue
            
            $content = New-Object System.Net.Http.MultipartFormDataContent
            $content.Add($streamContent)

            $response = $httpClient.PostAsync($Uri, $content).Result
    
            if (!$response.IsSuccessStatusCode) {
                $responseBody = $response.Content.ReadAsStringAsync().Result
                "= This report cannot be imported to target workspace. Skipping..."
                $errorMessage = "Status code {0}. Reason {1}. Server reported the following message: {2}." -f $response.StatusCode, $response.ReasonPhrase, $responseBody
                throw [System.Net.Http.HttpRequestException] $errorMessage
            } 
            
            # save the import IDs
            $import_job_id = (ConvertFrom-JSON($response.Content.ReadAsStringAsync().Result)).id

            # wait for import to complete
            $upload_in_progress = $true
            while ($upload_in_progress) {
                $uri = "https://api.powerbi.com/v1.0/$target_group_path/imports/$import_job_id"
                $response = Invoke-RestMethod -Uri $uri –Headers $target_auth_header –Method GET
                
                if ($response.importState -eq "Succeeded") {
                    "Publish succeeded!"
                    # update the report and dataset mappings
                    $report_id_mapping[$report_id] = $response.reports[0].id
                    Write-Host "Old reportId: " $report_id "- New reportId: " $report_id_mapping[$report_id]
                    $dataset_id_mapping[$dataset_id] = $response.datasets[0].id
                    Write-Host "Old datasetId: " $dataset_id "- New datasetId: " $dataset_id_mapping[$dataset_id]
                    $processed_datasets += 1
                    break
                }

                if ($response.importState -ne "Publishing") {
                    "Error: publishing failed, skipping this. More details: "
                    $response
                    break
                }
                
                Write-Host -NoNewLine "."
                Start-Sleep -s 5
            }            
        }
        catch [Exception] {
            Write-Host $_.Exception
            Write-Host "== Error: failed to import PBIX"
            Write-Host "= HTTP Status Code:" $_.Exception.Response.StatusCode.value__ 
            Write-Host "= HTTP Status Description:" $_.Exception.Response.StatusDescription
            continue
        }
    }

    Write-Host $processed_datasets " exported/imported of " $datasets.count
    Write-Host " "
    Write-Host "# of reports found: " $reports.count
    Write-Host " "
    Read-Host -Prompt "Press ENTER to proceed..."

    # For My Workspace, filter out reports that I don't own - e.g. those shared with me
    if ($source_group_ID -eq "me") {
        Write-Host "Filtering out reports not own by the user"
        Write-Host " "
        $reports_temp = @()
        Foreach ($report in $reports) {
            if ($report.isOwnedByMe -eq "True") {
                $reports_temp += $report
            }
        }
        $reports = $reports_temp
        Write-Host "# of reports left after the filter: " $reports
        Write-Host " "
        Read-Host -Prompt "Press ENTER to proceed..."
    }

    # == Export/import the reports that are built on PBIXes (this step creates the datasets)
    # for each report, try exporting and importing the PBIX
    "=== Exporting PBIX files to copy datasets..."
    Foreach ($report in $reports) {
        $dataset_to_delete_id = ""
        $report_id = $report.id
        $dataset_id = $report.datasetId
        $report_name = $report.name
        $temp_path = "$temp_path_root\$report_name.pbix"

        if (Test-Path $temp_path) {
            Write-Host "Report was previously imported, skipped: " $report_name
            continue;
        }

        $rebind = 0
        # # only rebind if this dataset has already been seen
        if ($dataset_ID_mapping[$dataset_id]) {
            $rebind = 1
        }

        "== Exporting $report_name with id: $report_id to $temp_path"
        $uri = "https://api.powerbi.com/v1.0/$source_group_path/reports/$report_id/Export"
        try {
            Invoke-RestMethod -Uri $uri –Headers $source_auth_header –Method GET -OutFile "$temp_path"
        }
        catch [Exception] {
            Write-Host $_.Exception
            Write-Host "== Error: failed to export PBIX"
            Write-Host "= HTTP Status Code:" $_.Exception.Response.StatusCode.value__ 
            Write-Host "= HTTP Status Description:" $_.Exception.Response.StatusDescription
            Write-Host "= This report and dataset cannot be copied, skipping: " $report_name
            continue
        }
        
        try {
            "== Importing $report_name to target workspace"
            $uri = "https://api.powerbi.com/v1.0/$target_group_path/imports/?datasetDisplayName=$report_name.pbix&nameConflict=Abort"

            # Here we switch to HttpClient class to help POST the form data for importing PBIX
            $httpClient = New-Object System.Net.Http.Httpclient $httpClientHandler
            $httpClient.DefaultRequestHeaders.Authorization = New-Object System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", $targetToken.AccessToken);
            $packageFileStream = New-Object System.IO.FileStream @($temp_path, [System.IO.FileMode]::Open)
            
            $contentDispositionHeaderValue = New-Object System.Net.Http.Headers.ContentDispositionHeaderValue "form-data"
            $contentDispositionHeaderValue.Name = "file0"
            $contentDispositionHeaderValue.FileName = $file_name
    
            $streamContent = New-Object System.Net.Http.StreamContent $packageFileStream
            $streamContent.Headers.ContentDisposition = $contentDispositionHeaderValue
            
            $content = New-Object System.Net.Http.MultipartFormDataContent
            $content.Add($streamContent)

            $response = $httpClient.PostAsync($Uri, $content).Result
    
            if (!$response.IsSuccessStatusCode) {
                $responseBody = $response.Content.ReadAsStringAsync().Result
                "= This report cannot be imported to target workspace. Skipping..."
                $errorMessage = "Status code {0}. Reason {1}. Server reported the following message: {2}." -f $response.StatusCode, $response.ReasonPhrase, $responseBody
                throw [System.Net.Http.HttpRequestException] $errorMessage
            } 
            
            # save the import IDs
            $import_job_id = (ConvertFrom-JSON($response.Content.ReadAsStringAsync().Result)).id

            # wait for import to complete
            $upload_in_progress = $true
            while ($upload_in_progress) {
                $uri = "https://api.powerbi.com/v1.0/$target_group_path/imports/$import_job_id"
                $response = Invoke-RestMethod -Uri $uri –Headers $target_auth_header –Method GET
                
                if ($response.importState -eq "Succeeded") {
                    "Publish succeeded!"
                    # update the report and dataset mappings
                    $report_id_mapping[$report_id] = $response.reports[0].id
                    Write-Host "Old reportId: " $report_id "- New reportId: " $report_id_mapping[$report_id]

                    # holds the dataset only if will not rebind
                    if ($rebind -eq 0) {
                        $dataset_id_mapping[$dataset_id] = $response.datasets[0].id
                        Write-Host "Old datasetId: " $dataset_id "- New datasetId: " $dataset_id_mapping[$dataset_id]
                    }
                    # holds the dataset id for deletion 
                    else {
                        $dataset_to_delete_id = $response.datasets[0].id
                    }
                    $processed_reports += 1
                    break
                }

                if ($response.importState -ne "Publishing") {
                    "Error: publishing failed, skipping this. More details: "
                    $response
                    break
                }
                
                Write-Host -NoNewLine "."
                Start-Sleep -s 5
            }

            
        }
        catch [Exception] {
            Write-Host $_.Exception
            Write-Host "== Error: failed to import PBIX"
            Write-Host "= HTTP Status Code:" $_.Exception.Response.StatusCode.value__ 
            Write-Host "= HTTP Status Description:" $_.Exception.Response.StatusDescription
            continue
        }

        # if this dataset has already been seen, the report was created online and needs to be rebound to dataset
        if ($rebind -eq 1) {
            Write-Host "Rebinding report: " $report.name " - its dataset was already created."
            $new_report_id = $report_id_mapping[$report_id]
            $uri = " https://api.powerbi.com/v1.0/$target_group_path/reports/$new_report_id/Rebind"
            $new_dataset_id = $dataset_id_mapping[$dataset_id]
            $body = "{`"datasetId`": `"$new_dataset_id`"}"
            Write-Host "New dataset: " $body
            try {
                Invoke-RestMethod -Uri $uri –Headers $target_auth_header –Method POST -Body $body
                Write-Host "Delete unnecessary dataset: " $dataset_to_delete_id
                $delete_uri = "https://api.powerbi.com/v1.0/$target_group_path/datasets/$dataset_to_delete_id"
                Invoke-RestMethod -Uri $delete_uri –Headers $target_auth_header –Method DELETE
            }
            catch [Exception] {
                Write-Host $_.Exception
                Write-Host "== Error: failed to rebind PBIX"
                Write-Host "= HTTP Status Code:" $_.Exception.Response.StatusCode.value__ 
                Write-Host "= HTTP Status Description:" $_.Exception.Response.StatusDescription
                Write-Host "= This report and dataset cannot be copied, skipping: " $report_name
                continue
            }
        }
    }

    Write-Host $processed_reports " exported/imported of " $reports.count
    Write-Host " "
    Read-Host -Prompt "Press ENTER to proceed..."

    # PART 3: Copy dashboards
    # ==================================================================
    "=== Cloning dashboards" 
    # get all dashboards in a workspace
    $uri = "https://api.powerbi.com/v1.0/$source_group_path/dashboards/"
    $dashboards = (Invoke-RestMethod -Uri $uri –Headers $source_auth_header –Method GET).value

    # For My Workspace, filter out dashboards that I don't own - e.g. those shared with me
    if ($source_group_ID -eq "me") {
        $dashboards_temp = @()
        Foreach ($dashboard in $dashboards) {
            if ($dashboard.isReadOnly -ne "True") {
                $dashboards_temp += $dashboard
            }
        }
        $dashboards = $dashboards_temp
    }

    Foreach ($dashboard in $dashboards) {
        #$dashboard_id = $dashboard.id
        $dashboard_name = $dashboard.displayName

        "== Cloning dashboard: $dashboard_name"

        # create new dashboard in the target workspace
        $uri = "https://api.powerbi.com/v1.0/$target_group_path/dashboards"
        $body = "{`"name`":`"$dashboard_name`"}"
        $response = Invoke-RestMethod -Uri $uri –Headers $target_auth_header –Method POST -Body $body
        #$target_dashboard_id = $response.id
        #
        # " = Cloning individual tiles..." 
        # # copy over tiles:
        # $uri = "https://api.powerbi.com/v1.0/$source_group_path/dashboards/$dashboard_id/tiles"
        # $response = Invoke-RestMethod -Uri $uri –Headers $source_auth_header –Method GET 
        # $tiles = $response.value
        # Foreach ($tile in $tiles) {
        #     try {
        #         $tile_id = $tile.id
        #         $tile_report_Id = $tile.reportId
        #         $tile_dataset_Id = $tile.datasetId
        #         if ($tile_report_id) { $tile_target_report_id = $report_id_mapping[$tile_report_id] }
        #         if ($tile_dataset_id) { $tile_target_dataset_id = $dataset_id_mapping[$tile_dataset_id] }

        #         # clone the tile only if a) it is not built on a dataset or b) if it is built on a report and/or dataset that we've moved
        #         if (!$tile_report_id -Or $dataset_id_mapping[$tile_dataset_id]) {
        #             $uri = " https://api.powerbi.com/v1.0/$source_group_path/dashboards/$dashboard_id/tiles/$tile_id/Clone"
        #             $body = @{}
        #             $body["TargetDashboardId"] = $target_dashboard_id
        #             $body["TargetWorkspaceId"] = $target_group_id
        #             if ($tile_report_id) { $body["TargetReportId"] = $tile_target_report_id } 
        #             if ($tile_dataset_id) { $body["TargetModelId"] = $tile_target_dataset_id } 
        #             $jsonBody = ConvertTo-JSON($body)
        #             $response = Invoke-RestMethod -Uri $uri –Headers $target_auth_header –Method POST -Body $jsonBody
        #             Write-Host -NoNewLine "."
        #         } else {
        #             $failure_log += $tile
        #         } 
            
        #     } catch [Exception] {
        #         "Error: skipping tile..."
        #         Write-Host $_.Exception
        #     }
        # }
        "Done!"
    }

    Read-Host -Prompt "Press ENTER to proceed..."

    # "Cleaning up temporary files"
    # Remove-Item -path $temp_path_root -Recurse

}