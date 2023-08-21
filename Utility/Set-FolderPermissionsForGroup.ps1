<#

Funktionen slår upp alla undermappar i en mapp till en GridView.
De markerade mapparna skickas vidare till en funktion som sätter behörighet på mappen

#>

[cmdletbinding()]
param(
    [Parameter(Mandatory)][string]$basePath,
    [Parameter(Mandatory)][string]$Groupname,
    [Parameter()][switch]$WritePermission
)

function Set-FolderPermissions {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [string]$FolderPath,

        [Parameter(Mandatory = $true)]
        [string]$GroupName,

        [switch]$WritePermission
    )

    begin {
        if ( $WritePermission ) {
            Write-verbose "Setting Write permissions for $GroupName"
        } else {
            Write-Verbose "Setting Read permissions for $GroupName"
        }
        
    }

    process {
        foreach ($folder in $FolderPath) {
            if (Test-Path $folder -PathType Container) {

                $acl = Get-Acl $folder

                # Remove existing permissions for the group
                $existingRule = $acl.Access | Where-Object { $_.IdentityReference -eq $GroupName }
                if ($existingRule) {
                    $acl.RemoveAccessRule($existingRule)
                }

                # Add new permission rule based on WritePermission switch
                if ($WritePermission) {
                    $permission = [System.Security.AccessControl.FileSystemRights]::Modify
                } else {
                    $permission = [System.Security.AccessControl.FileSystemRights]::ReadAndExecute
                }

                $inheritanceFlag = [System.Security.AccessControl.InheritanceFlags]::ContainerInherit -bor [System.Security.AccessControl.InheritanceFlags]::ObjectInherit
                $propagationFlag = [System.Security.AccessControl.PropagationFlags]::None

                $accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule($GroupName, $permission, $inheritanceFlag, $propagationFlag, "Allow")
                $acl.AddAccessRule($accessRule)

                Set-Acl $folder $acl

                Write-Verbose "Permissions updated for $folder"

            } else {
                Write-Host "$folder is not a valid folder path."
            }
        }
    }

    end {
        Write-Verbose "Done setting permissions for $GroupName."
    }
}

Get-ChildItem -Path $basePath -Directory | Select-Object -ExpandProperty FullName | Out-GridView -PassThru | Set-FolderPermissions -GroupName $Groupname -WritePermission:$WritePermission