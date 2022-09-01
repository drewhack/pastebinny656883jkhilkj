## Ver 8.31.22 - JB
$TestResult = $true
$TestResultTxt = ""
$verbosepreference = 'continue'
$Computer = Get-ImmyComputer
$PCName = $Computer.Name
$SerialNumber = $Computer.SerialNumber
Import-ImmyITGlueModule
$APIEndpoint = "https://api.itglue.com"
$APIKEy =  "REDACTED"
#$orgID = $ITGlueOrgID (No longer needed, extracted from configuration information)
#Settings IT Glue logon information
Add-ITGlueBaseURI -base_uri $APIEndpoint
Add-ITGlueAPIKey $APIKEy

# Check C: volume Bitlocker status
$BDEStatus = $Computer | Invoke-ImmyCommand {
    $BitlockStatus = (Get-BitlockerVolume -MountPoint "C:").protectionstatus
    Return $BitlockStatus
}

# If status is Off write Status to ITG, Else continue
if($BDEStatus -like "Off") {
    Write-Host "Bitlocker is Off"
    $Testresult = $false
    $TestResultTxt = $TestResultTxt + "Error: Bitlocker is Off`n"
    
    # Check for Coinfigurationin ITG
    $ComputerConfiguration = (Get-ITGlueConfigurations -filter_name $PCName -filter_serial_number $SerialNumber -filter_configuration_status_id "34770").data
    #Write-Verbose ($ComputerConfiguration | ConvertTo-Json -depth 5)
    if(!$ComputerConfiguration) {
        Write-Warning "Unable to find ITGlue configuration with SerialNumber: $SerialNumber"
        $Testresult = $false
        $TestResultTxt = $TestResultTxt + "Error: No IT Glue Configuration found`n"
    } Else { # Configuration(s) found
        # Added for loop to handle and update duplicate records
        foreach($RecordID in $ComputerConfiguration) {
            $orgID = ($RecordID).attributes | select organization-id
            $orgID = $orgID.'organization-id'
            $Record = $RecordID.id
            Write-Host "Org ID = $orgID, Record ID = $Record"
            $PasswordObjectName = "BDE Status: Off (ITG ID: $Record)"
            $PasswordObject = @{
                type = 'passwords'
                attributes = @{
                    name = $PasswordObjectName
                    password = "Bitlocker is Off"
                    notes = "Bitlocker status for $PCName"
                    resource_id = $Record
                    resource_type = "Configuration"
                }
            }
    
            # If there is NOT an existing ITG PW record name match, set test to fail.
            $ExistingPasswordAsset = (Get-ITGluePasswords -filter_organization_id $orgID -filter_name $PasswordObjectName).data
            if(!$ExistingPasswordAsset) {
                $Testresult = $false
                $TestResultTxt = $TestResultTxt + "Error: No IT Glue BDE status Off entry found`n"
            } else {
                $TestResultTxt = $TestResultTxt + "IT Glue BDE status Off entry found`n"
            }
            # Test or Deploy logic
            switch($Method) {
                "Test" {
                    #Return $Testresult
                }
                "Set" {   
                    # If the BDE ststus PW entry does not exist, write it to ITG.
                    if(!$ExistingPasswordAsset) {
                        Write-Output "Creating new Bitlocker Status Entry" -ForegroundColor yellow
                        $ITGNewPassword = New-ITGluePasswords -organization_id $orgID -data $PasswordObject
                    }
                }
            }
            # Reset objects each loop
            $PasswordObject = @{}
            $ExistingPasswordAsset = @{}
        } # End for loop
    } # End Update ITG withg Status

    } Else { # Else BDE is ON
    Write-Host "Bitlocker is On"
    $TestResultTxt = $TestResultTxt + "Bitlocker is On`n"

    # Get Bitlocker Volumes from Computer
    $BDEVols = $Computer | Invoke-ImmyCommand {
        $BitlockVolumes = Get-BitLockerVolume
        return $BitlockVolumes
    }
    #Write-Host "BitlockeVolumes:"
    #$BDEVols

    #Check that keys exist
    $BDEKeys = $Computer | Invoke-ImmyCommand {
        #$VerbosePreference = 'continue'
        $BitlockVolumes = Get-BitLockerVolume
        #Write-Verbose ($BitlockVolumes | fl * | Out-String)
        $RecoveryPasswords = $BitlockVolumes |  %{
            ([string]($_.KeyProtector).RecoveryPassword).Trim()
        } | ?{$_}
        return $RecoveryPasswords
    }
    if(!$BDEKeys) {
        Write-Host "No Bitlocker RecoveryPasswords Found"
        $Testresult = $false
        $TestResultTxt = $TestResultTxt + "Error: No Keys found`n"
    }
    #Write-Host "RecoveryPasswords:"
    #$BDEKeys

    if(!$SerialNumber) {
        $Testresult = $false
        $TestResultTxt = $TestResultTxt + "Error: Unable to determin PC Serail Number`n"
        Write-Error "Unable to determine SerialNumber for computer $PCName" -ErrorAction Stop
    }

    #The script uses the following line to find the correct asset by serialnumber, match it, and connect it if found.
    #$ComputerConfiguration = (Get-ITGlueConfigurations -organization_id $orgID -filter_name $PCName -filter_serial_number $SerialNumber -filter_configuration_status_id "34770").data
    $ComputerConfiguration = (Get-ITGlueConfigurations -filter_name $PCName -filter_serial_number $SerialNumber -filter_configuration_status_id "34770").data
    #Write-Verbose ($ComputerConfiguration | ConvertTo-Json -depth 5)
    if(!$ComputerConfiguration) {
        Write-Warning "Unable to find ITGlue configuration with SerialNumber: $SerialNumber"
        $Testresult = $false
        $TestResultTxt = $TestResultTxt + "Error: No IT Glue Configuration found`n"
    } Else { # Configuration found
        $TestResultTxt = $TestResultTxt + "IT Glue Configuration found`n"
        foreach($RecordID in $ComputerConfiguration) {
            #Extract Org ID
            $orgID = ($RecordID).attributes | select organization-id
            $orgID = $orgID.'organization-id'
            $Record = $RecordID.id
            Write-Host "Org ID = $orgID, Record ID = $Record"

            # Loop through BDE volumes
            foreach($BitlockVolume in $BDEVols) {
                If($BitlockVolume.ProtectionStatus -eq $true) {
                    $BitlockVolume.MountPoint
                    #$BitlockVolume.KeyProtector
                    $KeyTypes = $BitlockVolume.KeyProtector | Select-Object -ExpandProperty KeyProtectorType | Select-String "RecoveryPassword"
                    $PWLine = $KeyTypes.LineNumber -1
                    $BDEKeyID = ($BitlockVolume.KeyProtector.KeyProtectorId)[$PWLine]
                    $BDEKeyID = $BDEKeyID.Trim("{","}")
                    $BDEKey1 = ($BitlockVolume.KeyProtector.RecoveryPassword)[$PWLine]
                    Write-Host "Processing Bitlocker Volume $BitlockVolume Key ID $BDEKeyID"
                    #$BDEKey1
                    $PasswordObjectName = "BDE Key ID: $BDEKeyID (ITG ID: $Record)"
                    $PasswordObject = @{
                        type = 'passwords'
                        attributes = @{
                            name = $PasswordObjectName
                            password = ($BitlockVolume.KeyProtector.RecoveryPassword)[$PWLine]
                            notes = "Bitlocker key for $PCName - $($BitlockVolume.MountPoint) `n Key ID: $BDEKeyID"
                            resource_id = $Record
                            resource_type = "Configuration"
                        }
                    }
                    #Write-Verbose ($PasswordObject | ConvertTo-Json -depth 5)

                    # If there is NOT an existing IT Glue record, fail test.
                    $ExistingPasswordAsset = (Get-ITGluePasswords -filter_organization_id $orgID -filter_name $PasswordObjectName).data
                    if(!$ExistingPasswordAsset) {
                        $Testresult = $false
                        $TestResultTxt = $TestResultTxt + "Error: No IT Glue Key entry found`n"
                    } Else {
                        $TestResultTxt = $TestResultTxt + "IT Glue Key entry found`n"
                    }

                    # If there IS an existing IT Glue record that needs updating, fail test.
                    if($ExistingPasswordAsset) {
                        $ExistingPassword = (Get-ITGluePasswords -id $ExistingPasswordAsset.Id).data.attributes.password
                        if($ExistingPassword -ne ($BitlockVolume.KeyProtector.RecoveryPassword)[$PWLine]) {
                            $Testresult = $false
                            $TestResultTxt = $TestResultTxt + "Error: IT Glue key entry requires update`n"
                        } else {
                            $TestResultTxt = $TestResultTxt + "IT Glue key entry is up to date`n"
                        }
                    }

                    switch($Method) {
                        "Test" {
                            #Return $Testresult
                        }
                        "Set" {   
                            #If the Asset does not exist, we edit the body to be in the form of a new asset, if not, we update if needed.
                            if(!$ExistingPasswordAsset) {
                                Write-Output "Creating new Bitlocker Password" -ForegroundColor yellow
                                $TestResultTxt = $TestResultTxt + "Creating Bitlocker Password entry in IT Glue`n"
                                $ITGNewPassword = New-ITGluePasswords -organization_id $orgID -data $PasswordObject

                                #Remove BDE Status record
                                $PasswordObjectNameStatus = "BDE Status: Off (ITG ID: $Record)"
                                $ExistingPasswordAssetStatus = (Get-ITGluePasswords -filter_organization_id $orgID -filter_name $PasswordObjectNameStatus).data
                                if($ExistingPasswordAssetStatus) {
                                    $ITGRemoveStatus = Remove-ITGluePasswords -id $ExistingPasswordAssetStatus.id
                                }

                            } else {
                                If ($Testresult -ne $true) {
                                    Write-Output "Updating Bitlocker Password" -ForegroundColor Yellow
                                    $TestResultTxt = $TestResultTxt + "Updating Bitlocker Password entry in IT Glue`n"
                                    $ITGNewPassword = Set-ITGluePasswords -id $ExistingPasswordAsset.id -data $PasswordObject
                                }
                            }
                        }
                    }
                    # Reset objects each loop
                    $PasswordObject = @{}
                    $ExistingPasswordAsset = @{}
                } # End if Protection On
            } # End BDE volumes loop
        } # End multiple record for loop
    } # End Else found configuration
} # End of BDE Status check Else On

Write-Host $TestResultTxt
Write-Host $Testresult
switch($Method) {
    "Test" {
        if($Testresult -eq "True") {return $true} else {return $false}
    }
    "Set" {
        # No actions here
    }
}
