# Coded by A Tech With a Hammer - https://www.techwithahammer.com
#  Last Update March 25, 2019

[CmdletBinding()]
Param(
    # The OU Parameter will only work off the DistinguishedName of the OU
    [Parameter(Position=0,mandatory=$true)]
    [string]$OU,
    [Parameter(Position=1)]
    [bool]$TrialRun=$true
)

Import-Module ActiveDirectory
# Get all users in the specified OU
try {
    $Users = Get-ADUser -filter * -SearchBase $OU -Properties ProxyAddresses,Mail,DisplayName
} catch {
    Write-Host -foregroundcolor red "Unable to enumerate users in $OU. Use the distinguished name for the OU in this field."
    exit 1
}

# Parse each user
Foreach ($User in $Users) {
    # List the user details
    Write-Host $("Name: " + $User.DisplayName)
    Write-Host $("UPN: " + $User.UserPrincipalName)
    
    # Verify the mail attribute is not blank
    if ($User.mail -ne "" -and $User.mail -ne $null) {
        $UserEmail = $User.mail
        $UserEmailExist = $false # Used to verify the address existsa proxy address
        
        # Process the proxy addresses
        if ($User.ProxyAddresses.count -lt 1) { # Set if there are no proxy addresses
            $User.ProxyAddresses = @($("SMTP:" + $UserEmail))
            $UserEmailExist = $true
        }
        else {
            for ($i = 0; $i -lt ($User.ProxyAddresses).count; $i++) {
                if ($User.ProxyAddresses[$i].split(':')[0] -eq "SMTP") { # To ensure only the SMTP email addresses are manipulated
                    $UserProxyEmail = $User.ProxyAddresses[$i].split(':')[1]
                    if ($UserProxyEmail -eq $UserEmail) {
                        $User.ProxyAddresses[$i] = $("SMTP:" + ($User.ProxyAddresses[$i].split(':')[1]))
                        #Write-Host $("`tSMTP:" + ($User.ProxyAddresses[$i].split(':')[1]))
                        $UserEmailExist = $True
                    }
                    else {
                        $User.ProxyAddresses[$i] = ($User.ProxyAddresses[$i] = $("smtp:" + ($User.ProxyAddresses[$i].split(':')[1])))
                        #Write-Host $("`tsmtp:" + ($User.ProxyAddresses[$i].split(':')[1]))
                    }
                }
            }
            If (!($UserEmailExist)) { # Add if the address was not encountered
                $User.ProxyAddresses += $("SMTP:" + $UserEmail)
                #Write-Host $("`tSMTP:" + $UserEmail)
            }
            
            # List all the ProxyAddresses that will be set on the user
            $User.ProxyAddresses | Foreach-Object {
                Write-host $("`t" + $_)
            }
            
        }
        
        # Assume the user object was updated, and if not running in TrialRun mode
        if (!($TrialRun)) { Set-ADUser -Instance $User }
    }
    else {
        $UserUPN = $User.UserPrincipalName
        Write-Host -ForegroundColor RED -BackgroundColor BLACK "User $UserUPN does not have the mail attribute set. We are not assuming the format so please go set the Mail box in the first tab of the user."
    }
}

if ($TrialRun) { Write-host -ForegroundColor GREEN -BackgroundColor BLACK "`r`nThis was run in trial mode, if the attributes above look correct than run the script again with `"-TrialRun $false`"`r`n" }
