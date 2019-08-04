# Attention: do not start this script as long as the synchronization is active.

Install-Module MSOnline
Import-Module MSOnline
Connect-MsolService

# Get all users from an OU
$users = Get-ADUser -Filter { Name -Like "*" -and Enabled -eq $true } -Searchbase "OU=office365,DC=demo,DC=chrog,dc=net"

foreach ($user in $users){

    # Get Immutable IDs
    $wertad = [system.convert]::ToBase64String($user.ObjectGUID.ToByteArray()) 
    $wertaz = (Get-MsolUser -UserPrincipalName $user.UserPrincipalName).ImmutableID
    
    if ($wertad -ne $wertaz){
        $user.UserPrincipalName
        "AD Anchor:    " + $wertad
        "Azure Anchor: " + $wertaz

        # Create an OnMicrosoft email adress
        $bspStr = $user.UserPrincipalName.Split("@")
        $msemail = $bspStr[0] + "@chrognet.onmicrosoft.com"
        "E-Mail:       " + $msemail
        
        # Update azure account source anchor
        Set-MsolUserPrincipalName -UserPrincipalName $user.UserPrincipalName  -NewUserPrincipalName $msemail
        Set-MsolUser -UserPrincipalName $msemail -ImmutableId "$null"
        set-MsolUser -userprincipalname $msemail -immutableId $wertad
        Set-MsolUserPrincipalName -NewUserPrincipalName $user.UserPrincipalName -UserPrincipalName $msemail

        # Start delta-sync
        Start-ADSyncSyncCycle -PolicyType delta
        "-----------------------------------------------------------------"
    }
}
