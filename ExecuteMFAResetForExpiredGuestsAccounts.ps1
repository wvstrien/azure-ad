Clear-Host
Add-Type -Path "C:\Handytools\Microsoft.SharePoint2016.csom\Microsoft.SharePoint.Client.dll"
Add-Type -Path      "C:\Handytools\Microsoft.SharePoint2016.csom\Microsoft.SharePoint.Client.Runtime.dll"

#Specification:
# per invited guest:
# 1. Check when guest account created: if recent (< 2 months) then most likely it works, and MFA is still valid
# 2. If created longer ago; then look up lost logon to <tenant> Azure AD. Again: if recent, then most likely it works. Otherwise the MFA may
# have expired.

Function Get-FileName($initialdirectory)
{
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = "CSV (*.csv)| *.csv"
    $OpenFileDialog.ShowDialog() | Out-Null
    $OpenFileDialog.filename
}

$reportOutInfoCollection= @()

# Open connection with AzureAD, to allow check on whether already guest account is present in Azure AD
Import-Module AzureAD
try {
    $tenantDetail = Get-AzureADTenantDetail
} catch {
    #Not connected yet
    try {
         Connect-AzureAD -ErrorAction Stop
         $tenantDetail = Get-AzureADTenantDetail
    } catch {
         exit
    }
}

Import-Module MSOnline
try {
    Connect-MsolService  -ErrorAction Stop
} catch {
    exit
}

$verifiedDomainNames = @()
foreach($verifiedDomain in $tenantDetail.VerifiedDomains) {
    $verifiedDomainNames = $verifiedDomainNames + $verifiedDomain.Name.ToLower()
    if ($verifiedDomain._Default -eq $true) {
        $mainDomain = $verifiedDomain.Name
    }
}

# Load the input file for batch submission of new B2B guest requests.
$newGuestsListFile = Get-FileName
$newGuestsList = Import-Csv $newGuestsListFile
$nrOfNewGuests = 1
if ($newGuestsList -is [array]) {
    $nrOfNewGuests = $newGuestsList.Count
}

write-host "Nr of guests to check: " + $nrOfNewGuests

$progresscount = 0
$global:expiredDate = (get-date).addmonths(-2)

ForEach ($item in $newGuestsList){
    $progresscount++
    $guestEmail = $($item.Email).Trim().ToLower()
    $guestName = $($item.Name).Trim().ToLower()
    $ActivityInfo = "MFA necessity check handling of '$guestEmail' ...Completed: {0,3}%   " -f ([math]::floor(($progresscount*100)/$nrOfNewGuests))
    Write-Progress -Activity $ActivityInfo -Status "$($progresscount) of $($nrOfNewGuests)" -PercentComplete (($progresscount*100)/$nrOfNewGuests)
    try {
       $reportOutGuestInfo = new-object psobject        
       $reportOutGuestInfo | add-member noteproperty -name "Email" -value $guestEmail       
       $guestEmailDomain = $guestEmail.Split("@")[1].ToLower()
       if ($verifiedDomainNames.Contains($guestEmailDomain) -ne $true) {     
           #Convert to UPN for guest
           $guestUpn = $guestEmail.Replace('@','_') + "#EXT#@" + $mainDomain
           $userADObject =  Get-AzureADUser -Filter "userPrincipalName eq '$guestUpn'" | Select DisplayName, UserPrincipalName, Mail, UserType, UserState, UserStateChangedOn, RefreshTokensValidFromDateTime
           # Can be that person is administrated with another UPN as the presented email. Check whether the email is administrated as proxy address.
           # Note: as check on email, also for non-verified domains on the email value; not the 'guestified' UPN
           if ($userADObject -eq $null) {
               $userADObject = Get-AzureADUser -Filter "proxyAddresses/any(c:c eq 'smtp:$guestEmail')" | Select DisplayName, UserPrincipalName, Mail, UserType, UserState, UserStateChangedOn, RefreshTokensValidFromDateTime
           }                                                                                                                                                
           if ($userADObject -and $userADObject.UserType -eq "Guest") {  
               $guestUpn = $($userADObject.UserPrincipalName)
               $reportOutGuestInfo | add-member noteproperty -name "Name" -value $($userADObject.DisplayName)
               $reportOutGuestInfo | add-member noteproperty -name "O365 Logon" -value $($userADObject.Mail)        
               $reportOutGuestInfo | add-member noteproperty -name "UPN in ASML system" -value $guestUpn
               $reportOutGuestInfo | add-member noteproperty -name "RefreshTokensValidFromDateTime" -value $($userADObject.RefreshTokensValidFromDateTime)                                                   
               if ($userADObject.UserState -ne "PendingAcceptance" -and (get-date $userADObject.RefreshTokensValidFromDateTime) -lt $expiredDate) {
                   $reportOutGuestInfo | add-member noteproperty -name "Reset MFA" -value "Yes"
                   Write-Host "Reset-MsolStrongAuthenticationMethodByUpn -UserPrincipalName" $guestUpn
                   Reset-MsolStrongAuthenticationMethodByUpn -UserPrincipalName $guestUpn                                                              
               } else {
                   $reportOutGuestInfo | add-member noteproperty -name "Reset MFA" -value "No"
               }
            } else {  
                $reportOutGuestInfo | add-member noteproperty -name "UserState" -value "No guest account"                                 
            }
         } else {
            $appliedAction = "Skipped"
            $actionDetails = "Email qualifies as verifiedDomain in <tenant> Azure AD, and person must be authenticated via member account"
            Write-Host "Email qualifies as verifiedDomain in <tenant> Azure AD, and person must be authenticated via member account: " $guestEmail                  
            $reportOutGuestInfo | add-member noteproperty -name "UserState" -value "No guest account"                                                
         }
         $reportOutInfoCollection+=$reportOutGuestInfo              
      } catch { 
          $appliedAction = "Failed"
          $actionDetails = $($_.Exception.Message)
          Write-Host "Guest request  failure for: " $guestEmail -foregroundcolor red 
          write-host "$($_.Exception.Message)" -foregroundcolor red 
     }
}

if ($reportOutInfoCollection.Count -gt 0) {
    $reportOutInfoCollectionFile = $newGuestsListFile.replace(".csv", " (last-logon-state  $(get-date -f yyyy-MM-dd)).csv")
    $reportOutInfoCollection |  Export-Csv $reportOutInfoCollectionFile -NoTypeInformation
}
