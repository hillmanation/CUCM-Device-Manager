##################################
## Jabber Device Manager Script ##
## --Created By Jake Hillman--- ##
## ----GE Edison Works SDS----- ##
## -------v1.2 - 2/22/23------- ##
##################################

Import-Module ActiveDirectory

## Get todays date and a list of disabled users older than 120 days
$date = Get-Date
$users = Get-ADuser -filter * -Properties Name,samAccountName,Description,LastLogonDate,Created,Enabled | where { $_.Enabled -eq $true -and $_.samAccountName -notlike "svc_*" -and $_.samAccountName -notlike "root" -and $_.samAccountName -notlike "*service*" -and $_.samAccountName -notlike "*svc*" -and $_.LastLogonDate -ne $NULL -and (New-TimeSpan -start $_.LastLogonDate -end $date).Days -gt 120 -and (New-TimeSpan -start $_.Created -end $date).Days -gt 120 } | Select-Object Name,samAccountName,Description,LastLogonDate,Created,Enabled
$users += Get-ADUser -filter * -Properties Name,samAccountName,Description,LastLogonDate,Created,Enabled -SearchBase "OU=Disabled Users, OU=IT Users,OU=col-dev,DC=col-dev,DC=ge,DC=COM" -SearchScope Subtree | Select-Object Name,samAccountName,Description,LastLogonDate,Created,Enabled | where { $_.samAccountName -notin $users.samAccountName }
$expireduserlist = @()

# Enter CUCM Server
$cucmServer = "CUCM-1.col-dev.ge.com"
# Enter CUCM Version
$ver = '12.5'

# Enter your AD creds or your login to the /CCMAdmin page
if ( ! ($cred) ) {
$cred = Get-Credential -Message "Enter CUCM Credentials"
}

## Some stuff to ensure certificate rquests are ignored
add-type @"
    using System.Net;
    using System.Security.Cryptography.X509Certificates;
    public class TrustAllCertsPolicy : ICertificatePolicy {
        public bool CheckValidationResult(
            ServicePoint srvPoint, X509Certificate certificate,
            WebRequest request, int certificateProblem) {
            return true;
        }
    }
"@
[System.Net.ServicePointManager]::CertificatePolicy = New-Object TrustAllCertsPolicy
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls -bor [Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls12
[System.Net.ServicePointManager]::Expect100Continue = $false

## Function to wrap the xml queries with default syntax
Function Axl-Wrapper {
    Param ($body)

$wrap = @"
<?xml version="1.0" encoding="UTF-8"?>
<soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" xmlns:ns="http://www.cisco.com/AXL/API/12.0">
<soapenv:Header/>
    <soapenv:Body>
        $body
    </soapenv:Body>
</soapenv:Envelope>
"@

    return $wrap
}

## Function for sending the axl queries to the CUCM database
Function CUCM-Query {
    Param($query)

    $request = Axl-Wrapper $query

    try {
    $result = Invoke-RestMethod -Method Post -Uri "https://$cucmServer`:8443/axl/" -Headers @{'Content-Type'='text/xml';'SOAPAction'='CUCM:DB ver=' + $ver} -Body $request -Credential $cred
    } catch {
    $_ | select -ExpandProperty ErrorDetails | Select -ExpandProperty Message
    }

    ## Return result UNFORMATTED, will need formatting based on the query being handled
    return $result
}

## Get a list of phone devices for a specified user
Function Get-User-Devices {
    Param($userid)

$devicequery = @"
<ns:getUser sequence="?">
             <!--You have a CHOICE of the next 2 items at this level-->
             <userid>$userid</userid>
             <returnedTags uuid="?">
                 <associatedDevices>
                 </associatedDevices>
             </returnedTags>
      </ns:getUser>
"@

    $qresult = CUCM-Query $devicequery

    $output = $qresult.Envelope.Body.getUserResponse.return.user.associatedDevices.device
    
    return $output
}

## Remove all phone devices for a specified user
Function Remove-User-Devices {
    Param($userid)

$devicequery = @"
<ns:updateUser sequence="?">
            <userid>$userid</userid>
                <associatedDevices>
                </associatedDevices>
        </ns:updateUser>
"@

    $qresult = CUCM-Query $devicequery

    $output = $qresult.Envelope.Body.getUserResponse.return.user.associatedDevices.device

    return $output
}

## Remove a phone device from CUCM database
Function Remove-Phone {
    Param($devicename)

$devicequery = @"
<ns:removePhone>
     <name>$devicename</name>
  </ns:removePhone>
"@

    $qresult = CUCM-Query $devicequery
}

## Get phone device info from CUCM database
Function Get-Phone {
    Param($devicename)

$devicequery = @"
	<ns:getPhone>
		<name>$devicename</name>
	</ns:getPhone>
"@

    $qresult = CUCM-Query $devicequery

    $output = $qresult.Envelope.Body.getPhoneResponse.return.phone

    return $output
}

## Build a list of disabled users that have a device associated with them in CUCM
ForEach ($user in $users) {
    
    $userdevices = Get-User-Devices $user.samAccountName
    If ($userdevices -ne $NULL) {
        $expireduserlist += [pscustomobject]@{ NAME=$user.Name; samAccountName=$user.samAccountName; Enabled=$user.Enabled; LastLogonDate=$user.LastLogonDate; CreatedOn=$user.Created; Description=$user.Description; RegisteredDevices=$userdevices }
    }
}

## Display the list of expired user's devices
$expireduserlist | FT -AutoSize | out-host

## Prompt for removal of user devices
ForEach ($expireduser in $expireduserlist) {

    $username = $expireduser.samAccountName
    $name = $expireduser.NAME
    echo "`nRemove Devices for $name`? Y/N`n-----------------------------------------------"
    $expireduser
    $removeresponse = Read-Host "Y/N"

    ## Remove user devices if prompted
    If ($removeresponse.ToUpper() -eq 'Y') {
    
        Remove-User-Devices $username
        ## Check if the removal worked
        $check = Get-User-Devices $username
        If ($check -eq $NULL) { echo "Devices removed successfully for $username!" }
        Else { Write-Warning "Devices unable to be removed for $username!!!"; Continue }

        ForEach ($device in $expireduser.RegisteredDevices) {
            
            ## Remove the phone device from CUCM
            Remove-Phone $device

            $check2 = Get-Phone $device
            If ($check2 -eq $NULL) { echo "Phone $device for $username removed successfully from database!" }
            Else { Write-Warning "Phone $device for $username unable to be removed!`nPlease manually delete this phone from CUCM in order to free the license!!!"; Continue }
        }
        
    }

}