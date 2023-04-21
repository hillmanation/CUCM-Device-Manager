##################################
## Jabber Device Manager Script ##
## --Created By Jake Hillman--- ##
## ----GE Edison Works SDS----- ##
## ------v2.6 - 4/14/23-------- ##
##################################

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName presentationframework
Add-Type -AssemblyName System.Drawing

## Get todays date
$date = Get-Date

# Enter CUCM Version and default servername
$ver = '12.5'
If (!(Test-Path -Path $env:USERPROFILE\AppData\Local\CUCMServer.txt)) { $cucm = "CUCM-1.col-dev.ge.com" }
Else { $cucm = Get-Content -Path $env:USERPROFILE\AppData\Local\CUCMServer.txt }

## Some stuff to ensure certificate requests/TLS version disparity are ignored
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

$iconBase64 = "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"
$iconBytes = [Convert]::FromBase64String($iconBase64)
# initialize a Memory stream holding the bytes
$stream = [System.IO.MemoryStream]::new($iconBytes, 0, $iconBytes.Length)

## Form Specific Functions ##
Function Form-Reset {
    $cred = $NULL
    $loginbutton.Enabled = $true
    $userlistbox.Items.Clear()
    $logoutbutton.Enabled = $false
    $pingicon.BackColor = 'DarkGray'
    $calendar.Enabled = $false
    $30daybutton.Enabled = $false
    $90daybutton.Enabled = $false
    $180daybutton.Enabled = $false
    $servertextbox.Enabled = $true
    $disabledOUcheck.Enabled = $false
    $disabledOUcheck.Checked = $false
    $disabledUsercheck.Enabled = $false
    $disabledUsercheck.Checked = $false
    $findusersbutton.Enabled = $false
    $removedevicesbutton.Enabled = $false
    $vertextbox.Enabled = $true
    Clear-Console
    Write-Console White "-----------------Jabber Device Manager-----------------`r`nPlease login to the CUCM server to manage user devices."
}

Function Clear-Console {
    $consolebox.Text = ""
}

Function Write-Console {
    param( 
        [Parameter(Mandatory = $true, Position = 0)]
        [System.Drawing.Color]$color,
        [Parameter(Mandatory = $true, Position = 1)]
        [string]$text
    )
    $consolebox.SelectionStart = $consolebox.TextLength
    $consolebox.SelectionLength = 0
    $consolebox.SelectionColor = $color
    $consolebox.AppendText($text)
    $consolebox.AppendText([Environment]::NewLine)
}

#### Application Functions for Jabber/CUCM handling ####
## Function to verify login credential to CUCM
Function CUCM-Login {
    $userlistbox.Items.Clear()
    
    $global:cucmServer = $servertextbox.Text
    If (Test-Path -Path $env:USERPROFILE\AppData\Local -ErrorAction SilentlyContinue) { $cucmServer | Out-File -FilePath $env:USERPROFILE\AppData\Local\CUCMServer.txt -ErrorAction SilentlyContinue }
    $global:cred = Get-Credential -Message "Enter CUCM Credentials"
    $global:ver = $vertextbox.Text
    $username = $cred.UserName
    Write-Console White "Logging into $cucmServer as $username..."

    try {
    $result = Invoke-RestMethod -Uri "https://$cucmServer`:8443/axl/" -Credential $cred
    } catch {
    $_ | Select -Expand ErrorDetails | Select -ExpandProperty Message
    }
    $loggedin = "The AXL Web Service is working and accepting requests."

    If ($result -ne "" -and $result.html.body.p -like "*$loggedin*") {

        Write-Console Green "Logged in as $username"
        $pingicon.BackColor = 'Lime'
        $loginbutton.Enabled = $false
        $servertextbox.Enabled = $false
        $logoutbutton.Enabled = $true
        $calendar.Enabled = $true
        $30daybutton.Enabled = $true
        $90daybutton.Enabled = $true
        $180daybutton.Enabled = $true
        $disabledOUcheck.Enabled = $true
        $disabledUsercheck.Enabled = $true
        $userlistbox.Enabled = $true
        $findusersbutton.Enabled = $true
        $vertextbox.Enabled = $false
    }
    Else { $pingicon.BackColor = 'Crimson'
        Write-Console Red "Login Failed!"
    }
}

Function User-Search {

    $ADProps = @("Name","samAccountName","Description","LastLogonDate","Created","Enabled")
    $SearchBase = $disabledOUselect.SelectedItem
    $beforedate = $calendar.Value
    $searchusers = @()

    $age = [Math]::Round(($date - $beforedate).TotalDays, 0)

    ## Find users that have not logged in since before the selected calendar date
    $searchusers = Get-ADuser -filter * -Properties $ADProps | where { $_.samAccountName -notlike "root" -and $_.samAccountName -notlike "*service*" -and $_.samAccountName -notlike "*svc*" -and $_.LastLogonDate -ne $NULL -and (New-TimeSpan -start $_.LastLogonDate -end $date).Days -gt $age } | Select-Object $ADProps
    ## Include all users in the disabled OU
    If ($disabledOUcheck.Checked -eq $true -and $SearchBase -ne $NULL) {
        $searchusers += Get-ADUser -filter * -Properties $ADProps -SearchBase $SearchBase -SearchScope Subtree | where { $_.samAccountName -notin $searchusers.samAccountName } | Select-Object $ADProps
    }
    ## Include all disabled users (since some may not be in the disabled OU)
    If ($disabledUsercheck.Checked -eq $true) {
        $searchusers += Get-ADuser -filter * -Properties $ADProps | where { $_.Enabled -eq $false } | where { $_.samAccountName -notin $searchusers.samAccountName } | Select-Object $ADProps
    }

    return $searchusers
}

Function Find-Expired-Users {
    $expireduserlist = @()
    $userlistbox.Items.Clear()

    Write-Console White "Generating list of expired users with Jabber devices in CUCM..."

    $users = User-Search
    
    ## Build a list of disabled users that have a device associated with them in CUCM
    ForEach ($user in $users) {
    
        $userdevices = Get-User-Devices $user.samAccountName
        If ($userdevices -ne $NULL) {
            $expireduserlist += [pscustomobject]@{ NAME=$user.Name; samAccountName=$user.samAccountName; Enabled=$user.Enabled; LastLogonDate=$user.LastLogonDate; CreatedOn=$user.Created; Description=$user.Description; RegisteredDevices=$userdevices }
        }
    }
    If ($expireduserlist.Count -eq 0) { Write-Console Yellow "None found, try changing selected options..."; return }

    ## Add items to the listbox
    for ($i=0; $i -lt $expireduserlist.length; $i++) {
        $itemname = '$item' + "$i"
        $itemname = New-Object System.Windows.Forms.ListViewItem($expireduserlist[$i].Name)

        # Add columns
        $columns | % { If ($_.Name -ne 'Name') { [void]$itemname.SubItems.Add([string]$expireduserlist[$i].($_.Name)) }}
        # Add item to list
        [void]$userlistbox.Items.Add($itemname)
    }
 
    Write-Console Green "Complete!"
    $removedevicesbutton.Enabled = $true
}

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

## Build list of checked users from $userlistbox
Function Selected-Users {
    $checkedlist = @()

    ForEach ( $item in $userlistbox.CheckedItems) {
        $item = $item.Subitems.Text
        $checkedlist += [pscustomobject]@{ NAME=$item[0]; samAccountName=$item[1]; RegisteredDevices=$item[6] }
    }
    
    return $checkedlist
}

## Run removal of user devices
Function CUCM-Device-Removal {
    
    ## Get info from the user select box
    $removallist = Selected-Users

    ForEach ($expireduser in $removallist) {

        $username = $expireduser.samAccountName
        $name = $expireduser.NAME
        $userdevices = Get-User-Devices $username

        Write-Console White "Removing CUCM devices from user $name..."
    
        Remove-User-Devices $username
        ## Check if the removal worked
        $check = Get-User-Devices $username
        If ($check -eq $NULL) { Write-Console Green "Devices removed successfully from $name!" }
        Else { Write-Console Red "Devices unable to be removed from $name!!!"; Continue }

        ForEach ($device in $userdevices) {
            Write-Console White "Removing user $name $device device from CUCM..."
            
            ## Remove the phone device from CUCM
            Remove-Phone $device

            $check2 = Get-Phone $device
            If ($check2 -eq $NULL) { Write-Console Green "Phone $device for $name removed successfully from database!" }
            Else { Write-Console Red "Phone $device for $name unable to be removed!`nPlease manually delete this phone from CUCM in order to free the license!!!"; Continue }
        }
    }
    Write-Console White "Done!`r`n"
    Find-Expired-Users
}

## Start building the windows form
$form = New-Object System.Windows.Forms.Form
$form.Text = 'Jabber Device Manager'
$form.Size = New-Object System.Drawing.Size(940,660)
$form.StartPosition = 'CenterScreen'
$form.FormBorderStyle = 'Fixed3d'
$form.MaximizeBox = $false
$form.Icon = [System.Drawing.Icon]::FromHandle(([System.Drawing.Bitmap]::new($stream).GetHIcon()))

$cancelbutton = New-Object System.Windows.Forms.Button
$cancelbutton.Location = New-Object System.Drawing.Point(415,600)
$cancelbutton.Size = New-Object System.Drawing.Size(70,23)
$cancelbutton.Font = New-Object System.Drawing.Font("Segoe UI",10,[System.Drawing.FontStyle]::Regular)
$cancelbutton.Text = 'Exit'
$cancelbutton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$form.CancelButton = $cancelbutton
$form.Controls.Add($cancelbutton)

$loginlabel = New-Object System.Windows.Forms.Label
$loginlabel.Location = New-Object System.Drawing.Point(10,5)
$loginlabel.Size = New-Object System.Drawing.Size(220,22)
$loginlabel.Font = New-Object System.Drawing.Font("Segoe UI",12,[System.Drawing.FontStyle]::Regular)
$loginlabel.Text = 'Login to CUCM Server:'
$form.Controls.Add($loginlabel)

$pingicon = New-Object System.Windows.Forms.Button
$pingicon.Location = New-Object System.Drawing.Point(897,7)
$pingicon.Size = New-Object System.Drawing.Size(18, 18)
$pingicon.FlatStyle = 'Flat'
$pingicon.FlatAppearance.BorderSize = 0
$pingicon.BackColor = 'DarkGray'
$pingpath = [System.Drawing.Drawing2D.GraphicsPath]::new()
$pingpath.AddEllipse(0, 0, $pingicon.ClientSize.Width, $pingicon.ClientSize.Height)
$pingicon.Region = [System.Drawing.Region]::new($pingpath)
$pingicon.Enabled = $false
$form.Controls.Add($pingicon)

$loginbutton = New-Object System.Windows.Forms.Button
$loginbutton.Location = New-Object System.Drawing.Point(10,30)
$loginbutton.Size = New-Object System.Drawing.Size(153,28)
$loginbutton.Font = New-Object System.Drawing.Font("Segoe UI",12,[System.Drawing.FontStyle]::Bold)
$loginbutton.ForeColor = 'Green'
$loginbutton.Text = 'Login'
$loginbutton.Add_Click({ CUCM-Login })
$form.Controls.Add($loginbutton)

$logoutbutton = New-Object System.Windows.Forms.Button
$logoutbutton.Location = New-Object System.Drawing.Point(167,30)
$logoutbutton.Size = New-Object System.Drawing.Size(153,28)
$logoutbutton.Font = New-Object System.Drawing.Font("Segoe UI",12,[System.Drawing.FontStyle]::Bold)
$logoutbutton.Text = 'Logout'
$logoutbutton.Enabled = $false
$logoutbutton.Add_Click({ Form-Reset })
$form.Controls.Add($logoutbutton)

$vertextbox = New-Object System.Windows.Forms.TextBox
$vertextbox.Location = New-Object System.Drawing.Point(435,5)
$vertextbox.Size = New-Object System.Drawing.Size(45,23)
$vertextbox.Font = New-Object System.Drawing.Font("Lucida Console",12,[System.Drawing.FontStyle]::Regular)
$vertextbox.Text = $ver
$form.Controls.Add($vertextbox)

$verlabel = New-Object System.Windows.Forms.Label
$verlabel.Location = New-Object System.Drawing.Point(325,5)
$verlabel.Size = New-Object System.Drawing.Size(120,20)
$verlabel.Font = New-Object System.Drawing.Font("Segoe UI",12,[System.Drawing.FontStyle]::Regular)
$verlabel.Text = "CUCM Version:"
$form.Controls.Add($verlabel)

$servertextbox = New-Object System.Windows.Forms.TextBox
$servertextbox.Location = New-Object System.Drawing.Point(670,5)
$servertextbox.Size = New-Object System.Drawing.Size(250,25)
$servertextbox.Font = New-Object System.Drawing.Font("Lucida Console",12,[System.Drawing.FontStyle]::Regular)
$servertextbox.Text = $cucm
$form.Controls.Add($servertextbox)

$serverlabel = New-Object System.Windows.Forms.Label
$serverlabel.Location = New-Object System.Drawing.Point(477,5)
$serverlabel.Size = New-Object System.Drawing.Size(215,20)
$serverlabel.Font = New-Object System.Drawing.Font("Segoe UI",12,[System.Drawing.FontStyle]::Regular)
$serverlabel.Text = "Enter CUCM Server FQDN:"
$form.Controls.Add($serverlabel)

$30daybutton = New-Object System.Windows.Forms.Button
$30daybutton.Location = New-Object System.Drawing.Point(10,110)
$30daybutton.Size = New-Object System.Drawing.Size(100,28)
$30daybutton.Text = '30 Days'
$30daybutton.Enabled = $false
$30daybutton.Add_Click({ $calendar.Value = [DateTime]((Get-Date).AddDays(-30)) })
$form.Controls.Add($30daybutton)

$90daybutton = New-Object System.Windows.Forms.Button
$90daybutton.Location = New-Object System.Drawing.Point(115,110)
$90daybutton.Size = New-Object System.Drawing.Size(100,28)
$90daybutton.Text = '90 Days'
$90daybutton.Enabled = $false
$90daybutton.Add_Click({ $calendar.Value = [DateTime]((Get-Date).AddDays(-90)) })
$form.Controls.Add($90daybutton)

$180daybutton = New-Object System.Windows.Forms.Button
$180daybutton.Location = New-Object System.Drawing.Point(220,110)
$180daybutton.Size = New-Object System.Drawing.Size(100,28)
$180daybutton.Text = '180 Days'
$180daybutton.Enabled = $false
$180daybutton.Add_Click({ $calendar.Value = [DateTime]((Get-Date).AddDays(-180)) })
$form.Controls.Add($180daybutton)

$disabledOUcheck = New-Object System.Windows.Forms.CheckBox
$disabledOUcheck.Location = New-Object System.Drawing.Point(10,144)
$disabledOUcheck.Size = New-Object System.Drawing.Size(250,20)
$disabledOUcheck.Font = New-Object System.Drawing.Font("Segoe UI",12,[System.Drawing.FontStyle]::Regular)
$disabledOUcheck.Text = "Include Disabled OU Users?"
$tooltipdisabledOU = New-Object System.Windows.Forms.ToolTip
$tooltipdisabledOU.SetTooltip($disabledOUcheck, "Check Users in the Disabled User OU selected below to see if they have devices in Jabber.")
$disabledOUcheck.Add_MouseEnter({$tooltipdisabledOU})
$disabledOUcheck.Enabled = $false
$disabledOUcheck.Add_CheckStateChanged({ $disabledOUselect.Enabled = $disabledOUCheck.Checked })
$form.Controls.Add($disabledOUcheck)

$disabledOUselect = New-Object System.Windows.Forms.ComboBox
$disabledOUselect.Location = New-Object System.Drawing.Point(10,168)
$disabledOUselect.Size = New-Object System.Drawing.Size(310,28)
$disabledOUselect.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$disabledOUselect.Font = New-Object System.Drawing.Font("Arial",10,[System.Drawing.FontStyle]::Regular)
$disabledOUselect.Enabled = $false
$form.Controls.Add($disabledOUselect)

$disabledUsercheck = New-Object System.Windows.Forms.CheckBox
$disabledUsercheck.Location = New-Object System.Drawing.Point(10,200)
$disabledUsercheck.Size = New-Object System.Drawing.Size(310,20)
$disabledUsercheck.Font = New-Object System.Drawing.Font("Segoe UI",12,[System.Drawing.FontStyle]::Regular)
$disabledUsercheck.Text = "Include All Disabled Users?"
$tooltipdisableduser = New-Object System.Windows.Forms.ToolTip
$tooltipdisableduser.SetTooltip($disabledUsercheck, "Check all disabled Users to see if they have devices in Jabber.")
$disabledUsercheck.Add_MouseEnter({$tooltipdisableduser})
$disabledUsercheck.Enabled = $false
$form.Controls.Add($disabledUsercheck)

$calendarlabel = New-Object System.Windows.Forms.Label
$calendarlabel.Location = New-Object System.Drawing.Point(10,58)
$calendarlabel.Size = New-Object System.Drawing.Size(250,22)
$calendarlabel.Font = New-Object System.Drawing.Font("Segoe UI",12,[System.Drawing.FontStyle]::Regular)
$calendarlabel.Text = 'Users Last Logged in before:'
$form.Controls.Add($calendarlabel)

$calendar = New-Object Windows.Forms.DateTimePicker
$calendar.Location = New-Object System.Drawing.Point(10,80)
$calendar.Size = New-Object System.Drawing.Size(310, 40)
$calendar.Font = New-Object System.Drawing.Font("Sugoe UI",12,[System.Drawing.FontStyle]::Regular)
$calendar.Value = [DateTime]((Get-Date).AddDays(-90))
$calendar.MaxDate = [DateTime](Get-Date)
$calendar.DropDownAlign = 'Right'
$calendar.Enabled = $false
$form.Controls.Add($calendar)

$consolebox = New-Object System.Windows.Forms.RichTextbox
$consolebox.Location = New-Object System.Drawing.Point(325,30)
$consolebox.Size = New-Object System.Drawing.Size(595,260)
$consolebox.Multiline = $true
$consolebox.BorderStyle = 'Fixed3d'
$consolebox.Font = New-Object System.Drawing.Font("Lucida Console",12,[System.Drawing.FontStyle]::Regular)
$consolebox.BackColor = 'Black'
$consolebox.ReadOnly = $true
$consolebox.HideSelection = $false
$consolebox.Cursor = 'Arrow'
$form.Controls.Add($consolebox)

$findusersbutton = New-Object System.Windows.Forms.Button
$findusersbutton.Location = New-Object System.Drawing.Point(10,225)
$findusersbutton.Size = New-Object System.Drawing.Size(153,40)
$findusersbutton.Font = New-Object System.Drawing.Font("Sugoe UI",12,[System.Drawing.FontStyle]::Regular)
$findusersbutton.Text = 'Find Users'
$findusersbutton.Enabled = $false
$findusersbutton.Add_Click({ Find-Expired-Users })
$form.Controls.Add($findusersbutton)

$removedevicesbutton = New-Object System.Windows.Forms.Button
$removedevicesbutton.Location = New-Object System.Drawing.Point(167,225)
$removedevicesbutton.Size = New-Object System.Drawing.Size(153,40)
$removedevicesbutton.Font = New-Object System.Drawing.Font("Sugoe UI",12,[System.Drawing.FontStyle]::Regular)
$removedevicesbutton.Text = 'Remove Devices'
$removedevicesbutton.Enabled = $false
$removedevicesbutton.Add_Click({ CUCM-Device-Removal })
$form.Controls.Add($removedevicesbutton)

$userlistlabel = New-Object System.Windows.Forms.Label
$userlistlabel.Location = New-Object System.Drawing.Point(10,270)
$userlistlabel.Size = New-Object System.Drawing.Size(800,22)
$userlistlabel.Font = New-Object System.Drawing.Font("Sugoe UI",11,[System.Drawing.FontStyle]::Regular)
$userlistlabel.Text = 'Select users to remove Jabber Devices from:'
$form.Controls.Add($userlistlabel)

$userlistbox = New-Object System.Windows.Forms.listView
$userlistbox.View = 'Details'
$columns = [psCustomObject] @{ Name='Name'; width = 140 },
           [psCustomObject] @{ Name='samAccountName'; width = 110 },
           [psCustomObject] @{ Name='Enabled'; width = -2 },
           [psCustomObject] @{ Name='LastLogonDate'; width = 130 },
           [psCustomObject] @{ Name='CreatedOn'; width = 130 },
           [psCustomObject] @{ Name='Description'; width = 220 },
           [psCustomObject] @{ Name='RegisteredDevices'; width = 90 }
ForEach ($column in $columns) { [void]$userlistbox.Columns.Add($column.Name, $column.width) }
$userlistbox.FullRowSelect = $true
$userlistbox.CheckBoxes = $true
$userlistbox.Location = New-Object System.Drawing.Point(10,295)
$userlistbox.size = New-Object System.Drawing.Size(905,20)
$userlistbox.Font = New-Object System.Drawing.Font("Arial",10,[System.Drawing.FontStyle]::Regular)
$userlistbox.BackColor = 'LightGray'
$form.Controls.Add($userlistbox)
$userlistbox.Height = 300
$form.Add_Shown({$userlistbox.Select()})

Try { 
    Import-Module ActiveDirectory -ErrorAction Stop
    ## Get and display the Disabled User OU
    $disabledOU = Get-ADOrganizationalUnit -filter * -Properties Name,DistinguishedName | Select-Object Name,DistinguishedName | where { $_.Name -like "*Disabled*" -and $_.Name -like "*Users*" }
    $disabledOUselect.Items.Add($disabledOU.DistinguishedName)
    Write-Console White "-----------------Jabber Device Manager-----------------`r`nPlease login to the CUCM server to manage user devices."
} Catch { 
    Write-Console Red "Unable to load Active Directory Powershell module, please run this script from a machine that has it installed..."
    $loginbutton.Enabled = $false
    $servertextbox.Enabled = $false
}
$result = $form.ShowDialog()

$stream.Dispose()
$form.Dispose()