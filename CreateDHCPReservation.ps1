####################################################################################
# This code will minimize the Powershell console window.
if (-not $showWindowAsync) { $showWindowAsync = Add-Type –memberDefinition '[DllImport("user32.dll")]public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);' -name “Win32ShowWindowAsync” -namespace Win32Functions –passThru }; function Show-PowerShell() { $null = $showWindowAsync::ShowWindowAsync((Get-Process –id $pid).MainWindowHandle, 10) }; function Hide-PowerShell() { $null = $showWindowAsync::ShowWindowAsync((Get-Process –id $pid).MainWindowHandle, 2) }; Hide-PowerShell
####################################################################################


[void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") | Out-Null
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Collections") | Out-Null
[void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic") | Out-Null

[string[]]$global:fullservers = @("USHOUADDR01", "USDALDSIP01", "GBLONADDR01", "HKHKGADDR01", "AUMELCTX01", "INHYDADDR01") #, "GBHENADDR01", "HKHKGBCPADDR01"
[string[]]$global:servercomments = @("NA Production", "NA BR", "Europe Prod (BR is auto-updated)", "Asia Prod (BR is auto-updated)", "Australia Prod (BR is auto-updated)", "Hyderabad")
[string[]]$global:allservers = @("USHOUADDR01", "USDALDSIP01")
[object[]]$global:foundscopes = @()
[boolean]$global:usesingleform = $false
[boolean]$global:singleformadd = $true

$global:selectedservers = @()
$global:checkedservers = @()

[int32]$age_scopequery = 60
[int32]$age_rangequery = 60
[int32]$age_exclusionquery = 60
[int32]$age_reservationquery = 60
[int32]$age_leasequery = 60

[string]$scriptfolder = "D:\scripts\CreateDHCPReservation"
[string]$tempfolder = "$scriptfolder\DHCPTempInfo"
[string]$logfolder = "$scriptfolder\DHCPLogFiles"
[int32]$numinfocolumns = 3 

[string]$ButtonText_DoIt = "DO EET! DO EET NOOOW!!!"
[string]$ButtonText_Preview = "Preview"
[string]$ButtonText_NoActions = "No actions to perform"
[string]$ButtonText_CheckActions = "Check actions above"
[string]$error_invalidIP = "Invalid IP address"
[string]$error_invalidMAC = "Invalid MAC address"
[string]$error_serverproblem = "Cannot access the server"
[string]$error_scopenotfound = "Scope does not exist"
[string]$error_outofpool = "Not in the address pool"
[string]$error_excluded = "Excluded by rule"
[string]$error_ipreserved = "IP reserved with different MAC"
[string]$error_ipreserved2 = "Reservation found"
[string]$error_macreserved = "MAC reserved with different IP"
[string]$success_alreadyreserved = "This IP/MAC already reserved"
[string]$success_available = "Available"
[string]$success_available2 = "No reservation"
[string]$providernotfound = "The following command was not found: dhcp"
[string]$commandsuccessful = "Command completed successfully."
[string]$accessdeniederror = "Access is denied."
[string]$requireselevationerror = "The requested operation requires elevation."
[string]$serverversionerror = "Unable to determine the DHCP Server version"
[string]$validscopeerror = "The command needs a valid Scope IP Address."

[string]$checklabel_delete = "Enter IP addresses for deletion only"
[string]$checklabel_getleaseinfo = "Use existing lease info"
[string]$instructionsnormal = "Copy the reservation info to the clipboard with one entry per line:`nIP Address <tab> MAC Address <tab> Name (optional)`nThen click `"Get Info`"."
[string]$instructionsdelete = "Copy the list of IP addresses to the clipboard with one entry `nper line, then click `"Get Info`"."
[string]$instructionsgetleaseinfo = "Copy the list of device names to the clipboard with one entry `nper line, then click `"Get Info`". The existing lease info will be retrieved from the server."

[boolean]$fix_notinscope = $false
[boolean]$fix_exclusions = $false
[boolean]$fix_reservations = $false
[boolean]$fix_foundavailable = $false

[object[]]$allactions = @()

[string]$errorinfo = ""

[int]$global:sortcolumn = -1

$showScopeFormOnly = $false


add-type @" 
public struct Query_Result { 
   public string IP; 
   public string MAC; 
   public string Name;
   public object[] servers; 
   public int JobID;
} 
"@ 

add-type @" 
public struct Server_Info3 { 
    public string Name; 
    public bool ServerAccessError;
    public bool Success; 
    public string Scope; 
    public string ScopeActive; 
    public string Mask; 
    public string Range;
    public bool NotInScope;
    public string Exclusion;
    public string Reservation;
    public bool ReservationMatchExists;
    public string Error;
    public string TempFileScopes;
    public string TempFileRanges;
    public string TempFileExclusions;
    public string TempFileReservations;
    public string TempFileLeases;
} 
"@ 

add-type @" 
public struct Action {
    public string IP;
    public string MAC;
    public bool DeleteExclusion;
    public bool CreateExclusion;
    public bool DeleteReservation;
    public bool CreateReservation;
    public string Description;
    public string Command;
    public string Result;
    public string TempFile;
}
"@ 

$comparerClassString = @"
   using System;
   using System.Windows.Forms;
   using System.Drawing;
   using System.Collections;
   public class ListViewItemComparer : IComparer
   {
     private int col;
     private SortOrder order;
     public ListViewItemComparer()
     {
       col = 0;
       order = SortOrder.Ascending;
     }
     public ListViewItemComparer(int column, SortOrder order)
     {
       col = column;
       this.order = order;
     }
     public int Compare(object x, object y)
     {
       int returnVal= -1;
       returnVal = String.Compare(((ListViewItem)x).SubItems[col].Text,((ListViewItem)y).SubItems[col].Text);
       if (order == SortOrder.Descending) returnVal *= -1;
       return returnVal;
     }
   }
"@
Add-Type -TypeDefinition $comparerClassString -ReferencedAssemblies ('System.Windows.Forms', 'System.Drawing')


function run_queries {
    disable_form
    if ($usesingleform) {
        $inputvalue = "$($Txt2_IP.Text)`t$($Txt2_MAC.Text)"
        if ($Txt2_Name.Text -ne "") { $inputvalue += "`t$($Txt2_Name.Text)" }
        $inputvalues = @($inputvalue)
        $usesingleform = $false
    } else {
        $inputvalues = (& {powershell –sta {add-type –a system.windows.forms; [windows.forms.clipboard]::GetText()}}).Replace("`r`n", "`n").Split("`n")
        #$inputvalues = @("10.194.234.1`t0004c1c00d96", "10.194.234.16`t00-04-c1-c0-0d-97`tReservation 1", "10.194.234.17`t00 30 c1 c0 0d 98`tReservation 2", "10.194.242.49`t00:50:56:00:00:1f", "10.194.235.19`t00-04-c1-c0-0d-95", "10.191.50.124`t00:00:aa:d3:23:8e")
        #$inputvalues = @("10.194.235.19`t00-30-c1-c0-0d-98")
        #$inputvalues = @("10.191.50.124`t00:00:aa:d3:23:8e")
        #$inputvalues = @("10.194.242.49`t00:50:56:00:00:1f")
        #$inputvalues = @("10.194.235.49`t00-30-c1-c0-0d-96`tReservation 1", "10.194.235.50`t00-30-c1-c0-0d-97`tReservation 2")
        #$inputvalues = @("10.194.234.105`t00-30-c1-c0-0d-98", "10.194.242.49`t00-50-56-00-00-1f")
    }
    
    if ($Chk_RefreshTempData.Checked) {
        $tempfiles = @( Get-ChildItem "$tempfolder\*.*" -ErrorAction SilentlyContinue )  #| ? {($_.Name -match "Exclusions" -or $_.Name -match "Reservations")} 
        $tempfiles | ForEach-Object { Remove-Item $_ }
        $Chk_RefreshTempData.Checked = $false
    }
    $Txt_Results.Text = ""; $Rdo_ServerInfo.Checked = $true; $Form_DHCP.Refresh() 
    [string[]]$allips = @()
    [string[]]$allmacs = @()
    $allinputgood = $true
    $badIPlist = @()
    $badMAClist = @()
    $badNamelist = @()
    $badNameResolvelist = @()
    $dupelist = @()
    $duplicateIPs = @()
    $duplicateMACs = @()
    $incompletelist = @()
    for ($i=0; $i -lt $inputvalues.Count; $i++) {
        [string]$thisline = $inputvalues[$i]
        if ($thisline -eq "") { continue }
        if (-not $Chk_GetExistingLeases.checked) {
            [string]$ipaddress = $thisline.Trim().Split("`t")[0].Trim()
            [string]$check_ipaddress = validate_IP $ipaddress
            if ($check_ipaddress -eq "") { $badIPlist += $ipaddress ;$allinputgood = $false; continue }
            if ([array]::IndexOf($allips, $ipaddress) -gt -1) { $duplicateIPs += $ipaddress; $allinputgood = $false; continue } else { $allips += $ipaddress }
            if (-not $Chk_EnterIPsForDelete.Checked) {
                [string]$macaddress = ""
                if ($thisline.Split("`t").Length -eq 2 -or $thisline.Split("`t").Length -eq 3) {
                    $macaddress = clean_mac $thisline.Trim().Split("`t")[1].Trim()
                    $check_macaddress = validate_MAC $macaddress
                    if ($check_macaddress -eq "") { $badMAClist += $macaddress ;$allinputgood = $false; continue }
                    if ([array]::IndexOf($allmacs, $macaddress) -gt -1) { $duplicateMACs += $macaddress ;$allinputgood = $false; continue } else { $allmacs += $macaddress }
                } else {
                    $incompletelist += $thisline; $allinputgood = $false; continue
                }
            }
        } else {
            [string]$name = $thisline.Trim()
            if ($name.Split("`t").Count -gt 1) { $badNamelist += $name; $allinputgood = $false; continue }
            if ($name) {
                $ipaddress = Resolve-NameToIP $name
                if (validate_IP $name) { $badNamelist += $name; $allinputgood = $false; continue }
                if (validate_MAC $name) { $badNamelist += $name; $allinputgood = $false; continue }
                if (-not (validate_IP $ipaddress)) { $badNameResolvelist += $name; $allinputgood = $false; continue }
            }
        }
    }
    if (-not $allinputgood) {
        if ($badIPlist.Count -gt 0) { [Windows.Forms.Messagebox]::Show("Invalid IP addresses found.`n`n$(Convert-ArrayToStringList $badIPlist)"); enable_form; return }
        if ($badMAClist.Count -gt 0) { [Windows.Forms.Messagebox]::Show("Invalid MAC addresses found.`n`n$(Convert-ArrayToStringList $badMAClist)"); enable_form; return }
        if ($badNamelist.Count -gt 0) { [Windows.Forms.Messagebox]::Show("Invalid names found. Each line should contain one name only (no tabs). Names cannot be IP addresses or MAC addresses.`n`n$(Convert-ArrayToStringList $badNamelist)"); enable_form; return }
        if ($badNameResolvelist.Count -gt 0) { [Windows.Forms.Messagebox]::Show("Some names could not be resolved to IP addresses.`n`n$(Convert-ArrayToStringList $badNameResolvelist)"); enable_form; return }
        if ($incompletelist.Count -gt 0) { [Windows.Forms.Messagebox]::Show("Incomplete items found.  Each line must have both an IP address and a MAC address separated by a tab.`n`n$(Convert-ArrayToStringList $incompletelist)`n`nIf you want to enter a list of IP addresses for deletion only, check the checkbox `nat the top-right of the window."); enable_form; return }
        if ($duplicateIPs.Count -gt 0) { [Windows.Forms.Messagebox]::Show("Duplicate IP addresses found.  All IP addresses must be unique.`n`n$(Convert-ArrayToStringList $duplicateIPs)"); enable_form; return }
        if ($duplicateMACs.Count -gt 0) {
            [Windows.Forms.DialogResult]$response = [Windows.Forms.Messagebox]::Show("Duplicate MAC addresses found.  Multiple reservations with the same MAC address in a single scope are not allowed.  Are you sure you want to continue?`n`n$(Convert-ArrayToStringList $duplicateMACs)", "", [Windows.Forms.MessageboxButtons]::OKCancel)
            if ($response -eq [Windows.Forms.DialogResult]::Cancel) { enable_form; return }
        }
    }

    $LstVw_WorkItems.Items.Clear()
    for ($i=0; $i -lt $inputvalues.Count; $i++) {
        if (-not $Chk_GetExistingLeases.checked) {
            [string]$thisline = $inputvalues[$i].Trim()
            if ($thisline -eq "") { continue }
            [string]$ipaddress = validate_IP $thisline.Trim().Split("`t")[0].Trim()
            [string]$macaddress = ""
            [string]$name = ""
            if (-not $Chk_EnterIPsForDelete.Checked) {
                $macaddress = validate_MAC $thisline.Trim().Split("`t")[1].Trim()
                $macaddress = [System.Text.RegularExpressions.Regex]::Replace($macaddress,"[^a-fA-F0-9]","");
                if ($thisline.Split("`t").Length -eq 3) { $name = $thisline.Split("`t")[2].Trim() }
            }
            [Windows.Forms.ListViewItem]$newitem = New-Object Windows.Forms.ListViewItem($ipaddress)
            $newitem.SubItems.Add($macaddress)
            $newitem.SubItems.Add($name)
            for ($add=0; $add -le $allservers.Count; $add++) { $newitem.SubItems.Add("") }
            $LstVw_WorkItems.Items.Add($newitem)
        } else {
            $name = $inputvalues[$i].Trim()
            if ($name) {
                $ipaddress = Resolve-NameToIP $name
                [Windows.Forms.ListViewItem]$newitem = New-Object Windows.Forms.ListViewItem($ipaddress)
                $newitem.SubItems.Add("")
                $newitem.SubItems.Add($name)
                for ($add=0; $add -le $allservers.Count; $add++) { $newitem.SubItems.Add("") }
                $LstVw_WorkItems.Items.Add($newitem)
            }
        }
    }

    $LstVw_WorkItems.Refresh()
    reset_actions
    for ($i=0; $i -lt $inputvalues.Count; $i++) {
        [string]$thisline = $inputvalues[$i].Trim()
        if ($thisline -eq "") { continue }
        [string]$ipaddress = ""
        [string]$macaddress = ""
        [string]$name = ""
        if (-not $Chk_GetExistingLeases.Checked) {
            $ipaddress = validate_IP $thisline.Split("`t")[0].Trim()
            if (-not $Chk_EnterIPsForDelete.Checked) {
                $macaddress = validate_MAC $thisline.Split("`t")[1].Trim()
                if ($thisline.Split("`t").Length -eq 3) { $name = $thisline.Split("`t")[2].Trim() }
            }
        } else {
            $name = $thisline.Split("`t")[0].Trim()
            $ipaddress = validate_IP (Resolve-NameToIP $name)
            $macaddress = ""
        }
        [Query_Result]$result = New-Object Query_Result
        $result.IP = $ipaddress
        $result.MAC = $macaddress
        $result.Name = $name
        [Server_Info3]$actions = New-Object Server_Info3
        for ($d=0; $d -lt $allservers.Count; $d++) {
            [Server_Info3]$server = New-Object Server_Info3
            $server.Success = $true
            $server.Name = $allservers[$d]
            $server.NotInScope = $false
            $server.Exclusion = ""
            $server.Reservation = ""
            $server.ReservationMatchExists = $false
            $result.Servers += $server
            $result = GetScope $result $d
            if (-not $result.Servers[$d].Success) { $fix_notinscope = $true; update_status $result $d ""; continue }
            if (-not $Chk_EnterIPsForDelete.Checked) {
                $result = validate_Range $result $d
                if (-not $result.Servers[$d].Success) { $fix_notinscope = $true; update_status $result $d ""; continue }
                $result = validate_Exclusions $result $d
                if ($result.Servers[$d].Exclusion -ne "") { $fix_exclusions = $true }
            }
            if ($Chk_GetExistingLeases.Checked) {
                $result = Get-MACAddress $result $d
                $macaddress = $result.MAC
                update_status $result $d $result.MAC 1
            }
            $result = validate_Reservations $result $d
            if ($result.Servers[$d].Reservation -ne "") {
                update_status $result $d ""
                if (-not $result.Servers[$d].Success) { $fix_reservations = $true }
            }
            if ($result.Servers[$d].Exclusion -ne "") { update_status $result $d ""; continue }
            if ($result.Servers[$d].Reservation -ne "") { continue }
            $result.Servers[$d].Error = $success_available
            if ($Chk_EnterIPsForDelete.Checked) { $result.Servers[$d].Error = $success_available2 }
            update_status $result $d ""
            $fix_foundavailable = $true
        }
    }
    if ($fix_notinscope -or $fix_exclusions -or $fix_reservations -or $fix_foundavailable) {
        [boolean]$fix_canprocess = ($fix_exclusions -or $fix_reservations -or $fix_foundavailable)
        if ($fix_canprocess) {
            $Grp_Actions.Visible = $true
            $Chk_ShowNetShCmds.Visible = $true
            if ($fix_notinscope) { $Chk_SkipItemsNotInScope.Visible = $true }
            if ($fix_reservations) { $Chk_DelExistReservations.Visible = $true }
            if (-not $Chk_EnterIPsForDelete.Checked) {
                if ($fix_exclusions) { $Chk_RmvConfExclusions.Visible = $true }
                $Chk_DelExistReservations.Text = "Delete existing reservations"
                $Chk_CreateReservations.Visible = $true
            } else {
                $Chk_DelExistReservations.Text = "Delete reservations"
            }
        }
    }
    check_actions
    enable_form
}

function Convert-ArrayToStringList {
param($list)
    [string]::Join("`r", $list).Replace("`r", "`r`n")
}

function GetScope {
param([Query_Result]$result, [int32]$d)
    $dhcpserver = $result.Servers[$d].Name
    $ipaddress = $result.IP
    update_status $result $d "Retrieving scopes..."
    $tmpOutputFile = "$tempfolder\scopes-$dhcpserver.txt"
    $result.Servers[$d].TempfileScopes = $tmpOutputFile
    if (Test-Path $tmpOutputFile) {
        [string[]]$scopes = Get-Content $tmpOutputFile
        if ($scopes[0] -match $providernotfound) { Remove-Item $tmpOutputFile }
    }
    update_tempfile $tmpOutputFile $age_scopequery "show scope"
    [string[]]$scopes = Get-Content $tmpOutputFile
    $result.Servers[$d].ServerAccessError = $false
    if ($scopes | ? {$_ -match $providernotfound}) {
        $Txt_Results.Text = "DHCP provider not available on this computer.  You must run this script on Windows Vista, Windows 7 or Windows Server 2008."
        $result.Servers[$d].Success = $false
        $result.Servers[$d].ServerAccessError = $true
        $result.Servers[$d].NotInScope = $true
    } elseif ($scopes | ? {($_ -match $accessdeniederror) -or ($_ -match $requireselevationerror)}) {
        $Txt_Results.Text = "Access denied error - insufficient permissions to query the DHCP server."
        $result.Servers[$d].Success = $false
        $result.Servers[$d].ServerAccessError = $true
        $result.Servers[$d].NotInScope = $true
    } elseif ($scopes | ? {$_ -match $serverversionerror}) {
        $Txt_Results.Text = "DHCP server error.`r`n" + [string]::Join("`n", $scopes).Replace("`n", "`r`n")
        $result.Servers[$d].Success = $false
        $result.Servers[$d].ServerAccessError = $true
        $result.Servers[$d].NotInScope = $true
    } else {
        update_status $result $d "Locating scope..."
        $ipText = pad_IP $ipaddress
        $result.Servers[$d].Scope = ""
        [boolean]$scopematched = $false
        foreach ($thisscope in $foundscopes) {
            if ($thisscope.Name -eq $result.Servers[$d].Name) {
                $scopeID = $thisscope.Scope
                $scopeactive = $thisscope.ScopeActive
                $mask = $thisscope.Mask
                $startIP = $scopeID
                $endIP = ipadd $scopeID (((hex2dec (ip2hex $mask)) * -1) - 1)
                $startIPText = pad_IP $startIP
                $endIPText = pad_IP $endIP
                if (($iptext -ge $startIPText) -and ($iptext -le $endIPText)) {
                    $result.Servers[$d].Scope = $scopeID
                    $result.Servers[$d].ScopeActive = $scopeactive
                    $result.Servers[$d].Mask = $mask
                    $scopematched = $true
                    $startIP = $thisscope.Range.Split("-")[0].Trim()
                    $endIP = $thisscope.Range.Split("-")[1].Trim()
                    $startIPText = pad_IP $startIP
                    $endIPText = pad_IP $endIP
                    if (($iptext -ge $startIPText) -and ($iptext -le $endIPText)) { $result.Servers[$d].Range = $thisscope.Range }
                    break
                }
            }
        }
        if (-not $scopematched) {
            for ($i=0; $i -lt $scopes.Length; $i++) {
                if ($scopes[$i].Length -gt 15) {
                    $scopeID = $scopes[$i].substring(1, 15).trim()
                    if ((validate_IP $scopeID) -ne "") {
                        $mask = $scopes[$i].substring(18, 15).trim()
                        $startIP = $scopeID
                        $endIP = ipadd $scopeID (((hex2dec (ip2hex $mask)) * -1) - 1)
                        $startIPText = pad_IP $startIP
                        $endIPText = pad_IP $endIP
                        $scopestatus = $scopes[$i].substring(34, 14).trim()
                        if (($iptext -ge $startIPText) -and ($iptext -le $endIPText)) {
                            $result.Servers[$d].Scope = $scopeID
                            $result.Servers[$d].Mask = $mask
                            $result.Servers[$d].ScopeActive = $false
                            if ($scopestatus -eq "Active") { $result.Servers[$d].ScopeActive = $true }
                            $scopematched = $true
                            break
                        }
                    }
                }
            }
        }
        if (-not $scopematched) {
            $result.Servers[$d].Success = $false
            $result.Servers[$d].NotInScope = $true
        }
    }
    return $result
}

function validate_Range {
param([Query_Result]$result, [int32]$d)
    update_status $result $d "Checking address pool..."
    $dhcpserver = $result.Servers[$d].Name
    $ipaddress = $result.IP
    $ipscope = $result.Servers[$d].Scope
    $mask = $result.Servers[$d].Mask
    $tmpOutputFile = "$tempfolder\scope-$dhcpserver-$($ipscope)-Ranges.txt"
    $result.Servers[$d].TempfileRanges = $tmpOutputFile
    if ($result.Servers[$d].Range -ne $null) { return $result }
    update_tempfile $tmpOutputFile $age_rangequery "scope $ipscope show iprange"
    [string[]]$scopeRanges = Get-Content $tmpOutputFile
    if ($scopeRanges | ? {$_ -match $validscopeerror}) {
        $Txt_Results.Text = "Unable to query DCHP scope ranges.  This may be due to insufficient permissions."
        $result.Servers[$d].Success = $false
        $result.Servers[$d].NotInScope = $true
    } else {
        $result.Servers[$d].Error = ""
        $ipText = pad_IP $ipaddress
        [boolean]$scopematched = $false
        for ($i=0; $i -lt $scopeRanges.Length; $i++) {
            if ($scopeRanges[$i].Length -gt 38) {
                $startIP = $scopeRanges[$i].substring(3, 15).Trim()
                $endIP = $scopeRanges[$i].substring(23, 15).Trim()
                if (((validate_IP $startIP) -ne "") -and ((validate_IP $endIP) -ne "")) {
                    $scoperange += "$startIP - $endIP"
                    add_globalscope $result $d $scoperange
                    $startIPText = pad_IP $startIP
                    $endIPText = pad_IP $endIP
                    if (($iptext -ge $startIPText) -and ($iptext -le $endIPText)) {
                        $result.Servers[$d].Range = $scoperange
                        $scopematched = $true
                        break
                    }
                }
            }
        }
        if (-not $scopematched) {
            $result.Servers[$d].Success = $false
            $result.Servers[$d].NotInScope = $true
        }
    }
    return $result
}

function add_globalscope {
param([Query_Result]$result, [int32]$d, [string]$scoperange)
    $addscope = $true
    foreach ($thisscope in $foundscopes) {
        $thismatch = $true
        if ($thisscope.Name -ne $result.Servers[$d].Name) { $thismatch = $false }
        if ($thisscope.Scope -ne $result.Servers[$d].Scope) { $thismatch = $false }
        if ($thisscope.Mask -ne $result.Servers[$d].Mask) { $thismatch = $false }
        if ($thisscope.Range -ne $scoperange) { $thismatch = $false }
        if ($thismatch) { $addscope = $false }
    }
    if ($addscope) {
        $newscope = New-Object Server_Info3
        $newscope.Name = $result.Servers[$d].Name
        $newscope.Scope = $result.Servers[$d].Scope
        $newscope.ScopeActive = $result.Servers[$d].ScopeActive
        $newscope.Mask = $result.Servers[$d].Mask
        $newscope.Range = $scoperange
        $global:foundscopes += $newscope
    }
}

function validate_Exclusions {
param([Query_Result]$result, [int32]$d)
    $dhcpserver = $result.Servers[$d].Name
    $ipscope = $result.Servers[$d].Scope
    $ipaddress = $result.IP
    update_status $result $d "Checking exclusions..."
    $tmpOutputFile = "$tempfolder\scope-$dhcpserver-$($ipscope)-Exclusions.txt"
    $result.Servers[$d].TempfileExclusions = $tmpOutputFile
    $exclusionresult = Get-ExclusionRange $tmpOutputFile $ipscope $ipaddress $age_exclusionquery
    if ($exclusionresult -eq "error") {
        $Txt_Results.Text = "Unable to query DCHP exclusion info.  This may be due to insufficient permissions."
        $result.Servers[$d].Success = $false
    } else {
        $result.Servers[$d].Error = ""
        if ($exclusionresult) {
            $result.Servers[$d].Success = $false
            $result.Servers[$d].Exclusion = $exclusionresult
        }
    }
    return $result
}

function Get-ExclusionRange {
param($tmpOutputFile, $ipscope, $ipaddress, $fileage)
    update_tempfile $tmpOutputFile $fileage "scope $ipscope show excluderange" | Out-Null
    [string[]]$scopeExclusions = Get-Content $tmpOutputFile
    if ($scopeExclusions | ? {$_ -match $validscopeerror}) {
        "error"
    } else {
        $ipText = pad_IP $ipaddress
        for ($i=0; $i -lt $scopeExclusions.Length; $i++) {
            if ($scopeExclusions[$i].Length -gt 38) {
                $startIP = $scopeExclusions[$i].substring(3, 15).Trim()
                $endIP = $scopeExclusions[$i].substring(23, 15).Trim()
                if (((validate_IP $startIP) -ne "") -and ((validate_IP $endIP) -ne "")) {
                    $startIPText = pad_IP $startIP
                    $endIPText = pad_IP $endIP
                    if (($iptext -ge $startIPText) -and ($iptext -le $endIPText)) {
                        "$startIP - $endIP"
                        break
                    }
                }
            }
        }
    }
}

function Get-MACAddress {
param([Query_Result]$result, [int32]$d)
    $dhcpserver = $result.Servers[$d].Name
    $ipscope = $result.Servers[$d].Scope
    $ipaddress = $result.IP
    if ($result.Servers[$d].ScopeActive) {
        update_status $result $d "Getting MAC address..."
        $tmpOutputFile = "$tempfolder\scope-$dhcpserver-$($ipscope)-Leases.txt"
        $result.Servers[$d].TempfileLeases = $tmpOutputFile
        update_tempfile $tmpOutputFile $age_leasequery "scope $ipscope show clients"
        [string[]]$scopeLeases = Get-Content $tmpOutputFile
        if ($scopeReservations | ? {$_ -match $validscopeerror}) {
            $Txt_Results.Text = "Unable to query DCHP reservation info.  This may be due to insufficient permissions."
            $result.Servers[$d].Success = $false
        } else {
            $result.Servers[$d].Error = ""
            for ($i=0; $i -lt $scopeLeases.Length; $i++) {
                if ($scopeLeases[$i].Length -gt 52) {
                    $reservedIP = validate_IP ($scopeLeases[$i].substring(0, 15).Trim())
                    if (($reservedIP -ne "") -and ($ipaddress -eq $reservedIP)) {
                        $result.MAC = clean_mac ($scopeLeases[$i].substring(34, $scopeLeases[$i].indexOf(" -", 34) - 34).Trim())
                        break
                    }
                }
            }
        }
    }
    return $result
}

function validate_Reservations {
param([Query_Result]$result, [int32]$d)
    $dhcpserver = $result.Servers[$d].Name
    $ipscope = $result.Servers[$d].Scope
    $ipaddress = $result.IP
    $macaddress = $result.MAC
    update_status $result $d "Checking reservations..."
    $tmpOutputFile = "$tempfolder\scope-$dhcpserver-$($ipscope)-Reservations.txt"
    $result.Servers[$d].TempfileReservations = $tmpOutputFile
    update_tempfile $tmpOutputFile $age_reservationquery "scope $ipscope show reservedip"
    [string[]]$scopeReservations = Get-Content $tmpOutputFile
    if ($scopeReservations | ? {$_ -match $validscopeerror}) {
        $Txt_Results.Text = "Unable to query DCHP reservation info.  This may be due to insufficient permissions."
        $result.Servers[$d].Success = $false
    } else {
        $result.Servers[$d].Error = ""
        for ($i=0; $i -lt $scopeReservations.Length; $i++) {
            if ($scopeReservations[$i].Length -gt 38) {
                $reservedIP = validate_IP ($scopeReservations[$i].substring(3, 15).Trim())
                if (($reservedIP -ne "") -and ($ipaddress -eq $reservedIP)) {
                    $result.Servers[$d].Success = $false
                    $reservedMACraw = $scopeReservations[$i].substring(24).Trim()
                    $reservedMAC = ""
                    if ($reservedMACraw.length -eq 18) { $reservedMAC = validate_MAC $reservedMACraw }
                    if (-not $reservedMAC) { $reservedMAC = "$reservedMACraw (invalid)" }
                    $result.Servers[$d].Reservation = "$reservedIP $reservedMAC"
                    if ($reservedMAC -eq $macaddress) {
                        $result.Servers[$d].Error = $success_alreadyreserved
                        $result.Servers[$d].ReservationMatchExists = $true
                    } else {
                        $result.Servers[$d].Error = $error_ipreserved
                    }
                    break
                }
            }
        }
        if ($result.Servers[$d].Success -and (-not $result.Servers[$d].ReservationMatchExists) -and (-not $Chk_EnterIPsForDelete.Checked)) {
            for ($i=0; $i -lt $scopeReservations.Length; $i++) {
                if ($scopeReservations[$i].Length -gt 38) {
                    $reservedMAC = validate_MAC ($scopeReservations[$i].substring(24, 20).Trim())
                    if (($reservedMAC -ne "") -and ($macaddress -eq $reservedMAC)) {
                        $result.Servers[$d].Success = $false
                        $reservedIP = validate_IP ($scopeReservations[$i].substring(3, 15).Trim())
                        $result.Servers[$d].Reservation = "$reservedIP $reservedMAC"
                        if ($reservedIP -ne $ipaddress) { $result.Servers[$d].Error = $error_macreserved }
                        break
                    }
                }
            }
        }
    }
    return $result
}

function update_status {
param([Query_Result]$result, [int32]$d, [string]$status, [int32]$updatecolumn = -1)
    [string]$dhcpserver = $result.Servers[$d].Name
    [string]$ipaddress = $result.IP
    [string]$name = ""
    if ($result.name -ne "") { $name = $result.name }
    if ($status -eq "") { $status = get_error $result $d "" }
    [int32]$updateIndex = -1
    for ($j=0; $j -lt $LstVw_WorkItems.Items.Count; $j++) { if ($LstVw_WorkItems.Items[$j].Text -eq $ipaddress -or ($name -ne "" -and $LstVw_WorkItems.Items[$j].Subitems[2].Text -eq $name)) { $updateIndex = $j; break } }
    if ($updatecolumn -eq -1) { $updatecolumn = $numinfocolumns + [array]::IndexOf($allservers, $dhcpserver) }
    $LstVw_WorkItems.Items[$updateIndex].SubItems[$updatecolumn].Text = $status
    $LstVw_WorkItems.Items[$updateIndex].Tag = $result
    $LstVw_WorkItems.Refresh()
    $Form_DHCP.Refresh()
    $LstVw_WorkItems.EnsureVisible($updateIndex)
}

function get_error {
param([Query_Result]$result, [int32]$d, [string]$errortype)
    [string[]]$status = @()
    if ($result.Servers[$d].ServerAccessError -and $errortype -eq "") {
        $status += $error_serverproblem
    } else {
        if ($result.Servers[$d].Scope -eq "" -and $errortype -eq "") { $status += $error_scopenotfound }
        if ($result.Servers[$d].NotInScope -and $result.Servers[$d].Scope -ne "" -and $errortype -eq "") { $status += $error_outofpool }
        if ($result.Servers[$d].Exclusion -ne "" -and ($errortype -eq "" -or $errortype -eq "exclusion")) { $status += $error_excluded }
        if ($result.Servers[$d].Reservation -ne "" -and ($errortype -eq "" -or $errortype -eq "reservation")) {
            [string]$thisstatus = $result.Servers[$d].Error
            if ($Chk_EnterIPsForDelete.Checked -and $thisstatus -eq $error_ipreserved) { $thisstatus = $error_ipreserved2 }
            $status += $thisstatus
        } 
        if ($result.Servers[$d].Success -and -not $result.Servers[$d].ReservationMatchExists) {
            [string]$thisstatus = $success_available
            if ($Chk_EnterIPsForDelete.Checked) { $thisstatus = $success_available2 }
            $status += $thisstatus
        }
    }
    return [string]::Join(";", $status).Replace(";", "; ")
}

function check_actions {
    [boolean]$enable_preview = $true
    if ($Chk_SkipItemsNotInScope.Visible -and -not $Chk_SkipItemsNotInScope.Checked) { $enable_preview = $false }
    #if ($Chk_RmvConfExclusions.Visible -and -not $Chk_RmvConfExclusions.Checked) { $enable_preview = $false }
    #if ($Chk_DelExistReservations.Visible -and -not $Chk_DelExistReservations.Checked) { $enable_preview = $false }
    if ($Chk_CreateReservations.Visible -and -not $Chk_CreateReservations.Checked -and -not ($Chk_SkipItemsNotInScope.Visible -or $Chk_RmvConfExclusions.Visible -or $Chk_DelExistReservations.Visible)) { $enable_preview = $false }
    $onebuttonchecked = $false
    if ($Chk_SkipItemsNotInScope.Visible -and $Chk_SkipItemsNotInScope.Checked) { $onebuttonchecked = $true }
    if ($Chk_RmvConfExclusions.Visible -and $Chk_RmvConfExclusions.Checked) { $onebuttonchecked = $true }
    if ($Chk_DelExistReservations.Visible -and $Chk_DelExistReservations.Checked) { $onebuttonchecked = $true }
    if ($Chk_CreateReservations.Visible -and $Chk_CreateReservations.Checked) { $onebuttonchecked = $true }
    if (-not $onebuttonchecked) { $enable_preview = $false }
    if (-not $Chk_RmvConfExclusions.Visible -and -not $Chk_DelExistReservations.Visible -and -not $Chk_CreateReservations.Visible) {
        $Btn_Preview.Text = $ButtonText_NoActions
        $Chk_SkipItemsNotInScope.Visible = $false
    } else {
        $Btn_Preview.Enabled = $enable_preview
        if ($enable_preview) { $Btn_Preview.Text = $ButtonText_Preview } else { $Btn_Preview.Text = $ButtonText_CheckActions }
    }
}

function pad_IP {
param([string]$ipaddress)
    $iptext = ""
    for ($i=0; $i -le 3; $i++) { $iptext += $ipaddress.split(".")[$i].ToString().PadLeft(3, "0") }
    return $iptext
}

function reset_to_preview {
    if ($Btn_Preview.Text -eq $ButtonText_DoIt) { $Btn_Preview.Text = $ButtonText_Preview }
}

function show_info {
    if ($LstVw_WorkItems.SelectedItems.Count -eq 0) { return }
    set_fixedfont($true)
    reset_to_preview
    [string]$infotext = ""
    [Windows.Forms.ListViewItem]$item = $LstVw_WorkItems.SelectedItems[0]
    [Query_Result]$result = $item.Tag
    if ($Rdo_ServerInfo.Checked) {
        $infotext += "IP Address:   $($result.IP)`r`n"
        $infotext += "MAC Address: $($result.MAC)`r`n"
        for ($d=0; $d -lt $allservers.Count; $d++) {
            $infotext += "`r`n"
            $infotext += "Results for $($allservers[$d])`r`n"
            if ($result.Servers[$d].Scope -ne "" -and $result.Servers[$d].Scope -ne $null) { $infotext += "Scope: $($result.Servers[$d].Scope)`r`n" }
            if (-not $result.Servers[$d].Success) {
                if ($result.Servers[$d].NotInScope) {
                    $infotext += "Error: " + (get_error $result $d "") + "`r`n"
                    if ($result.Servers[$d].Range -ne "" -and $result.Servers[$d].Range -ne $null) { $infotext += "Address range for the scope is: $($result.Servers[$d].Range)`r`n" }
                }
                if ($result.Servers[$d].Exclusion -ne "" -and $result.Servers[$d].Exclusion -ne $null) {
                    if (-not $Chk_EnterIPsForDelete.Checked) { $infotext += "Error: " + (get_error $result $d "exclusion") + "`r`n" }
                    $infotext += "Exclusion rule: $($result.Servers[$d].Exclusion)`r`n"
                }
                if ($result.Servers[$d].Reservation -ne "" -and $result.Servers[$d].Reservation -ne $null) {
                    if (-not $Chk_EnterIPsForDelete.Checked) { $infotext += "Error: " + (get_error $result $d "reservation") + "`r`n" }
                    $infotext += "Existing Reservation: $($result.Servers[$d].Reservation)`r`n"
                }
            } else {
                if (-not $Chk_EnterIPsForDelete.Checked) {
                    if ($result.Servers[$d].ReservationMatchExists) {
                        $infotext += "The IP address is already reserved with the specified MAC address.`r`n"
                    } else {
                        $infotext += "Address is available for reservation.`r`n"
                    }
                } else {
                    $infotext += "No reservation for this address.`r`n"
                }
            }
        }
        $Txt_Results.Text = $infotext
        return
    }
    if ($Rdo_AllScopes.Checked -or $Rdo_AddressPool.Checked -or $Rdo_Exclusions.Checked -or $Rdo_Reservations.Checked) {
        for ($d=0; $d -lt $allservers.Count; $d++) {
            if ($Rdo_AllScopes.Checked) { $tmpOutputFile = $result.Servers[$d].TempfileScopes }
            if ($Rdo_AddressPool.Checked) { $tmpOutputFile = $result.Servers[$d].TempfileRanges }
            if ($Rdo_Exclusions.Checked) { $tmpOutputFile = $result.Servers[$d].TempfileExclusions }
            if ($Rdo_Reservations.Checked) { $tmpOutputFile = $result.Servers[$d].TempfileReservations }
            if ($tmpOutputFile -ne $null) {
                $tmpOutputData = Get-Content $tmpOutputFile
                if (-not $Rdo_AllScopes.Checked) { $tmpOutputData = sort_data $tmpOutputData }
                $filecontent = [string]::Join("`n", $tmpOutputData).Replace("`n", "`r`n")
                if ($filecontent.substring(0,2) -eq "`r`n") { $filecontent = $filecontent.substring(2) }
                $infotext += $allservers[$d] + "`r`n" + $filecontent + "`r`n`r`n"
            } else {
                $infotext += $allservers[$d] + "$`r`n`r`nInfo not available because the query was skipped.`r`n`r`n"
            }
        }
        $Txt_Results.Text = $infotext
        return
    }
}

function sort_data {
param([string[]]$alltext)
    disable_form
    $header = @()
    $body = @()
    $footer = @()
    $section = 1
    $alltext | ForEach-Object {
        if ($section -eq 1) { if ($_ -match "\b\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}\b" -and -not ($_ -match "Changed the current scope")) { $section = 2 } else { $header += $_ } }
        if ($section -eq 2) { if ($_ -eq "") { $section = 3 } else { $body += $_ } }
        if ($section -eq 3) { $footer += $_ }
    }
    $body = sortbyIP $body
    $alltext = @()
    $alltext = $header + $body + $footer
    enable_form
    return $alltext
}

function sortbyIP {
param([string[]]$lines)
    [string[]]$sortlines = @()
    $lines | ForEach-Object {
        $regmatch = $_ -match "\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}"
        $ipaddress = $matches[0]
        $sortlines += ($_ -replace $ipaddress, (ip2hex $ipaddress))
    }
    $sortlines = ($sortlines | sort)
    [string[]]$newlines = @()
    $sortlines | ForEach-Object {
        $regmatch = $_ -match "\w{7,8}"
        $iphex = $matches[0]
        $newlines += ($_ -replace $iphex, (hex2ip $iphex))
    }    
    return $newlines
}

function enable_form {
    $Form_DHCP.Cursor = "Default"
    $Btn_GetInfo.Enabled = $true
    $Btn_Close.Enabled = $true
}

function disable_form {
    $Form_DHCP.Cursor = "WaitCursor"
    $Btn_GetInfo.Enabled = $false
    $Btn_Close.Enabled = $false
}

function reset_actions {
    $fix_notinscope = $false
    $fix_exclusions = $false
    $fix_reservations = $false
    $Grp_Actions.Visible = $false
    $Grp_Actions.Enabled = $true
    $Chk_ShowNetShCmds.Visible = $false
    $Chk_SkipItemsNotInScope.Visible = $false
    $Chk_RmvConfExclusions.Visible = $false
    $Chk_DelExistReservations.Visible = $false
    $Chk_CreateReservations.Visible = $false
    $Chk_SkipItemsNotInScope.Checked = $false
    $Chk_RmvConfExclusions.Checked = $false
    $Chk_DelExistReservations.Checked = $false
    $Chk_CreateReservations.Checked = $false
    $Btn_Preview.Enabled = $false
}

function validate_IP {
param([string]$ipaddress)
    $ipaddress = $ipaddress.Trim()
    if ($ipaddress.split(".").count -ne 4) { return "" }
    for ($i=0; $i -lt $ipaddress.length; $i++) {
        $char = [byte][char]$ipaddress.substring($i, 1)
        if ((($char -lt 48) -or ($char -gt 57)) -and ($char -ne 46)) { return "" }
    }
    for ($i=0; $i -le 3; $i++) {
        $ippart = [int]$ipaddress.split(".")[$i]
        if (($ippart -lt 0) -or ($ippart -gt 255)) { return "" }
    }
    return $ipaddress
}

function validate_MAC {
param([string]$macaddress)
    $macaddress = clean_mac $macaddress
    if ($macaddress.length -ne 12) { return "" }
    return $macaddress
}

function clean_mac {
param([string]$macaddress)
    $macaddress = [System.Text.RegularExpressions.Regex]::Replace($macaddress,"[^a-fA-F0-9]","");
    return $macaddress.ToLower()
}

function update_tempfile {
param([string]$tempfilename, [int32]$fileage, [string]$netshcmd)
    $tempfile = $false
    if (Test-Path $tempfilename) {
        $tempfile = @( Get-ChildItem $tempfilename -ErrorAction SilentlyContinue | ? {($_.LastWriteTime -gt ((Get-date).AddMinutes($fileage * -1)))} )
    }
    if (-not $tempfile) {
        $objStartInfo = New-Object System.Diagnostics.ProcessStartInfo
        $objStartInfo.FileName = "cmd.exe"
        $objStartInfo.windowStyle = "Hidden"
        $objStartInfo.arguments = "/C netsh.exe dhcp server \\$dhcpserver $netshcmd >`"$tempfilename`""
        [void][System.Diagnostics.Process]::Start($objStartInfo).WaitForExit()
        if ($tempfilename -match "scopes-" -or $tempfilename -match "-Ranges") { $global:foundscopes = @() }
    }
}

function build_actions {
param([int32]$actioncount = -1)
    [Action[]]$allactions = @()
    $startindex = 0
    $endindex = $LstVw_WorkItems.Items.Count - 1
    $exclusionrefresh = $age_exclusionquery
    $actionmaychangewarning = " (may change during processing)"
    if ($actioncount -ge 0) {
        $startindex = $actioncount
        $endindex = $actioncount
        $exclusionrefresh = -1
        $actionmaychangewarning = ""
    }
    for ($j=$startindex; $j -le $endindex; $j++) {
        [Query_Result]$result = $LstVw_WorkItems.Items[$j].Tag
        for ($d=0; $d -lt $allservers.Count; $d++) {
            [Server_Info3]$server = $result.Servers[$d]
            if ($server.NotInScope) { continue }
            [string]$dhcpserver = $server.Name
            [string]$ipscope = $server.Scope
            [string]$netshcmd = "netsh.exe dhcp server \\$dhcpserver scope $ipscope"
            [string]$ipaddress = $result.IP
            [string]$macaddress = $result.MAC
            [string]$name = $result.Name
            $tmpOutputFile = "$tempfolder\scope-$dhcpserver-$($ipscope)-Exclusions.txt"
            $server.Exclusion = Get-ExclusionRange $tmpOutputFile $ipscope $ipaddress $exclusionrefresh
            if ($Chk_RmvConfExclusions.Checked -and $server.Exclusion -ne "") {
                [string]$startIP = $server.Exclusion.Replace(" ", "").Split("-")[0]
                [string]$endIP = $server.Exclusion.Replace(" ", "").Split("-")[1]
                if ($startIP -eq $endIP) {
                    #just delete the exclusion
                    [Action]$newaction = new_action $result $d "DeleteExclusion"
                    $newaction.Description = "Delete exclusion range `"$($server.Exclusion)`" on $dhcpserver."
                    $newaction.Command = "$netshcmd delete excluderange $startIP $endIP"
                    $allactions += $newaction
                } else {
                    #determine how to split the exclusion, then delete the old one and create the new one(s)
                    [string[]]$newrange1 = Get-NewExclusionRange $ipaddress $startIP $endIP 1
                    [string[]]$newrange2 = Get-NewExclusionRange $ipaddress $startIP $endIP 2
                    [Action]$newaction = new_action $result $d "DeleteExclusion"
                    $newaction.Description = "Delete exclusion range `"$startIP - $endIP`" on $dhcpserver.$actionmaychangewarning"
                    $newaction.Command = "$netshcmd delete excluderange $startIP $endIP"
                    $allactions += $newaction
                    [Action]$newaction = new_action $result $d "CreateExclusion"
                    $newaction.Description = "Create new exclusion range `"$($newrange1[0]) - $($newrange1[1])`" on $dhcpserver.$actionmaychangewarning"
                    $newaction.Command = "$netshcmd add excluderange $($newrange1[0]) $($newrange1[1])"
                    $allactions += $newaction
                    if ($newrange2.Count -gt 0) {
                        [Action]$newaction = new_action $result $d "CreateExclusion"
                        $newaction.Description = "Create new exclusion range `"$($newrange2[0]) - $($newrange2[1])`" on $dhcpserver.$actionmaychangewarning"
                        $newaction.Command = "$netshcmd add excluderange $($newrange2[0]) $($newrange2[1])"
                        $allactions += $newaction
                    }
                }
            }
            if (-not ($Chk_DelExistReservations.Checked -and $Chk_CreateReservations.Checked -and $server.ReservationMatchExists)) {
                if ($Chk_DelExistReservations.Checked -and $server.Reservation) {
                    #delete the reservation
                    [string]$oldIP = $server.Reservation.Split(" ")[0]
                    [string]$oldMAC = $server.Reservation.Split(" ")[1]
                    [Action]$newaction = new_action $result $d "DeleteReservation"
                    $newaction.Description = "Delete reservation `"$oldIP - $oldMAC`" on $dhcpserver."
                    $newaction.Command = "$netshcmd delete reservedip $oldIP $oldMAC"
                    $allactions += $newaction
                }
                if ($Chk_CreateReservations.Checked) {
                    #create the new reservation
                    [Action]$newaction = new_action $result $d "CreateReservation"
                    $newaction.Description = "Create reservation `"$ipaddress - $macaddress`" on $dhcpserver."
                    if ($name -eq "") { $name = $macaddress }
                    $newaction.Command = "$netshcmd add reservedip $ipaddress $macaddress `"$name`""
                    $allactions += $newaction
                }
            }
        }
    }
    $allactions = Order-Actions $allactions
    return $allactions
}

function Order-Actions {
param($allactions)
    [Action[]]$act_delete_exclusion = @()
    [Action[]]$act_create_exclusion = @()
    [Action[]]$act_delete_reservation = @()
    [Action[]]$act_create_reservation = @()
    $allactions | ForEach-Object {
        if ($_.DeleteExclusion) { $act_delete_exclusion += $_ }
        if ($_.CreateExclusion) { $act_create_exclusion += $_ }
        if ($_.DeleteReservation) { $act_delete_reservation += $_ }
        if ($_.CreateReservation) { $act_create_reservation += $_ }
    }
    [Action[]]$newactions = $act_delete_exclusion + $act_create_exclusion + $act_delete_reservation + $act_create_reservation
    $newactions
}

function Get-NewExclusionRange {
param($ipaddress, $startIP, $endIP, $mode)
    [string[]]$newrange1 = @()
    [string[]]$newrange2 = @()
    if ((pad_IP $ipaddress) -eq (pad_IP $startIP)) {
        $newrange1 += @((ipadd $startIP 1), $endIP)
    } elseif ((pad_IP $ipaddress) -eq (pad_IP $endIP)) {
        $newrange1 += @($startIP, (ipadd $endIP -1))
    } elseif ((pad_IP $ipaddress) -gt (pad_IP $startIP) -and (pad_IP $ipaddress) -lt (pad_IP $endIP)) {
        $newrange1 += @($startIP, (ipadd $ipaddress -1))
        $newrange2 += @((ipadd $ipaddress 1), $endIP)
    }
    if ($mode -eq 1) { $newrange1 }
    if ($mode -eq 2) { $newrange2 }
}

function show_preview {
    set_fixedfont($false)
    $count_DeleteExclusion = 0
    $count_CreateExclusion = 0
    $count_DeleteReservation = 0
    $count_CreateReservation = 0
    $allactions = build_actions
    $allactions | ForEach-Object {
        if ($_.DeleteExclusion) { $count_DeleteExclusion ++ }
        if ($_.CreateExclusion) { $count_CreateExclusion ++ }
        if ($_.DeleteReservation) { $count_DeleteReservation ++ }
        if ($_.CreateReservation) { $count_CreateReservation ++ }
    }
    $countsummary = ""
    if ($count_DeleteExclusion -gt 0) { $addsummary = "delete $count_DeleteExclusion exclusion rule, "; if ($count_DeleteExclusion -gt 1) { $addsummary = $addsummary.Replace("rule", "rules") }; $countsummary += $addsummary }
    if ($count_CreateExclusion -gt 0) { $addsummary = "create $count_CreateExclusion exclusion rule, "; if ($count_CreateExclusion -gt 1) { $addsummary = $addsummary.Replace("rule", "rules") }; $countsummary += $addsummary }
    if ($count_DeleteReservation -gt 0) { $addsummary = "delete $count_DeleteReservation reservation, "; if ($count_DeleteReservation -gt 1) { $addsummary = $addsummary.Replace("tion", "tions") }; $countsummary += $addsummary }
    if ($count_CreateReservation -gt 0) { $addsummary = "create $count_CreateReservation reservation, "; if ($count_CreateReservation -gt 1) { $addsummary = $addsummary.Replace("tion", "tions") }; $countsummary += $addsummary }
    if ($countsummary -ne "") {
        $countsummary = "Will " + $countsummary.Substring(0, $countsummary.Length - 2) + "."
        [regex]$r = New-Object System.Text.RegularExpressions.Regex(",", [System.Text.RegularExpressions.RegexOptions]::RightToLeft)
        $countsummary = $r.Replace($countsummary, ", and", 1)
        [string[]]$pvwtext = @()
        $pvwtext += @("Preview of changes", $countsummary, "")
        [int32]$count = 0
        $allactions | ForEach-Object {
            $count++
            $pvwtext += "Step $($count.ToString()): $($_.Description)"
            if ($Chk_ShowNetShCmds.Checked) {
                $pvwtext += "   Command: $($_.Command)"
                $pvwtext += ""
            }
        }
        $Txt_Results.Text = [string]::Join("`n", $pvwtext).Replace("`n", "`r`n")
        $Btn_Preview.Text = $ButtonText_DoIt
    } else {
        $Txt_Results.Text = "No actions to perform."
        $Btn_Preview.Text = $ButtonText_NoActions
    }
}

function new_action {
param([Query_Result]$result, [string]$d, [string]$action)
    [Action]$newaction = New-Object Action
    $newaction.IP = $result.IP
    $newaction.MAC = $result.MAC
    if ($action -eq "DeleteExclusion") { $newaction.DeleteExclusion = $true }
    if ($action -eq "CreateExclusion") { $newaction.CreateExclusion = $true }
    if ($action -eq "DeleteReservation") { $newaction.DeleteReservation = $true }
    if ($action -eq "CreateReservation") { $newaction.CreateReservation = $true }
    if ($action -match "Exclusion") { $newaction.TempFile = $result.Servers[$d].TempFileExclusions }
    if ($action -match "Reservation") { $newaction.TempFile = $result.Servers[$d].TempFileReservations }
    return $newaction
}

function do_changes {
    $Grp_Actions.Enabled = $false
    $Chk_ShowNetShCmds.Enabled = $false
    $Txt_Results.Text = ""
    
    for ($j=0; $j -lt $LstVw_WorkItems.Items.Count; $j++) {
        [Action[]]$allactions = build_actions $j
        $numactions = 0
        if ($allactions) { $numactions = $allactions.count }
        for ($actionloop=0; $actionloop -lt $numactions; $actionloop++) {
            if ($allactions) { if ($allactions.gettype().name -eq "Action") { $thisaction = $allactions } else { $thisaction = $allactions[$actionloop] } }
            [string]$logfiledate = "{0:yyyyMMddhhmmss}" -f [DateTime](Get-Date)
            [string]$logfiledescr = $thisaction.Description.Replace("`"", "")
            if ($thisaction.DeleteExclusion -or $thisaction.CreateExclusion -or $thisaction.DeleteReservation -or $thisaction.CreateReservation) { $logfilename = "$logfolder\$logfiledate $($logfiledescr).txt" }
            $objStartInfo = New-Object System.Diagnostics.ProcessStartInfo
            $objStartInfo.FileName = "cmd.exe"
            $objStartInfo.WindowStyle = "Hidden"
            $objStartInfo.Arguments = "/C $($thisaction.Command) >`"$logfilename`""
            [void][System.Diagnostics.Process]::Start($objStartInfo).WaitForExit()
            $logfileinfo = Get-Content $logfilename
            $success = $false
            $logfileinfo | Foreach-Object { if ($_ -match $commandsuccessful) { $success = $true } }
            $commandoutput = @()
            if ($success) {
                $commandoutput += ("Success: " + $thisaction.Description)
            } else {
                $commandoutput += ("Failed: " + $thisaction.Description)
                $commandoutput += ($logfileinfo | ? {$_ -ne ""} )
            }
            $commandoutput += ""
            $Txt_Results.Text += [string]::Join("`n", $commandoutput).Replace("`n", "`r`n")
            $Txt_Results.Refresh()
            Scroll-ToEnd $Txt_Results
            if (Test-Path $thisaction.TempFile) { Remove-Item $thisaction.TempFile }
        }
    }
        
    $Grp_Actions.Visible = $false
    $Grp_Actions.Enabled = $true
    $Chk_ShowNetShCmds.Visible = $false
    $Chk_ShowNetShCmds.Enabled = $true
    $Txt_Results.Text += "`r`nProcessing completed."
    Scroll-ToEnd $Txt_Results
}

function Scroll-ToEnd {
param($control)
    $control.Select($control.Text.Length - 1, 0)
    $control.ScrollToCaret()
}

function change_servers {
    ask_servers
    set_listview_columns
}

function ask_servers {
    $handler_form2_Load = {
        $Form_DHCP.Enabled = $false
    }
    $handler_form2_FormClosed = {
        $Form_DHCP.Enabled = $true
    }
    $handler_form2_OKButton = {
        $newservers = @()
        foreach ($item in $listvw1.Items) {
            if ($item.Checked) { $newservers += $item.Text }
        }
        if ($newservers.Count -gt 0) {
            $global:allservers = $newservers
            $form2.Close()
        } else {
            [Windows.Forms.Messagebox]::Show("Please select at least one server.")
        }
    }
    [Windows.Forms.Form]$form2 = New-Object Windows.Forms.Form
    $form2.Text = "Select the target servers."
    $form2.Size = New-Object Drawing.Point 360,300
    $form2.MinimumSize = $form2.Size
    $form2.add_Load($handler_form2_Load)
    $form2.add_FormClosed($handler_form2_FormClosed)
    
    [Windows.Forms.ListView]$listvw1 = New-Object Windows.Forms.ListView
    $listvw1.Location = new-object Drawing.Point 10,25
    $listvw1.Size = new-object Drawing.Point ($form2.Width - 35), ($form2.Height - 110) 
    $listvw1.View = "Details"
    $listvw1.Anchor = "Top,Left,Right,Bottom"
    $listvw1.HideSelection = $false
    $listvw1.MultiSelect = $false
    $listvw1.Checkboxes = $true
    [void]$listvw1.Columns.Add("DHCP Server Name",120)
    [void]$listvw1.Columns.Add("Description",175)
    $listvw1.Items.Clear()
    for ($i=0; $i -lt $fullservers.Count; $i++) {
        [Windows.Forms.ListViewItem]$newitem = New-Object Windows.Forms.ListViewItem($fullservers[$i])
        $newitem.SubItems.Add($servercomments[$i])
        if ($allservers -contains $fullservers[$i]) { $newitem.Checked = $true }
        $listvw1.Items.Add($newitem)
    }

    [Windows.Forms.Button]$buttonCancel = New-Object Windows.Forms.Button
    $buttonCancel.Text = "Cancel"
    $buttonCancel.Location = New-Object Drawing.Point ($listvw1.Left + $listvw1.Width - $buttonCancel.Width),($listvw1.Top + $listvw1.Height + $boxspacing)
    $buttonCancel.Anchor = "Right,Bottom"

    [Windows.Forms.Button]$buttonOK = New-Object Windows.Forms.Button
    $buttonOK.Text = "OK"
    $buttonOK.add_click($handler_form2_OKButton)
    $buttonOK.Location = New-Object Drawing.Point ($buttonCancel.Left - $buttonOK.Width - $boxspacing),$buttonCancel.Top
    $buttonOK.Anchor = "Right,Bottom"

    $form2.Controls.Add($listvw1)
    $form2.Controls.Add($buttonOK)
    $form2.Controls.Add($buttonCancel)
    $form2.AcceptButton = $buttonOK
    $form2.CancelButton = $buttonCancel
    $form2.ShowDialog($Form_DHCP)
}

function set_listview_columns {
    [void]$LstVw_WorkItems.Clear()
    [void]$LstVw_WorkItems.Columns.Add("IP Address",100)
    [void]$LstVw_WorkItems.Columns.Add("MAC Address",100)
    [void]$LstVw_WorkItems.Columns.Add("Name",80)
    foreach ($dhcpserver in $allservers) { [void]$LstVw_WorkItems.Columns.Add("Status on $dhcpserver",180) }
}

function ip2hex {
param($ip)
    [string]$iphex = ""
    $ip.Split(".") | ForEach-Object { $iphex += (“{0:x}” -f [Int]$_).PadLeft(2, "0") }
    if ($iphex.substring(0,1) -eq "0") { $iphex = $iphex.substring(1) }
    return $iphex
}

function dec2hex {
param([int32]$dec)
    return (“{0:x}” -f [Int]$dec)
}

function hex2dec {
param([string]$hexval)
    return ([Convert]::toInt32($hexval, 16))
}

function hex2ip {
param([string]$iphex)
    $iphex = $iphex.PadLeft(8, "0")
    [string[]]$iparray = @()
    for ($i=0; $i -le 3; $i++) {
        $iparray += [Convert]::toInt32($iphex.substring(($i * 2), 2), 16)
    }
    return ([string]::Join(".", $iparray))
}

function ipadd {
param([string]$ip, [int32]$addval)
    $pos = $true
    if ($addval -lt 0) { $pos = $false; $addval = $addval * -1 }
    [int32]$decval = [Convert]::ToInt32((ip2hex $ip), 16)
    [int32]$adddec = $addval #“{0:x}” -f $addval
    if ($pos) {
        [string]$hexnew = dec2hex ($decval + $adddec)
    } else {
        [string]$hexnew = dec2hex ($decval - $adddec)
    }
    return (hex2ip $hexnew)
}

function set_fixedfont {
param([boolean]$fixed)
    $fontname = "Microsoft Sans Serif"; $fontsize = 11.0
    if ($fixed) { $fontname = "Courier New"; $fontsize = 11.0 }
    $Txt_Results.Font = New-Object Drawing.Font($fontname, $fontsize, [Drawing.GraphicsUnit]::pixel)
}

function Open-SingleForm {
param([switch]$add, [switch]$delete)
    #$txt2_IP.text = ""
    #$txt2_MAC.text = ""
    #$txt2_Name.text = ""
    if ($add) {
        $Chk_EnterIPsForDelete.Checked = $false
        $Form2_NewRes.Text = "Add One Entry"
        $lbl2_MAC.Visible = $true
        $lbl2_Name.Visible = $true
        $txt2_MAC.Visible = $true
        $txt2_Name.Visible = $true
    } elseif ($delete) {
        $Chk_EnterIPsForDelete.Checked = $true
        $Form2_NewRes.Text = "Delete One Entry"
        $lbl2_MAC.Visible = $false
        $lbl2_Name.Visible = $false
        $txt2_MAC.Visible = $false
        $txt2_Name.Visible = $false
    }
    $Form2_NewRes.ShowDialog()
    if ($usesingleform) { run_queries }
}

function Sort-Columns {
    if ($sortcolumn -ne $_.Column) {
        $global:sortcolumn = $_.Column
        $LstVw_WorkItems.Sorting = "Ascending"
    } else {
        if ($LstVw_WorkItems.Sorting -eq "Ascending") {
            $LstVw_WorkItems.Sorting = "Descending"
        } else {
            $LstVw_WorkItems.Sorting = "Ascending"
        }
    }
    $LstVw_WorkItems.Sort()
    $LstVw_WorkItems.ListViewItemSorter = New-Object ListViewItemComparer($_.Column, $LstVw_WorkItems.Sorting)
}




##################  Functions for Scope Creation.  #####################

function Load-NewScopeForm {
    Fill-NewScopeFormServerList
    Set-SubnetMaskValue
}

function Close-NewScopeForm {
    if ($showScopeFormOnly) { $Form_DHCP.Close() }
}

function Fill-NewScopeFormServerList {
    $LstVw_Servers.Items.Clear()
    $fullservers | foreach {
        $newitem = New-Object System.Windows.Forms.ListViewItem
        $newitem.Text = $_
        $LstVw_Servers.Items.Add($newitem)
    }
    $LstVw_Servers.Refresh()
}

function Test-IPAddress {
param($address)
    $regexIPaddress = "\b(25[0-5]|2[0-4][0-9]|[1]?[0-9][0-9]?)\.(25[0-5]|2[0-4][0-9]|[1]?[0-9][0-9]?)\.(25[0-5]|2[0-4][0-9]|[1]?[0-9][0-9]?)\.(25[0-5]|2[0-4][0-9]|[1]?[0-9][0-9]?)\b"
    if ($address -match $regexIPaddress) { return $true }
    return $false
}

function Get-SelectedServers {
    $global:selectedservers = @()
    $LstVw_Servers.SelectedIndices | foreach { $global:selectedservers += $_ }
    $global:checkedservers = @()
    $LstVw_Servers.CheckedIndices | foreach { $global:checkedservers += $_ }
}

function Check-SelectedServers {
    if ($LstVw_Servers.SelectedItems.Count -gt 1) { Reset-SelectedServers }
    $checkedservers = @(); $LstVw_Servers.CheckedItems | foreach { $checkedservers += $_.Text }
    $testActiveServer = $true; if ($LstVw_Servers.SelectedItems.Count -ge 1) { $LstVw_Servers.SelectedItems | foreach { if ($checkedservers -notcontains $_.Text) { $testActiveServer = $false } } }
    if (-not $testActiveServer) { Reset-SelectedServers }
}

function Test-SelectedNotChecked {
    $LstVw_Servers.Items | foreach { if ((-not $_.Checked) -and $_.Selected) { $_.Selected = $false } }
}

function Reset-SelectedServers {
    $LstVw_Servers.Items | foreach {
        if ($selectedservers -contains $_.Index) { $_.Selected = $true } else { $_.Selected = $false }
    }
}

function Check-ServersSelected { 
    if ($LstVw_Servers.CheckedItems.Count -eq 0) {
        if ($LstVw_Servers.SelectedItems.Count -eq 1) { $LstVw_Servers.SelectedItems | foreach { $_.Selected = $false } }
        [Windows.Forms.Messagebox]::Show("Place a checkmark next to each server you want to create the scope on.") | Out-Null
        return $false
    }
    if ($LstVw_Servers.SelectedItems.Count -eq 0) {
        [Windows.Forms.DialogResult]$confirm = [Windows.Forms.Messagebox]::Show("Are you sure you don't want to activate the scope on any servers?", "", [Windows.Forms.MessageboxButtons]::YesNo)
        if ($confirm -eq [Windows.Forms.DialogResult]::No) { return $false }
    }
    return $true
}

function Verify-ScopeInput {
param([NewScope]$scope, [switch]$showmessage = $false)
    if (-not (Check-ServersSelected)) { return $false }
    if (-not (Test-IPAddress $scope.Subnet)) { if ($showmessage) { [Windows.Forms.Messagebox]::Show("Enter a valid Subnet Name.") | Out-Null }; return $false }
    if ($scope.Name -eq "") {  if ($showmessage) { [Windows.Forms.Messagebox]::Show("Enter a name for the scope.") | Out-Null }; return $false }
    if (-not (Test-IPAddress $scope.StartIP)) { if ($showmessage) { [Windows.Forms.Messagebox]::Show("Enter a valid Start IP.") | Out-Null }; return $false }
    if (-not (Test-IPAddress $scope.EndIP)) { if ($showmessage) { [Windows.Forms.Messagebox]::Show("Enter a valid End IP.") | Out-Null }; return $false }
    if ($scope.SubnetLength -lt 1 -or $scope.SubnetLength -gt 31) { if ($showmessage) { [Windows.Forms.Messagebox]::Show("Enter a valid Subnet Mask Length.") | Out-Null }; return $false }
    if (-not (Test-IPAddress $scope.RouterIP)) { if ($showmessage) { [Windows.Forms.Messagebox]::Show("Enter a valid Router IP (Default Gateway).") | Out-Null }; return $false }
    if ($scope.Option1Num -ne "" -and $scope.Option1Num -notmatch "^\d+$") { if ($showmessage) { [Windows.Forms.Messagebox]::Show("Scope Options line 1 `"Number`" value must be a number.") | Out-Null }; return $false }
    if ($scope.Option2Num -ne "" -and $scope.Option2Num -notmatch "^\d+$") { if ($showmessage) { [Windows.Forms.Messagebox]::Show("Scope Options line 2 `"Number`" value must be a number.") | Out-Null }; return $false }
    if ($scope.Option3Num -ne "" -and $scope.Option3Num -notmatch "^\d+$") { if ($showmessage) { [Windows.Forms.Messagebox]::Show("Scope Options line 3 `"Number`" value must be a number.") | Out-Null }; return $false }
    if ($scope.Option1Num -eq "" -xor $scope.Option1Value -eq "") { if ($showmessage) { [Windows.Forms.Messagebox]::Show("Scope Options line 1 has missing data.") | Out-Null }; return $false }
    if ($scope.Option2Num -eq "" -xor $scope.Option2Value -eq "") { if ($showmessage) { [Windows.Forms.Messagebox]::Show("Scope Options line 2 has missing data.") | Out-Null }; return $false }
    if ($scope.Option3Num -eq "" -xor $scope.Option3Value -eq "") { if ($showmessage) { [Windows.Forms.Messagebox]::Show("Scope Options line 3 has missing data.") | Out-Null }; return $false }
    return $true
}

function Confirm-ScopeCreation {
    [Windows.Forms.DialogResult]$response = [Windows.Forms.Messagebox]::Show("Are you sure you want to create the scope?", "", [Windows.Forms.MessageboxButtons]::YesNo)
    if ($response -eq [Windows.Forms.DialogResult]::Yes) { return $true }
    return $false
}

function Get-ScopeInput {
    [NewScope]$scope = Get-ScopeValues 1 $null
    if (-not (Verify-ScopeInput $scope -showmessage)) { return $false }
    if (-not (Confirm-ScopeCreation)) { return $false }
    Create-Scope $scope -showoutput
}

function Create-Scope {
param([NewScope]$scope, [switch]$showoutput)
    $successservers = ""
    $activateserver = ""
    $LstVw_Servers.SelectedItems | foreach { $activateserver = $_.Text }
    $LstVw_Servers.CheckedItems | foreach {
        $dhcpserver = $_.Text
        $netshcmd = "netsh.exe dhcp server \\$dhcpserver"
        #Create the scope
        $command = "$netshcmd add scope $($scope.Subnet) $(ConvertTo-Mask $scope.SubnetLength) `"$($scope.Name)`""
        if ($scope.Description) { $command += " `"$($scope.Description)`"" }
        $description = "Create scope $($scope.Subnet) on $dhcpserver with subnet mask $(ConvertTo-Mask $scope.SubnetLength) and name `"$($scope.Name)`"."
        if (-not (Do-ScopeCommand $scope $dhcpserver $command $description)) { return $false }
        #set active state
        $scopestate = 0; if ($dhcpserver -eq $activateserver) { $scopestate = 1 }
        $command = "$netshcmd scope $($scope.Subnet) set state $scopestate"
        $description = "Set state=$scopestate on scope $($scope.Subnet) on $dhcpserver."
        if (-not (Do-ScopeCommand $scope $dhcpserver $command $description -rollback)) { return $false }
        #add IP range
        $command = "$netshcmd scope $($scope.Subnet) add iprange $($scope.StartIP) $($scope.EndIP)"
        $description = "Add IP range $($scope.StartIP) to $($scope.EndIP) on scope $($scope.Subnet) on $dhcpserver."
        if (-not (Do-ScopeCommand $scope $dhcpserver $command $description -rollback)) { return $false }
        #add router
        $command = "$netshcmd scope $($scope.Subnet) set optionvalue 003 IPADDRESS $($scope.RouterIP)"
        $description = "Add default gateway $($scope.RouterIP) on scope $($scope.Subnet) on $dhcpserver."
        if (-not (Do-ScopeCommand $scope $dhcpserver $command $description -rollback)) { return $false }
        #add domain name
        $command = "$netshcmd scope $($scope.Subnet) set optionvalue 015 STRING `"corp.amvescap.net`""
        $description = "Add domain name `"corp.amvescap.net`" on scope $($scope.Subnet) on $dhcpserver."
        if (-not (Do-ScopeCommand $scope $dhcpserver $command $description -rollback)) { return $false }
        #add options
        if ($scope.Option1Num) {
            $optionnum = $scope.Option1Num.ToString().PadLeft(3, "0")
            $value = Get-OptionValue $scope.Option1Value $scope.Option1Type
            $command = "$netshcmd scope $($scope.Subnet) set optionvalue $optionnum $($scope.Option1Type) $value"
            $description = "Add option $optionnum with value `"$($scope.Option1Value)`" on scope $($scope.Subnet) on $dhcpserver."
            if (-not (Do-ScopeCommand $scope $dhcpserver $command $description -rollback)) { return $false }
        }
        if ($scope.Option2Num) {
            $optionnum = $scope.Option2Num.ToString().PadLeft(3, "0")
            $value = Get-OptionValue $scope.Option2Value $scope.Option2Type
            $command = "$netshcmd scope $($scope.Subnet) set optionvalue $optionnum $($scope.Option2Type) $value"
            $description = "Add option $optionnum with value `"$($scope.Option2Value)`" on scope $($scope.Subnet) on $dhcpserver."
            if (-not (Do-ScopeCommand $scope $dhcpserver $command $description -rollback)) { return $false }
        }
        if ($scope.Option3Num) {
            $optionnum = $scope.Option3Num.ToString().PadLeft(3, "0")
            $value = Get-OptionValue $scope.Option3Value $scope.Option3Type
            $command = "$netshcmd scope $($scope.Subnet) set optionvalue $optionnum $($scope.Option3Type) $value"
            $description = "Add option $optionnum with value `"$($scope.Option3Value)`" on scope $($scope.Subnet) on $dhcpserver."
            if (-not (Do-ScopeCommand $scope $dhcpserver $command $description -rollback)) { return $false }
        }
        $successservers += "$dhcpserver`r`n"
    }
    if ($successservers -and $showoutput) { [Windows.Forms.Messagebox]::Show("Scope created successfully on:`r`n`r`n$successservers") }
    $Btn_Cancel.Text = "Close"
    $Btn_Cancel.Refresh()
    return $true
}

function Get-OptionValue {
param($value, $type)
    switch ($type) {
        "STRING" { return "`"$($value)`"" }
        "IPADDRESS" { return $value }
        "BINARY" {
            $hexvalue = ""
            foreach ($element in $value.ToCharArray()) { $hexvalue += [System.String]::Format("{0:X}", [System.Convert]::ToUInt32($element)) }
            return $hexvalue
        }
    }
}


function Do-ScopeCommand {
param([NewScope]$scope, $server, $command, $description, [switch]$rollback)
    [string]$logfiledate = "{0:yyyyMMddhhmmss}" -f [DateTime](Get-Date)
    $descriptionfixed = $description.Replace("`"", "") -replace "[^\w\s,;.-]", ""
    if ($descriptionfixed.length -gt 64) { $descriptionfixed = $descriptionfixed.substring(0,64) }
    $logfilename = "$logfolder\$logfiledate $descriptionfixed.txt"
    $objStartInfo = New-Object System.Diagnostics.ProcessStartInfo
    $objStartInfo.FileName = "cmd.exe"
    $objStartInfo.WindowStyle = "Hidden"
    $objStartInfo.Arguments = "/C $command >`"$logfilename`""
    [System.Diagnostics.Process]::Start($objStartInfo).WaitForExit() | Out-Null
    $logfileinfo = Get-Content $logfilename
    $success = $false
    $logfileinfo | Foreach-Object { if ($_ -match $commandsuccessful) { $success = $true } }
    if (-not $success) {
        [Windows.Forms.Messagebox]::Show("$command`r`n`r`n$([String]::Join("`r`n", $logfileinfo))", "Error encountered") | Out-Null
        if ($rollback) { Rollback-ScopeCreation $scope $server | Out-Null }
        return $false
    }
    return $true
}

function Rollback-ScopeCreation {
param([NewScope]$scope, $server)
    [Windows.Forms.DialogResult]$confirmrollback = [Windows.Forms.Messagebox]::Show("The scope configuration failed on $server.  Do you want to rollback the operation?", "Error",  [Windows.Forms.MessageboxButtons]::YesNo)
    if ($confirmrollback -eq [Windows.Forms.DialogResult]::Yes) {
        $netshcmd = "netsh.exe dhcp server \\$dhcpserver"
        $command = "$netshcmd delete scope $($scope.Subnet) DHCPFULLFORCE"
        $description = "Delete scope $($scope.Subnet) on $dhcpserver (rollback creation)."
        $success = Do-ScopeCommand $scope $dhcpserver $command $description
        if ($success) {
            [Windows.Forms.Messagebox]::Show("Rollback successful.") | Out-Null
        } else {
            [Windows.Forms.Messagebox]::Show("Rollback failed.  The scope will need to be deleted manually.") | Out-Null
        }
    }
}

function Set-SubnetMaskValue {
    $Lbl_SubnetIPValue.Text = "IP value: $(ConvertTo-Mask $Num_SubnetLength.Value)"
    $Lbl_SubnetIPValue.Refresh()
}

Function ConvertTo-Mask( [Byte]$MaskLength ) {
  Return ConvertTo-DottedDecimalIP ([Convert]::ToUInt32($(("1" * $MaskLength).PadRight(32, "0")), 2))
}

Function ConvertTo-DottedDecimalIP( [String]$IP ) {
  Switch -RegEx ($IP) {
    "([01]{8}\.){3}[01]{8}" { Return [String]::Join('.', $( $IP.Split('.') | %{ [Convert]::ToInt32($_, 2) } )) }
    "\d" { $IP = [UInt32]$IP; $DottedIP = $( For ($i = 3; $i -gt -1; $i--) { $Remainder = $IP % [Math]::Pow(256, $i); ($IP - $Remainder) / [Math]::Pow(256, $i); $IP = $Remainder } ); Return [String]::Join('.', $DottedIP); }
  }
}

add-type @" 
public struct NewScope { 
   public string Subnet; 
   public string Name; 
   public string Description;
   public string StartIP;
   public string EndIP;
   public int SubnetLength;
   public string RouterIP;
   public bool Activate;
   public string Option1Num;
   public string Option1Type;
   public string Option1Value;
   public string Option2Num;
   public string Option2Type;
   public string Option2Value;
   public string Option3Num;
   public string Option3Type;
   public string Option3Value;
} 
"@ 

function Import-ScopeFile {
    if (-not (Check-ServersSelected)) { return $false }
    $openFileDialog1.ShowHelp = $True
    $openFileDialog1.FileName = ""
    $openFileDialog1.ShowDialog()
    if ($openFileDialog1.FileName) {
        $csv = Import-Csv $openFileDialog1.FileName
        $linecount = 0
        $alldataok = $true
        $csv | foreach {
            $linecount++
            [NewScope]$scope = Get-ScopeValues 2 $_
            if (-not (Verify-ScopeInput $scope)) {
                [Windows.Forms.Messagebox]::Show("Invalid data on line $linecount.")
                $alldataok = $false
            }
        }
        if ($alldataok) {
            if ((Confirm-ScopeCreation)) {
                $linecount = 0
                $csv | foreach {
                    $linecount++
                    [NewScope]$scope = Get-ScopeValues 2 $_
                    $success = Create-Scope $scope
                    if (-not $success) {
                        [Windows.Forms.Messagebox]::Show("Scope creation failed on line $linecount.")
                        return
                    }
                }
                [Windows.Forms.Messagebox]::Show("Scope creation finished.")
            }
        }
    }
}

function Concat-Subnet {
param($subnet, $lastoctets)
    $numlastoctets = $lastoctets.Split(".").length
    $returnvalue = ""
    for ($x=0;$x -le (3 - $numlastoctets); $x++) {
        $returnvalue += "$($subnet.Split(`".`")[$x])."
    }
    $returnvalue += $lastoctets
    $returnvalue
}

function Get-ScopeValues {
param($mode, $csvline)
    [NewScope]$scope = New-Object NewScope
    if ($mode -eq 1) {
        $scope.Subnet = $Txt_Subnet.Text
        $scope.Name = $Txt_Name.Text
        $scope.Description = $Txt_Description.Text
        $scope.StartIP = $Txt_StartIP.Text
        $scope.EndIP = $Txt_EndIP.Text
        $scope.SubnetLength = $Num_SubnetLength.Value
        $scope.RouterIP = $Txt_RouterIP.Text
        $scope.Option1Num = $Txt_Option1Num.Text
        $scope.Option1Type = $Cbo_Option1Type.Text
        $scope.Option1Value = $Txt_Option1Value.Text
        $scope.Option2Num = $Txt_Option2Num.Text
        $scope.Option2Type = $Cbo_Option2Type.Text
        $scope.Option2Value = $Txt_Option2Value.Text
        $scope.Option3Num = $Txt_Option3Num.Text
        $scope.Option3Type = $Cbo_Option3Type.Text
        $scope.Option3Value = $Txt_Option3Value.Text
    } else {
        $scope.Subnet = $csvline.Subnet
        $scope.Name = $csvline.Name
        $scope.Description = $csvline.Description
        $scope.StartIP = $csvline.StartIP
        if ($scope.StartIP.Split(".").length -ne 4) { $scope.StartIP = Concat-Subnet $scope.Subnet $scope.StartIP }
        $scope.EndIP = $csvline.EndIP
        if ($scope.EndIP.Split(".").length -ne 4) { $scope.EndIP = Concat-Subnet $scope.Subnet $scope.EndIP }
        $scope.SubnetLength = [int]$csvline.SubnetLength
        $scope.RouterIP = $csvline.RouterIP
        if ($scope.RouterIP.Split(".").length -ne 4) { $scope.RouterIP = Concat-Subnet $scope.Subnet $scope.RouterIP }
        $scope.Option1Num = $csvline.Option1Num
        $scope.Option1Type = "STRING"
        $scope.Option1Value = $csvline.Option1Value
        $scope.Option2Num = $csvline.Option2Num
        $scope.Option2Type = "STRING"
        $scope.Option2Value = $csvline.Option2Value
        $scope.Option3Num = $csvline.Option3Num
        $scope.Option3Type = "STRING"
        $scope.Option3Value = $csvline.Option3Value
    }
    $scope
}

function Resolve-NameToIP {
param($name)
    try {
        $ipresolve = [System.Net.Dns]::GetHostAddresses($name) | ? { $_.AddressFamily -eq "InterNetwork" }
        if ($ipresolve) {
            $ipaddress = $ipresolve.IPAddressToString
        } else {
            $ipaddress += "Error: Unable to get IP address"
        }
    } catch {
        if ($_.exception.message.contains("No such host is known")) {
            $ipaddress += "Error: Host not found"
        } else {
            $ipaddress += "Error: $($_.exception.message)"
        }
    }
    $ipaddress
}

function Set-Instructions{
    if ($this.text -eq $checklabel_delete) {
        if ($Chk_GetExistingLeases.Checked) { $Chk_GetExistingLeases.Checked = $false }
        $Lbl_Instructions.Text = $instructionsdelete
    } elseif ($this.text -eq $checklabel_getleaseinfo) {
        if ($Chk_EnterIPsForDelete.Checked) { $Chk_EnterIPsForDelete.Checked = $false }
        $Lbl_Instructions.Text = $instructionsgetleaseinfo
    } 
    if (-not $Chk_GetExistingLeases.Checked -and -not $Chk_EnterIPsForDelete.Checked) {
        $Lbl_Instructions.Text = $instructionsnormal
    }
}

function Copy-ListToClipboard {
    $output = ""
    $LstVw_WorkItems.Items | foreach {
        $line = ""
        for ($x=0; $x -lt $_.SubItems.count; $x++) { $line += "`t" + $_.SubItems[$x].text }
        $output += $line.substring(1) + "`r`n"
    }
    [windows.forms.clipboard]::SetText($output)
}



########################################################################
# Define Windows Form for Main Form.
########################################################################

$columnClick = { Sort-Columns }

$basewidth = 310
$baseheight = 360
$edgespacing = 5
$boxspacing = 10
$checkspacing = 25
$radiospacing = 20
$rightcheckcolumn = 480

$Form_Components = New-Object System.ComponentModel.Container

[Windows.Forms.Form]$Form_DHCP = New-Object Windows.Forms.Form
$Form_DHCP.Text = "DHCP Reservation Tool"
$Form_DHCP.Size = New-Object Drawing.Point 710,700
$Form_DHCP.MinimumSize = $Form_DHCP.Size
$Form_DHCP.add_FormClosing({ $Form_DHCP.Visible = $false })

[Windows.Forms.Label]$Lbl_Instructions = New-Object Windows.Forms.Label
$Lbl_Instructions.Location = New-Object Drawing.Point $boxspacing,$boxspacing
$Lbl_Instructions.Size = New-Object Drawing.Point 340,40
$Lbl_Instructions.Text = $instructionsnormal

[Windows.Forms.Checkbox]$Chk_RefreshTempData = New-Object Windows.Forms.Checkbox
$Chk_RefreshTempData.Text = "Refresh temporary scope data"
$Chk_RefreshTempData.Location = New-Object Drawing.Point $rightcheckcolumn,$boxspacing
$Chk_RefreshTempData.Size = New-Object Drawing.Point 250, 25
$Chk_RefreshTempData.Anchor = "Top,Right"
$Chk_RefreshTempData.add_click({ check_actions })

[Windows.Forms.Checkbox]$Chk_EnterIPsForDelete = New-Object Windows.Forms.Checkbox
$Chk_EnterIPsForDelete.Text = $checklabel_delete
$Chk_EnterIPsForDelete.Location = New-Object Drawing.Point $rightcheckcolumn,($Chk_RefreshTempData.Top + $checkspacing - 5)
$Chk_EnterIPsForDelete.Size = New-Object Drawing.Point 250, 25
$Chk_EnterIPsForDelete.Anchor = "Top,Right"
$Chk_EnterIPsForDelete.add_click({ Set-Instructions })

[Windows.Forms.Checkbox]$Chk_GetExistingLeases = New-Object Windows.Forms.Checkbox
$Chk_GetExistingLeases.Text = $checklabel_getleaseinfo
$Chk_GetExistingLeases.Location = New-Object Drawing.Point $rightcheckcolumn,($Chk_EnterIPsForDelete.Top + $checkspacing - 5)
$Chk_GetExistingLeases.Size = New-Object Drawing.Point 250, 25
$Chk_GetExistingLeases.Anchor = "Top,Right"
$Chk_GetExistingLeases.add_click({ Set-Instructions })

$ContextMenuStrip1 = New-Object System.Windows.Forms.ContextMenuStrip($Form_Components)

$CopyListMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$CopyListMenuItem.Name = "CopyListMenuItem"
$CopyListMenuItem.Text = "&Copy list to clipboard"
$CopyListMenuItem.add_Click({ Copy-ListToClipboard })
$ContextMenuStrip1.Items.Add($CopyListMenuItem) | Out-Null


[Windows.Forms.ListView]$LstVw_WorkItems = New-Object Windows.Forms.ListView
$LstVw_WorkItems.Location = new-object Drawing.Point $boxspacing,($Lbl_Instructions.Bottom + 20 + $boxspacing)
$LstVw_WorkItems.Size = new-object Drawing.Point ($Form_DHCP.Width - 35), ($Form_DHCP.Height - 505) 
$LstVw_WorkItems.FullRowSelect = $True
$LstVw_WorkItems.View = "Details"
$LstVw_WorkItems.Anchor = "Top,Left,Right,Bottom"
$LstVw_WorkItems.add_click({ show_info })
$LstVw_WorkItems.add_SelectedIndexChanged({ show_info })
$LstVw_WorkItems.HideSelection = $false
$LstVw_WorkItems.MultiSelect = $false
$LstVw_WorkItems.Scrollable = $true
set_listview_columns
$LstVw_WorkItems.Add_ColumnClick($columnClick)
$LstVw_WorkItems.ContextMenuStrip = $ContextMenuStrip1

[Windows.Forms.Textbox]$Txt_Results = New-Object Windows.Forms.Textbox
$Txt_Results.Multiline = $true
$Txt_Results.ScrollBars = "Vertical,Horizontal"
$Txt_Results.Location = New-Object Drawing.Point $boxspacing,($LstVw_WorkItems.Top + $LstVw_WorkItems.Height + $boxspacing)
$Txt_Results.Size = New-Object Drawing.Point ($LstVw_WorkItems.Width),180
$Txt_Results.Text = ""
$Txt_Results.Anchor = "Right,Left,Bottom"
$Txt_Results.WordWrap = $false
$Txt_Results.ReadOnly = $true
$Txt_Results.BackColor = "White"

[Windows.Forms.Checkbox]$Chk_ShowNetShCmds = New-Object Windows.Forms.Checkbox
$Chk_ShowNetShCmds.Text = "Show netsh commands in preview"
$Chk_ShowNetShCmds.Location = New-Object Drawing.Point $rightcheckcolumn,($Txt_Results.Top + $Txt_Results.Height + $boxspacing)
$Chk_ShowNetShCmds.Size = New-Object Drawing.Point 250, 25
$Chk_ShowNetShCmds.Anchor = "Bottom,Right"
$Chk_ShowNetShCmds.Visible = $false
$Chk_ShowNetShCmds.add_click({ if ($Btn_Preview.Text -eq $ButtonText_DoIt) { show_preview } })

[Windows.Forms.GroupBox]$Grp_Show = New-Object Windows.Forms.GroupBox
$Grp_Show.Location = New-Object Drawing.Point $boxspacing,($Txt_Results.Top + $Txt_Results.Height + $boxspacing)
$Grp_Show.Size = New-Object Drawing.Point 200, 125
$Grp_Show.Text = "Show:"
$Grp_Show.Anchor = "Left,Bottom"

[Windows.Forms.RadioButton]$Rdo_ServerInfo = New-Object Windows.Forms.RadioButton
$Rdo_ServerInfo.Text = "Server info"
$Rdo_ServerInfo.Location = New-Object Drawing.Point $boxspacing, 15
$Rdo_ServerInfo.Size = New-Object Drawing.Point 180, 25
$Rdo_ServerInfo.add_click({ show_info })
$Rdo_ServerInfo.Parent = $Grp_Show
$Rdo_ServerInfo.Checked = $true

[Windows.Forms.RadioButton]$Rdo_AllScopes = New-Object Windows.Forms.RadioButton
$Rdo_AllScopes.Text = "All Scopes"
$Rdo_AllScopes.Location = New-Object Drawing.Point $boxspacing, ($Rdo_ServerInfo.Top + $radiospacing)
$Rdo_AllScopes.Size = New-Object Drawing.Point 180, 25
$Rdo_AllScopes.add_click({ show_info })
$Rdo_AllScopes.Parent = $Grp_Show

[Windows.Forms.RadioButton]$Rdo_AddressPool = New-Object Windows.Forms.RadioButton
$Rdo_AddressPool.Text = "Address pool for this scope"
$Rdo_AddressPool.Location = New-Object Drawing.Point $boxspacing, ($Rdo_AllScopes.Top + $radiospacing)
$Rdo_AddressPool.Size = New-Object Drawing.Point 180, 25
$Rdo_AddressPool.add_click({ show_info })
$Rdo_AddressPool.Parent = $Grp_Show

[Windows.Forms.RadioButton]$Rdo_Exclusions = New-Object Windows.Forms.RadioButton
$Rdo_Exclusions.Text = "Exclusions in this scope"
$Rdo_Exclusions.Location = New-Object Drawing.Point $boxspacing, ($Rdo_AddressPool.Top + $radiospacing)
$Rdo_Exclusions.Size = New-Object Drawing.Point 180, 25
$Rdo_Exclusions.add_click({ show_info })
$Rdo_Exclusions.Parent = $Grp_Show

[Windows.Forms.RadioButton]$Rdo_Reservations = New-Object Windows.Forms.RadioButton
$Rdo_Reservations.Text = "Reservations in this scope"
$Rdo_Reservations.Location = New-Object Drawing.Point $boxspacing, ($Rdo_Exclusions.Top + $radiospacing)
$Rdo_Reservations.Size = New-Object Drawing.Point 180, 25
$Rdo_Reservations.add_click({ show_info })
$Rdo_Reservations.Parent = $Grp_Show

[Windows.Forms.GroupBox]$Grp_Actions = New-Object Windows.Forms.GroupBox
$Grp_Actions.Location = New-Object Drawing.Point ($Grp_Show.Width + 40),($Txt_Results.Top + $Txt_Results.Height + $boxspacing)
$Grp_Actions.Size = New-Object Drawing.Point 220, 175
$Grp_Actions.Text = "Actions"
$Grp_Actions.Anchor = "Left,Bottom"
$Grp_Actions.Visible = $false

[Windows.Forms.Checkbox]$Chk_SkipItemsNotInScope = New-Object Windows.Forms.Checkbox
$Chk_SkipItemsNotInScope.Text = "Skip items that are not in any scope"
$Chk_SkipItemsNotInScope.Location = New-Object Drawing.Point $boxspacing,25
$Chk_SkipItemsNotInScope.Size = New-Object Drawing.Point 205, 25
$Chk_SkipItemsNotInScope.Parent = $Grp_Actions
$Chk_SkipItemsNotInScope.add_click({ check_actions })

[Windows.Forms.Checkbox]$Chk_RmvConfExclusions = New-Object Windows.Forms.Checkbox
$Chk_RmvConfExclusions.Text = "Remove conflicting exclusions"
$Chk_RmvConfExclusions.Location = New-Object Drawing.Point $boxspacing, ($Chk_SkipItemsNotInScope.Top + $checkspacing)
$Chk_RmvConfExclusions.Size = New-Object Drawing.Point 200, 25
$Chk_RmvConfExclusions.Parent = $Grp_Actions
$Chk_RmvConfExclusions.add_click({ check_actions })

[Windows.Forms.Checkbox]$Chk_DelExistReservations = New-Object Windows.Forms.Checkbox
$Chk_DelExistReservations.Text = "Delete existing reservations"
$Chk_DelExistReservations.Location = New-Object Drawing.Point $boxspacing, ($Chk_RmvConfExclusions.Top + $checkspacing)
$Chk_DelExistReservations.Size = New-Object Drawing.Point 200, 25
$Chk_DelExistReservations.Parent = $Grp_Actions
$Chk_DelExistReservations.add_click({ check_actions })

[Windows.Forms.Checkbox]$Chk_CreateReservations = New-Object Windows.Forms.Checkbox
$Chk_CreateReservations.Text = "Create reservations"
$Chk_CreateReservations.Location = New-Object Drawing.Point $boxspacing, ($Chk_DelExistReservations.Top + $checkspacing)
$Chk_CreateReservations.Size = New-Object Drawing.Point 200, 25
$Chk_CreateReservations.Parent = $Grp_Actions
$Chk_CreateReservations.add_click({ check_actions })

[Windows.Forms.Button]$Btn_GetInfo = New-Object Windows.Forms.Button
$Btn_GetInfo.Text = "Get Info"
$Btn_GetInfo.add_click({ $usesingleform = $false; run_queries })
$Btn_GetInfo.Location = New-Object Drawing.Point $boxspacing,($Grp_Show.Top + $Grp_Show.Height + $boxspacing)
$Btn_GetInfo.Anchor = "Left,Bottom"

[Windows.Forms.Button]$Btn_AddOne = New-Object Windows.Forms.Button
$Btn_AddOne.Text = "Add One Entry"
$Btn_AddOne.add_click({ Open-SingleForm -add })
$Btn_AddOne.Location = New-Object Drawing.Point $boxspacing, ($Grp_Show.Top + $Grp_Show.Height + $boxspacing + $Btn_GetInfo.Height + 5)
$Btn_AddOne.Anchor = "Left,Bottom"
$Btn_AddOne.Width = 100

[Windows.Forms.Button]$Btn_DeleteOne = New-Object Windows.Forms.Button
$Btn_DeleteOne.Text = "Delete One Entry"
$Btn_DeleteOne.add_click({ Open-SingleForm -delete })
$Btn_DeleteOne.Location = New-Object Drawing.Point ($boxspacing + $Btn_AddOne.Width + 5), ($Grp_Show.Top + $Grp_Show.Height + $boxspacing + $Btn_GetInfo.Height + 5)
$Btn_DeleteOne.Anchor = "Left,Bottom"
$Btn_DeleteOne.Width = 100

[Windows.Forms.Button]$Btn_Close = New-Object Windows.Forms.Button
$Btn_Close.Text = "Close"
$Btn_Close.Location = New-Object Drawing.Point ($LstVw_WorkItems.Left + $LstVw_WorkItems.Width - $Btn_Close.Width),($Grp_Show.Top + $Grp_Show.Height + $boxspacing + $Btn_GetInfo.Height + 5)
$Btn_Close.Anchor = "Right,Bottom"

[Windows.Forms.Button]$Btn_NewScope = New-Object Windows.Forms.Button
$Btn_NewScope.Text = "Create New Scopes"
$Btn_NewScope.Width = 120
$Btn_NewScope.Location = New-Object Drawing.Point ($LstVw_WorkItems.Right - $Btn_Close.Width - $Btn_NewScope.Width - 10),($Grp_Show.Top + $Grp_Show.Height + $boxspacing + $Btn_GetInfo.Height + 5)
$Btn_NewScope.Anchor = "Right,Bottom"
$Btn_NewScope.add_click({ $Form_CreateScope.ShowDialog($Form_DHCP) })

[Windows.Forms.Button]$Btn_GetLeaseInfo = New-Object Windows.Forms.Button
$Btn_GetLeaseInfo.Text = "Get Info For Existing Leases"
$Btn_GetLeaseInfo.Width = 205
$Btn_GetLeaseInfo.Location = New-Object Drawing.Point ($Btn_NewScope.Left),($Btn_GetInfo.Top)
$Btn_GetLeaseInfo.Anchor = "Right,Bottom"
$Btn_GetLeaseInfo.add_click({ Get-LeaseInfo })
$Btn_GetLeaseInfo.Visible = $false

[Windows.Forms.Button]$Btn_Preview = New-Object Windows.Forms.Button
$Btn_Preview.Text = $ButtonText_Preview
$Btn_Preview.Location = New-Object Drawing.Point $boxspacing, ($Chk_CreateReservations.Top + $checkspacing + 15)
$Btn_Preview.Size = New-Object Drawing.Point 160, 25
$Btn_Preview.Anchor = "Right,Bottom"
$Btn_Preview.Enabled = $false
$Btn_Preview.Parent = $Grp_Actions
$Btn_Preview.add_click({ if ($Btn_Preview.Text -eq $ButtonText_Preview) { show_preview } elseif ($Btn_Preview.Text -eq $ButtonText_DoIt) { do_changes } })

[Windows.Forms.Button]$Btn_SelectServers = New-Object Windows.Forms.Button
$Btn_SelectServers.Text = "Select Servers"
$Btn_SelectServers.add_click({ change_servers })
$Btn_SelectServers.Location = New-Object Drawing.Point 360, ($boxspacing + 5)
$Btn_SelectServers.Size = New-Object Drawing.Point 100, $Btn_GetInfo.Height
$Btn_SelectServers.Anchor = "Right,Top"

$Form_DHCP.Controls.Add($Lbl_Instructions)
$Form_DHCP.Controls.Add($Chk_GetExistingLeases)
$Form_DHCP.Controls.Add($Chk_EnterIPsForDelete)
$Form_DHCP.Controls.Add($Chk_ShowNetShCmds)
$Form_DHCP.Controls.Add($LstVw_WorkItems)
$Form_DHCP.Controls.Add($Txt_Results)
$Form_DHCP.Controls.Add($Btn_GetInfo)
$Form_DHCP.Controls.Add($Btn_Close)
$Form_DHCP.Controls.Add($Grp_Show)
$Form_DHCP.Controls.Add($Grp_Actions)
$Form_DHCP.Controls.Add($Chk_RefreshTempData)
$Form_DHCP.Controls.Add($Btn_SelectServers)
$Form_DHCP.Controls.Add($Btn_AddOne)
$Form_DHCP.Controls.Add($Btn_DeleteOne)
$Form_DHCP.Controls.Add($Btn_NewScope)
$Form_DHCP.Controls.Add($Btn_GetLeaseInfo)

$Grp_Show.Controls.Add($Rdo_ServerInfo)
$Grp_Show.Controls.Add($Rdo_AllScopes)
$Grp_Show.Controls.Add($Rdo_AddressPool)
$Grp_Show.Controls.Add($Rdo_Exclusions)
$Grp_Show.Controls.Add($Rdo_Reservations)

$Grp_Actions.Controls.Add($Chk_SkipItemsNotInScope)
$Grp_Actions.Controls.Add($Chk_RmvConfExclusions)
$Grp_Actions.Controls.Add($Chk_DelExistReservations)
$Grp_Actions.Controls.Add($Chk_CreateReservations)
$Grp_Actions.Controls.Add($Btn_Preview)

$Form_DHCP.AcceptButton = $Btn_GetInfo
$Form_DHCP.CancelButton = $Btn_Close
$Form_DHCP.add_Load($Form_DHCP_Load)


########################################################################
# End Windows Form for Main Form.
########################################################################


########################################################################
# Define Windows Form for New DHCP Reservation.
########################################################################

#region Generated Form Objects
$Form2_NewRes = New-Object System.Windows.Forms.Form
$Lbl2_IP = New-Object System.Windows.Forms.Label
$Lbl2_MAC = New-Object System.Windows.Forms.Label
$Lbl2_Name = New-Object System.Windows.Forms.Label
$Txt2_IP = New-Object System.Windows.Forms.TextBox
$Txt2_MAC = New-Object System.Windows.Forms.TextBox
$Txt2_Name = New-Object System.Windows.Forms.TextBox
$Btn2_Clear = New-Object System.Windows.Forms.Button
$Btn2_OK = New-Object System.Windows.Forms.Button
$Btn2_Cancel = New-Object System.Windows.Forms.Button
$InitialFormWindowState = New-Object System.Windows.Forms.FormWindowState
#endregion Generated Form Objects

#----------------------------------------------
#Generated Event Script Blocks
#----------------------------------------------
#Provide Custom Code for events specified in PrimalForms.
$Btn2_Clear_OnClick = {
    $Txt2_IP.Text = ""
    $Txt2_MAC.Text = ""
    $Txt2_Name.Text = ""
}

$Btn2_Cancel_OnClick = {
    $usesingleform = $false  
}

$Btn2_OK_OnClick = {
    $usesingleform = $true
    $formvalidation = $true
    if ($Txt2_MAC.Visible -and -not ($Txt2_MAC.Text)) { $formvalidation = $false }
    if ($formvalidation) {
        $Form2_NewRes.Close()
    } else {
        [Windows.Forms.Messagebox]::Show("Some fields are incomplete.")
    }
}

$OnLoadForm_StateCorrection = {
	$Form2_NewRes.WindowState = $InitialFormWindowState
}

$Form_DHCP_Load = {
    if ($showScopeFormOnly) { $Btn_NewScope.PerformClick() }
}

#----------------------------------------------
#region Generated Form Code
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 300
$System_Drawing_Size.Height = 150
$Form2_NewRes.Size = $System_Drawing_Size
$Form2_NewRes.Text = "New DCHP Reservation"
$Form2_NewRes.Name = "Form2_NewRes"
$Form2_NewRes.DataBindings.DefaultDataSourceUpdateMode = 0
$Form2_NewRes.CancelButton = $Btn2_Cancel
$Form2_NewRes.FormBorderStyle = 5

$Lbl2_IP.TabIndex = 0
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 68
$System_Drawing_Size.Height = 23
$Lbl2_IP.Size = $System_Drawing_Size
$Lbl2_IP.Text = "IP Address:"
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 12
$System_Drawing_Point.Y = 9
$Lbl2_IP.Location = $System_Drawing_Point
$Lbl2_IP.DataBindings.DefaultDataSourceUpdateMode = 0
$Lbl2_IP.Name = "Lbl2_IP"
$Form2_NewRes.Controls.Add($Lbl2_IP)

$Lbl2_MAC.TabIndex = 1
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 80
$System_Drawing_Size.Height = 23
$Lbl2_MAC.Size = $System_Drawing_Size
$Lbl2_MAC.Text = "MAC Address:"
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 12
$System_Drawing_Point.Y = 32
$Lbl2_MAC.Location = $System_Drawing_Point
$Lbl2_MAC.DataBindings.DefaultDataSourceUpdateMode = 0
$Lbl2_MAC.Name = "Lbl2_MAC"
$Form2_NewRes.Controls.Add($Lbl2_MAC)

$Lbl2_Name.TabIndex = 2
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 68
$System_Drawing_Size.Height = 37
$Lbl2_Name.Size = $System_Drawing_Size
$Lbl2_Name.Text = "Name: (optional)"
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 12
$System_Drawing_Point.Y = 55
$Lbl2_Name.Location = $System_Drawing_Point
$Lbl2_Name.DataBindings.DefaultDataSourceUpdateMode = 0
$Lbl2_Name.Name = "Lbl2_Name"
$Form2_NewRes.Controls.Add($Lbl2_Name)

$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 196
$System_Drawing_Size.Height = 20
$Txt2_IP.Size = $System_Drawing_Size
$Txt2_IP.DataBindings.DefaultDataSourceUpdateMode = 0
$Txt2_IP.Anchor = 13
$Txt2_IP.Name = "Txt2_IP"
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 86
$System_Drawing_Point.Y = 6
$Txt2_IP.Location = $System_Drawing_Point
$Txt2_IP.TabIndex = 3
$Form2_NewRes.Controls.Add($Txt2_IP)
$Txt2_IP.BringToFront()

$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 196
$System_Drawing_Size.Height = 20
$Txt2_MAC.Size = $System_Drawing_Size
$Txt2_MAC.DataBindings.DefaultDataSourceUpdateMode = 0
$Txt2_MAC.Anchor = 13
$Txt2_MAC.Name = "Txt2_MAC"
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 86
$System_Drawing_Point.Y = 29
$Txt2_MAC.Location = $System_Drawing_Point
$Txt2_MAC.TabIndex = 4
$Form2_NewRes.Controls.Add($Txt2_MAC)
$Txt2_MAC.BringToFront()

$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 196
$System_Drawing_Size.Height = 20
$Txt2_Name.Size = $System_Drawing_Size
$Txt2_Name.DataBindings.DefaultDataSourceUpdateMode = 0
$Txt2_Name.Anchor = 13
$Txt2_Name.Name = "Txt2_Name"
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 86
$System_Drawing_Point.Y = 51
$Txt2_Name.Location = $System_Drawing_Point
$Txt2_Name.TabIndex = 5
$Form2_NewRes.Controls.Add($Txt2_Name)
$Txt2_Name.BringToFront()

$Btn2_Clear.TabIndex = 6
$Btn2_Clear.Anchor = 10
$Btn2_Clear.Name = "Btn2_Clear"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 75
$System_Drawing_Size.Height = 23
$Btn2_Clear.Size = $System_Drawing_Size
$Btn2_Clear.UseVisualStyleBackColor = $True
$Btn2_Clear.Text = "Clear"
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 6
$System_Drawing_Point.Y = 93
$Btn2_Clear.Location = $System_Drawing_Point
$Btn2_Clear.DataBindings.DefaultDataSourceUpdateMode = 0
$Btn2_Clear.add_Click($Btn2_Clear_OnClick)
$Form2_NewRes.Controls.Add($Btn2_Clear)

$Btn2_OK.TabIndex = 6
$Btn2_OK.Anchor = 10
$Btn2_OK.Name = "Btn2_OK"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 75
$System_Drawing_Size.Height = 23
$Btn2_OK.Size = $System_Drawing_Size
$Btn2_OK.UseVisualStyleBackColor = $True
$Btn2_OK.Text = "OK"
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 126
$System_Drawing_Point.Y = 93
$Btn2_OK.Location = $System_Drawing_Point
$Btn2_OK.DataBindings.DefaultDataSourceUpdateMode = 0
$Btn2_OK.add_Click($Btn2_OK_OnClick)
$Form2_NewRes.Controls.Add($Btn2_OK)

$Btn2_Cancel.TabIndex = 7
$Btn2_Cancel.Anchor = 10
$Btn2_Cancel.Name = "Btn2_Cancel"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 75
$System_Drawing_Size.Height = 23
$Btn2_Cancel.Size = $System_Drawing_Size
$Btn2_Cancel.UseVisualStyleBackColor = $True
$Btn2_Cancel.Text = "Cancel"
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 207
$System_Drawing_Point.Y = 93
$Btn2_Cancel.Location = $System_Drawing_Point
$Btn2_Cancel.DataBindings.DefaultDataSourceUpdateMode = 0
$Btn2_Cancel.add_Click($Btn2_Cancel_OnClick)
$Form2_NewRes.Controls.Add($Btn2_Cancel)

#endregion Generated Form Code

#Save the initial state of the form
$InitialFormWindowState = $Form2_NewRes.WindowState
#Init the OnLoad event to correct the initial state of the form
$Form2_NewRes.add_Load($OnLoadForm_StateCorrection)


########################################################################
# End Windows Form for New DHCP Reservation.
########################################################################



########################################################################
# Define Windows Form for Scope Creation.
########################################################################

$Form_CreateScope = New-Object System.Windows.Forms.Form
$Lbl_Servers = New-Object System.Windows.Forms.Label
$LstVw_Servers = New-Object System.Windows.Forms.ListView
$LstVw_Columns = New-Object System.Windows.Forms.ColumnHeader
$Btn_Import = New-Object System.Windows.Forms.Button
$Lbl_Subnet = New-Object System.Windows.Forms.Label
$Txt_Subnet = New-Object System.Windows.Forms.TextBox
$Lbl_SubnetHelp = New-Object System.Windows.Forms.Label
$Lbl_Name = New-Object System.Windows.Forms.Label
$Txt_Name = New-Object System.Windows.Forms.TextBox
$Lbl_Description = New-Object System.Windows.Forms.Label
$Txt_Description = New-Object System.Windows.Forms.TextBox
$Lbl_StartIP = New-Object System.Windows.Forms.Label
$Txt_StartIP = New-Object System.Windows.Forms.TextBox
$Lbl_EndIP = New-Object System.Windows.Forms.Label
$Txt_EndIP = New-Object System.Windows.Forms.TextBox
$Lbl_SubnetLength = New-Object System.Windows.Forms.Label
$Num_SubnetLength = New-Object System.Windows.Forms.NumericUpDown
$Lbl_SubnetIPValue = New-Object System.Windows.Forms.Label
$Lbl_RouterIP = New-Object System.Windows.Forms.Label
$Txt_RouterIP = New-Object System.Windows.Forms.TextBox
$Grp_OtherOptions = New-Object System.Windows.Forms.GroupBox
$Lbl_OtherOptionNumber = New-Object System.Windows.Forms.Label
$Lbl_OtherOptionType = New-Object System.Windows.Forms.Label
$Lbl_OtherOptionValue = New-Object System.Windows.Forms.Label
$Txt_Option1Num = New-Object System.Windows.Forms.TextBox
$Cbo_Option1Type = New-Object System.Windows.Forms.ComboBox
$Txt_Option1Value = New-Object System.Windows.Forms.TextBox
$Txt_Option2Num = New-Object System.Windows.Forms.TextBox
$Cbo_Option2Type = New-Object System.Windows.Forms.ComboBox
$Txt_Option2Value = New-Object System.Windows.Forms.TextBox
$Txt_Option3Num = New-Object System.Windows.Forms.TextBox
$Cbo_Option3Type = New-Object System.Windows.Forms.ComboBox
$Txt_Option3Value = New-Object System.Windows.Forms.TextBox
$Lbl_OptionHelp = New-Object System.Windows.Forms.Label
$Btn_Create = New-Object System.Windows.Forms.Button
$Btn_Cancel = New-Object System.Windows.Forms.Button
$openFileDialog1 = New-Object System.Windows.Forms.OpenFileDialog

$Btn_Create_OnClick = { Get-ScopeInput }
$Btn_Cancel_OnClick = { $Form_CreateScope.Close() }
$Btn_Import_OnClick = { Import-ScopeFile }

$Form_CreateScope.Name = "Form_CreateScope"
$Form_CreateScope.Size = New-Object System.Drawing.Size(390,496)
$Form_CreateScope.Text = "Create DHCP Scope"
$Form_CreateScope.DataBindings.DefaultDataSourceUpdateMode = 0
$Form_CreateScope.MinimumSize = $Form_CreateScope.Size
$Form_CreateScope.MaximumSize = $Form_CreateScope.Size
$Form_CreateScope.CancelButton = $Btn_Cancel
$Form_CreateScope.add_Load({ Load-NewScopeForm })
$Form_CreateScope.add_FormClosed({ Close-NewScopeForm })

$Lbl_Servers.Name = "Lbl_Servers"
$Lbl_Servers.Text = "Check each server to create the scope on.  Select one server to activate the scope on."
$Lbl_Servers.TabIndex = 10
$Lbl_Servers.AutoSize = $false
$Lbl_Servers.Size = New-Object System.Drawing.Size(375,26)
$Lbl_Servers.Location = New-Object System.Drawing.Point(12,8)
$Lbl_Servers.DataBindings.DefaultDataSourceUpdateMode = 0
$Form_CreateScope.Controls.Add($Lbl_Servers)

$LstVw_Servers.Name = "LstVw_Servers"
$LstVw_Servers.UseCompatibleStateImageBehavior = $False
$LstVw_Servers.Size = New-Object System.Drawing.Size(203,61)
$LstVw_Servers.DataBindings.DefaultDataSourceUpdateMode = 0
$LstVw_Servers.View = 1
$LstVw_Servers.TabIndex = 0
$LstVw_Servers.Location = New-Object System.Drawing.Point(12,36)
$LstVw_Servers.Columns.Add($LstVw_Columns)|Out-Null
$LstVw_Servers.CheckBoxes = $True
$LstVw_Servers.HideSelection = $false
$LstVw_Servers.add_MouseDown({ Get-SelectedServers })
$LstVw_Servers.add_MouseUp({ Test-SelectedNotChecked })
$LstVw_Servers.add_KeyUp({ Test-SelectedNotChecked })
$LstVw_Servers.add_SelectedIndexChanged({ Check-SelectedServers })
$Form_CreateScope.Controls.Add($LstVw_Servers)

$LstVw_Columns.Text = "Server name"
$LstVw_Columns.Width = 157

$Btn_Import.Name = "Btn_Import"
$Btn_Import.Size = New-Object System.Drawing.Size(120,23)
$Btn_Import.Location = New-Object System.Drawing.Point(($LstVw_Servers.Right + 5),($LstVw_Servers.Top + $LstVw_Servers.Height  - $Btn_Import.Height))
$Btn_Import.Text = "Import from File"
$Btn_Import.TabIndex = 8
$Btn_Import.UseVisualStyleBackColor = $True
$Btn_Import.DataBindings.DefaultDataSourceUpdateMode = 0
$Btn_Import.add_Click($Btn_Import_OnClick)
$Form_CreateScope.Controls.Add($Btn_Import)

$openFileDialog1.Filter = "CSV files (*.csv)|*.csv|Text files (*.txt)|*.txt"
$openFileDialog1.Multiselect = $false
$openFileDialog1.InitialDirectory = $scriptfolder

$verticalspacingsmall = 26
$verticalspacinglarge = 35
$ipaddressfieldwidth = 90

$Lbl_Subnet.Name = "Lbl_Subnet"
$Lbl_Subnet.Size = New-Object System.Drawing.Size(38,13)
$Lbl_Subnet.Location = New-Object System.Drawing.Point(12,110)
$Lbl_Subnet.Text = "Subnet Name:"
$Lbl_Subnet.TabIndex = 11
$Lbl_Subnet.AutoSize = $True
$Lbl_Subnet.DataBindings.DefaultDataSourceUpdateMode = 0
$Form_CreateScope.Controls.Add($Lbl_Subnet)

$Txt_Subnet.Name = "Txt_Subnet"
$Txt_Subnet.Size = New-Object System.Drawing.Size($ipaddressfieldwidth,20)
$Txt_Subnet.Location = New-Object System.Drawing.Point(120,($Lbl_Subnet.Top - 3))
$Txt_Subnet.DataBindings.DefaultDataSourceUpdateMode = 0
$Txt_Subnet.TabIndex = 1
$Form_CreateScope.Controls.Add($Txt_Subnet)

$Lbl_SubnetHelp.Name = "Lbl_Subnet"
$Lbl_SubnetHelp.Size = New-Object System.Drawing.Size(38,13)
$Lbl_SubnetHelp.Location = New-Object System.Drawing.Point(($Txt_Subnet.Left + $Txt_Subnet.Width + 5),($Lbl_Subnet.Top))
$Lbl_SubnetHelp.Text = "Use IP Address format (x.x.x.x)"
$Lbl_SubnetHelp.Font = New-Object System.Drawing.Font("", 7)
$Lbl_SubnetHelp.TabIndex = 11
$Lbl_SubnetHelp.AutoSize = $True
$Lbl_SubnetHelp.DataBindings.DefaultDataSourceUpdateMode = 0
$Form_CreateScope.Controls.Add($Lbl_SubnetHelp)

$Lbl_Name.Name = "Lbl_Name"
$Lbl_Name.Size = New-Object System.Drawing.Size(38,13)
$Lbl_Name.Location = New-Object System.Drawing.Point(12,($Lbl_Subnet.Top + $verticalspacingsmall))
$Lbl_Name.Text = "Name (description):"
$Lbl_Name.TabIndex = 11
$Lbl_Name.AutoSize = $True
$Lbl_Name.DataBindings.DefaultDataSourceUpdateMode = 0
$Form_CreateScope.Controls.Add($Lbl_Name)

$Txt_Name.Name = "Txt_Name"
$Txt_Name.Size = New-Object System.Drawing.Size(220,20)
$Txt_Name.Location = New-Object System.Drawing.Point($Txt_Subnet.Left,($Txt_Subnet.Top + $verticalspacingsmall))
$Txt_Name.DataBindings.DefaultDataSourceUpdateMode = 0
$Txt_Name.TabIndex = 1
$Form_CreateScope.Controls.Add($Txt_Name)

$Lbl_Description.Name = "Lbl_Description"
$Lbl_Description.Size = New-Object System.Drawing.Size(63,13)
$Lbl_Description.Location = New-Object System.Drawing.Point(12,($Lbl_Name.Top + $verticalspacingsmall))
$Lbl_Description.Text = "Additional comment:"
$Lbl_Description.TabIndex = 12
$Lbl_Description.AutoSize = $True
$Lbl_Description.DataBindings.DefaultDataSourceUpdateMode = 0
$Form_CreateScope.Controls.Add($Lbl_Description)

$Txt_Description.Name = "Txt_Description"
$Txt_Description.Size = New-Object System.Drawing.Size(220,20)
$Txt_Description.Location = New-Object System.Drawing.Point($Txt_Subnet.Left,($Txt_Name.Top + $verticalspacingsmall))
$Txt_Description.DataBindings.DefaultDataSourceUpdateMode = 0
$Txt_Description.TabIndex = 2
$Form_CreateScope.Controls.Add($Txt_Description)

$Lbl_StartIP.Name = "Lbl_StartIP"
$Lbl_StartIP.Size = New-Object System.Drawing.Size(86,13)
$Lbl_StartIP.Location = New-Object System.Drawing.Point(12,($Lbl_Description.Top + $verticalspacinglarge))
$Lbl_StartIP.Text = "Start IP Address:"
$Lbl_StartIP.TabIndex = 13
$Lbl_StartIP.AutoSize = $True
$Lbl_StartIP.DataBindings.DefaultDataSourceUpdateMode = 0
$Form_CreateScope.Controls.Add($Lbl_StartIP)

$Txt_StartIP.Name = "Txt_StartIP"
$Txt_StartIP.Size = New-Object System.Drawing.Size($ipaddressfieldwidth,20)
$Txt_StartIP.Location = New-Object System.Drawing.Point($Txt_Subnet.Left,($Txt_Description.Top + $verticalspacinglarge))
$Txt_StartIP.DataBindings.DefaultDataSourceUpdateMode = 0
$Txt_StartIP.TabIndex = 3
$Form_CreateScope.Controls.Add($Txt_StartIP)

$Lbl_EndIP.Name = "Lbl_EndIP"
$Lbl_EndIP.Size = New-Object System.Drawing.Size(83,13)
$Lbl_EndIP.Location = New-Object System.Drawing.Point(12,($Lbl_StartIP.Top + $verticalspacingsmall))
$Lbl_EndIP.Text = "End IP Address:"
$Lbl_EndIP.TabIndex = 14
$Lbl_EndIP.AutoSize = $True
$Lbl_EndIP.DataBindings.DefaultDataSourceUpdateMode = 0
$Form_CreateScope.Controls.Add($Lbl_EndIP)

$Txt_EndIP.Name = "Txt_EndIP"
$Txt_EndIP.Size = New-Object System.Drawing.Size($ipaddressfieldwidth,20)
$Txt_EndIP.Location = New-Object System.Drawing.Point($Txt_Subnet.Left,($Txt_StartIP.Top + $verticalspacingsmall))
$Txt_EndIP.DataBindings.DefaultDataSourceUpdateMode = 0
$Txt_EndIP.TabIndex = 4
$Form_CreateScope.Controls.Add($Txt_EndIP)

$Lbl_SubnetLength.Name = "Lbl_SubnetLength"
$Lbl_SubnetLength.Size = New-Object System.Drawing.Size(76,13)
$Lbl_SubnetLength.Location = New-Object System.Drawing.Point(12,($Lbl_EndIP.Top + $verticalspacingsmall))
$Lbl_SubnetLength.Text = "Subnet mask length:"
$Lbl_SubnetLength.TabIndex = 15
$Lbl_SubnetLength.AutoSize = $True
$Lbl_SubnetLength.DataBindings.DefaultDataSourceUpdateMode = 0
$Form_CreateScope.Controls.Add($Lbl_SubnetLength)

$Num_SubnetLength.Name = "Num_SubnetLength"
$Num_SubnetLength.Size = New-Object System.Drawing.Size(50,20)
$Num_SubnetLength.Location = New-Object System.Drawing.Point($Txt_Subnet.Left,($Txt_EndIP.Top + $verticalspacingsmall))
$Num_SubnetLength.DataBindings.DefaultDataSourceUpdateMode = 0
$Num_SubnetLength.Maximum = 31
$Num_SubnetLength.Minimum = 1
$Num_SubnetLength.TabIndex = 5
$Num_SubnetLength.Value = 24
$Num_SubnetLength.add_ValueChanged({Set-SubnetMaskValue})
$Num_SubnetLength.add_LostFocus({Set-SubnetMaskValue})
$Form_CreateScope.Controls.Add($Num_SubnetLength)

$Lbl_SubnetIPValue.Name = "Lbl_SubnetIPValue"
$Lbl_SubnetIPValue.Size = New-Object System.Drawing.Size(121,13)
$Lbl_SubnetIPValue.Location = New-Object System.Drawing.Point(($Num_SubnetLength.Right + 10),$Lbl_SubnetLength.Top)
$Lbl_SubnetIPValue.Text = "Subnet mask:"
$Lbl_SubnetIPValue.TabIndex = 16
$Lbl_SubnetIPValue.AutoSize = $True
$Lbl_SubnetIPValue.DataBindings.DefaultDataSourceUpdateMode = 0
$Form_CreateScope.Controls.Add($Lbl_SubnetIPValue)

$Lbl_RouterIP.Name = "Lbl_RouterIP"
$Lbl_RouterIP.Size = New-Object System.Drawing.Size(55,13)
$Lbl_RouterIP.Location = New-Object System.Drawing.Point(12,($Lbl_SubnetLength.Top + $verticalspacingsmall))
$Lbl_RouterIP.Text = "Router IP:"
$Lbl_RouterIP.TabIndex = 17
$Lbl_RouterIP.AutoSize = $True
$Lbl_RouterIP.DataBindings.DefaultDataSourceUpdateMode = 0
$Form_CreateScope.Controls.Add($Lbl_RouterIP)

$Txt_RouterIP.Name = "Txt_RouterIP"
$Txt_RouterIP.Size = New-Object System.Drawing.Size($ipaddressfieldwidth,20)
$Txt_RouterIP.Location = New-Object System.Drawing.Point($Txt_Subnet.Left,($Num_SubnetLength.Top + $verticalspacingsmall))
$Txt_RouterIP.DataBindings.DefaultDataSourceUpdateMode = 0
$Txt_RouterIP.TabIndex = 6
$Form_CreateScope.Controls.Add($Txt_RouterIP)

$Grp_OtherOptions.Name = "Grp_OtherOptions"
$Grp_OtherOptions.Size = New-Object System.Drawing.Size(355,115)
$Grp_OtherOptions.Location = New-Object System.Drawing.Point(15,($Lbl_RouterIP.Top + 23))
$Grp_OtherOptions.Text = "Other scope options"
$Grp_OtherOptions.TabStop = $False
$Grp_OtherOptions.TabIndex = 18
$Grp_OtherOptions.DataBindings.DefaultDataSourceUpdateMode = 0
$Form_CreateScope.Controls.Add($Grp_OtherOptions)

$OptionFieldSpacing = 6

$Txt_Option1Num.Name = "Txt_Option1Num"
$Txt_Option1Num.Size = New-Object System.Drawing.Size(48,20)
$Txt_Option1Num.Location = New-Object System.Drawing.Point(6,32)
$Txt_Option1Num.DataBindings.DefaultDataSourceUpdateMode = 0
$Txt_Option1Num.TabIndex = 0
$Grp_OtherOptions.Controls.Add($Txt_Option1Num)

$Cbo_Option1Type.Name = "Cbo_Option1Type"
$Cbo_Option1Type.Size = New-Object System.Drawing.Size(90,$Txt_Option1Num.Top)
$Cbo_Option1Type.Location = New-Object System.Drawing.Point(($Txt_Option1Num.Right + $OptionFieldSpacing),$Txt_Option1Num.Top)
$Cbo_Option1Type.DataBindings.DefaultDataSourceUpdateMode = 0
$Cbo_Option1Type.TabIndex = 0
$Cbo_Option1Type.DropDownStyle = "DropDownList"
$Grp_OtherOptions.Controls.Add($Cbo_Option1Type)
$Cbo_Option1Type.Items.AddRange(@("STRING", "IPADDRESS", "BINARY"))
$Cbo_Option1Type.SelectedIndex = 0

$Txt_Option1Value.Name = "Txt_Option1Value"
$Txt_Option1Value.Size = New-Object System.Drawing.Size(185,$Txt_Option1Num.Top)
$Txt_Option1Value.Location = New-Object System.Drawing.Point(($Cbo_Option1Type.Right + $OptionFieldSpacing),$Cbo_Option1Type.Top)
$Txt_Option1Value.DataBindings.DefaultDataSourceUpdateMode = 0
$Txt_Option1Value.TabIndex = 1
$Grp_OtherOptions.Controls.Add($Txt_Option1Value)

$Txt_Option2Num.Name = "Txt_Option2Num"
$Txt_Option2Num.Size = $Txt_Option1Num.Size
$Txt_Option2Num.Location = New-Object System.Drawing.Point(6,58)
$Txt_Option2Num.DataBindings.DefaultDataSourceUpdateMode = 0
$Txt_Option2Num.TabIndex = 2
$Grp_OtherOptions.Controls.Add($Txt_Option2Num)

$Cbo_Option2Type.Name = "Cbo_Option2Type"
$Cbo_Option2Type.Size = $Cbo_Option1Type.Size
$Cbo_Option2Type.Location = New-Object System.Drawing.Point(($Txt_Option2Num.Right + $OptionFieldSpacing),$Txt_Option2Num.Top)
$Cbo_Option2Type.DataBindings.DefaultDataSourceUpdateMode = 0
$Cbo_Option2Type.TabIndex = 0
$Cbo_Option2Type.DropDownStyle = "DropDownList"
$Grp_OtherOptions.Controls.Add($Cbo_Option2Type)
$Cbo_Option2Type.Items.AddRange(@("STRING", "IPADDRESS", "BINARY"))
$Cbo_Option2Type.SelectedIndex = 0

$Txt_Option2Value.Name = "Txt_Option2Value"
$Txt_Option2Value.Size = $Txt_Option1Value.Size
$Txt_Option2Value.Location = New-Object System.Drawing.Point(($Cbo_Option2Type.Right + $OptionFieldSpacing),$Cbo_Option2Type.Top)
$Txt_Option2Value.DataBindings.DefaultDataSourceUpdateMode = 0
$Txt_Option2Value.TabIndex = 3
$Grp_OtherOptions.Controls.Add($Txt_Option2Value)

$Txt_Option3Num.Name = "Txt_Option3Num"
$Txt_Option3Num.Size = $Txt_Option1Num.Size
$Txt_Option3Num.Location = New-Object System.Drawing.Point(6,84)
$Txt_Option3Num.DataBindings.DefaultDataSourceUpdateMode = 0
$Txt_Option3Num.TabIndex = 4
$Grp_OtherOptions.Controls.Add($Txt_Option3Num)

$Cbo_Option3Type.Name = "Cbo_Option3Type"
$Cbo_Option3Type.Size = $Cbo_Option1Type.Size
$Cbo_Option3Type.Location = New-Object System.Drawing.Point(($Txt_Option3Num.Right + $OptionFieldSpacing),$Txt_Option3Num.Top)
$Cbo_Option3Type.DataBindings.DefaultDataSourceUpdateMode = 0
$Cbo_Option3Type.TabIndex = 0
$Cbo_Option3Type.DropDownStyle = "DropDownList"
$Grp_OtherOptions.Controls.Add($Cbo_Option3Type)
$Cbo_Option3Type.Items.AddRange(@("STRING", "IPADDRESS", "BINARY"))
$Cbo_Option3Type.SelectedIndex = 0

$Txt_Option3Value.Name = "Txt_Option3Value"
$Txt_Option3Value.Size = $Txt_Option1Value.Size
$Txt_Option3Value.Location = New-Object System.Drawing.Point(($Cbo_Option3Type.Right + $OptionFieldSpacing),$Cbo_Option3Type.Top)
$Txt_Option3Value.DataBindings.DefaultDataSourceUpdateMode = 0
$Txt_Option3Value.TabIndex = 5
$Grp_OtherOptions.Controls.Add($Txt_Option3Value)

$Lbl_OtherOptionNumber.Name = "Lbl_OtherScopeNumber"
$Lbl_OtherOptionNumber.Size = New-Object System.Drawing.Size($Txt_Option1Num.Width,13)
$Lbl_OtherOptionNumber.Location = New-Object System.Drawing.Point($Txt_Option1Num.Left,15)
$Lbl_OtherOptionNumber.Text = "Number"
$Lbl_OtherOptionNumber.TabIndex = 6
$Lbl_OtherOptionNumber.AutoSize = $True
$Lbl_OtherOptionNumber.DataBindings.DefaultDataSourceUpdateMode = 0
$Grp_OtherOptions.Controls.Add($Lbl_OtherOptionNumber)

$Lbl_OtherOptionType.Name = "Lbl_OtherOptionType"
$Lbl_OtherOptionType.Size = New-Object System.Drawing.Size($Cbo_Option1Type.Width,13)
$Lbl_OtherOptionType.Location = New-Object System.Drawing.Point($Cbo_Option1Type.Left,15)
$Lbl_OtherOptionType.Text = "Type"
$Lbl_OtherOptionType.TabIndex = 7
$Lbl_OtherOptionType.AutoSize = $True
$Lbl_OtherOptionType.DataBindings.DefaultDataSourceUpdateMode = 0
$Grp_OtherOptions.Controls.Add($Lbl_OtherOptionType)

$Lbl_OtherOptionValue.Name = "Lbl_OtherScopeValue"
$Lbl_OtherOptionValue.Size = New-Object System.Drawing.Size($Txt_Option1Value.Width,13)
$Lbl_OtherOptionValue.Location = New-Object System.Drawing.Point($Txt_Option1Value.Left,15)
$Lbl_OtherOptionValue.Text = "Value"
$Lbl_OtherOptionValue.TabIndex = 7
$Lbl_OtherOptionValue.AutoSize = $True
$Lbl_OtherOptionValue.DataBindings.DefaultDataSourceUpdateMode = 0
$Grp_OtherOptions.Controls.Add($Lbl_OtherOptionValue)

$Lbl_OptionHelp.Name = "Lbl_OptionHelp"
$Lbl_OptionHelp.AutoSize = $true
$Lbl_OptionHelp.Location = New-Object System.Drawing.Point($Grp_OtherOptions.Left, $Grp_OtherOptions.Bottom)
$Lbl_OptionHelp.Text = "Common scope options:`n006 - DNS Servers`n044 - WINS Servers`n(separate IP addresses with spaces)"
$Lbl_OptionHelp.Font = New-Object Drawing.Font("Microsoft Sans Serif", 9.0, [Drawing.GraphicsUnit]::pixel)
$Lbl_OptionHelp.TabIndex = 7
$Lbl_OptionHelp.AutoSize = $True
$Lbl_OptionHelp.DataBindings.DefaultDataSourceUpdateMode = 0
$Form_CreateScope.Controls.Add($Lbl_OptionHelp)

$Btn_Create.Name = "Btn_Create"
$Btn_Create.Size = New-Object System.Drawing.Size(75,23)
$Btn_Create.Location = New-Object System.Drawing.Point(214,428)
$Btn_Create.Text = "Create"
$Btn_Create.TabIndex = 8
$Btn_Create.UseVisualStyleBackColor = $True
$Btn_Create.DataBindings.DefaultDataSourceUpdateMode = 0
$Btn_Create.add_Click($Btn_Create_OnClick)
$Form_CreateScope.Controls.Add($Btn_Create)

$Btn_Cancel.Name = "Btn_Cancel"
$Btn_Cancel.Size = New-Object System.Drawing.Size(75,23)
$Btn_Cancel.Location = New-Object System.Drawing.Point(295,$Btn_Create.Top)
$Btn_Cancel.Text = "Cancel"
$Btn_Cancel.TabIndex = 9
$Btn_Cancel.UseVisualStyleBackColor = $True
$Btn_Cancel.DataBindings.DefaultDataSourceUpdateMode = 0
$Btn_Cancel.add_Click($Btn_Cancel_OnClick)
$Form_CreateScope.Controls.Add($Btn_Cancel)
$Form_CreateScope.CancelButton = $Btn_Cancel

########################################################################
# End Windows Form for Scope Creation.
########################################################################


#Open the main form
$Form_DHCP.ShowDialog() | Out-Null
exit




    #$allactions | ForEach-Object {
    #    [string]$logfiledate = "{0:yyyyMMddhhmmss}" -f [DateTime](Get-Date)
    #    [string]$logfiledescr = $_.Description.Replace("`"", "")
    #    if ($_.DeleteExclusion -or $_.CreateExclusion -or $_.DeleteReservation -or $_.CreateReservation) { $logfilename = "$logfolder\$logfiledate $($logfiledescr).txt" }
    #    $objStartInfo = New-Object System.Diagnostics.ProcessStartInfo
    #    $objStartInfo.FileName = "cmd.exe"
    #    $objStartInfo.WindowStyle = "Hidden"
    #    $objStartInfo.Arguments = "/C $($_.Command) >`"$logfilename`""
    #    [void][System.Diagnostics.Process]::Start($objStartInfo).WaitForExit()
    #    $logfileinfo = Get-Content $logfilename
    #    $success = $false
    #    $logfileinfo | Foreach-Object { if ($_ -match $commandsuccessful) { $success = $true } }
    #    $commandoutput = @()
    #    if ($success) {
    #        $commandoutput += ("Success: " + $_.Description)
    #    } else {
    #        $commandoutput += ("Failed: " + $_.Description)
    #        $commandoutput += ($logfileinfo | ? {$_ -ne ""} )
    #    }
    #    $commandoutput += ""
    #    $Txt_Results.Text += [string]::Join("`n", $commandoutput).Replace("`n", "`r`n"); $Txt_Results.Refresh()
    #    if (Test-Path $_.TempFile) { Remove-Item $_.TempFile }
    #}
