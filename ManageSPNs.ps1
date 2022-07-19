﻿#[void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null - Commenting out "void", because if we keep both "void and Out-Null in single line, it's preventing script to execute in newer version of powershell.
[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null

####################################################################################
# This code will minimize the Powershell console window.
if (-not $showWindowAsync) { $showWindowAsync = Add-Type –memberDefinition '[DllImport("user32.dll")]public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);' -name “Win32ShowWindowAsync” -namespace Win32Functions –passThru }; function Show-PowerShell() { $null = $showWindowAsync::ShowWindowAsync((Get-Process –id $pid).MainWindowHandle, 10) }; function Hide-PowerShell() { $null = $showWindowAsync::ShowWindowAsync((Get-Process –id $pid).MainWindowHandle, 2) }; Hide-PowerShell
####################################################################################



[DirectoryServices.DirectoryEntry]$global:objUser = $nothing
$global:allSPNs = @()
$global:username = ""

$formspacingv = 23
$formspacingx = 76
$formspacingindent = 16

#Generated Form Function
function GenerateForm {
########################################################################
# Code Generated By: SAPIEN Technologies PrimalForms (Community Edition) v1.0.8.0
# Generated On: 4/18/2011 3:20 PM
# Generated By: d-jacobje
########################################################################

#region Import the Assemblies
[reflection.assembly]::loadwithpartialname("System.Windows.Forms") | Out-Null
[reflection.assembly]::loadwithpartialname("System.Drawing") | Out-Null
#endregion

#region Generated Form Objects
$Form_Main = New-Object System.Windows.Forms.Form
$Lbl_AccountID = New-Object System.Windows.Forms.Label
$Txt_AccountID = New-Object System.Windows.Forms.TextBox
$Btn_GetSPNs = New-Object System.Windows.Forms.Button
$List_SPNs = New-Object System.Windows.Forms.ListBox
$ContextMenuStrip1 = New-Object System.Windows.Forms.ContextMenuStrip
$ToolStripMenuItem1 = New-Object System.Windows.Forms.ToolStripMenuItem
$Rdo_Format1 = New-Object System.Windows.Forms.RadioButton
$Lbl_Service = New-Object System.Windows.Forms.Label
$Txt_Service = New-Object System.Windows.Forms.TextBox
$Lbl_HostName = New-Object System.Windows.Forms.Label
$Txt_HostName = New-Object System.Windows.Forms.TextBox
$Lbl_Port = New-Object System.Windows.Forms.Label
$Txt_Port = New-Object System.Windows.Forms.TextBox
$Rdo_Format2 = New-Object System.Windows.Forms.RadioButton
$Txt_SPNString = New-Object System.Windows.Forms.TextBox
$Lbl_ServerDN = New-Object System.Windows.Forms.Label
$Txt_ServerDN = New-Object System.Windows.Forms.TextBox
$Btn_Add = New-Object System.Windows.Forms.Button
$Btn_Delete = New-Object System.Windows.Forms.Button
$Btn_Close = New-Object System.Windows.Forms.Button
$InitialFormWindowState = New-Object System.Windows.Forms.FormWindowState
$Lbl_Filter = New-Object System.Windows.Forms.Label
$Txt_Filter = New-Object System.Windows.Forms.TextBox
$Btn_ClearFilter = New-Object System.Windows.Forms.Button
#endregion Generated Form Objects

#----------------------------------------------
#Generated Event Script Blocks
#----------------------------------------------
#Provide Custom Code for events specified in PrimalForms.
$Btn_GetSPNs_OnClick= { $global:username = $Txt_AccountID.Text; Get-SPNs }
$Btn_Add_OnClick= { Add-SPN }
$Btn_Delete_OnClick= { Delete-SPN }
$Btn_Close_OnClick= { $Form_Main.Close() }
$Btn_ClearFilter_OnClick = { $Txt_Filter.Text = "" }
$Btn_ShowServerDN_OnClick = { Show-ServerDN }

$OnLoadForm_StateCorrection=
{#Correct the initial state of the form to prevent the .Net maximized form issue
	$Form_Main.WindowState = $InitialFormWindowState
    Set-RadioOptions
}

#----------------------------------------------
#region Generated Form Code
$Form_Main.Name = "Form_Main"
$Form_Main.Text = "Manage SPNs"
$Form_Main.DataBindings.DefaultDataSourceUpdateMode = 0
$Form_Main.Size = New-Object System.Drawing.Size(417,548)
$Form_Main.MinimumSize = $Form_Main.Size

$Lbl_AccountID.Name = "Lbl_AccountID"
$Lbl_AccountID.Text = "Account ID:"
$Lbl_AccountID.Size = New-Object System.Drawing.Size(64,13)
$Lbl_AccountID.Location = New-Object System.Drawing.Point(12,15)
$Lbl_AccountID.TabIndex = 2
$Lbl_AccountID.AutoSize = $True
$Lbl_AccountID.DataBindings.DefaultDataSourceUpdateMode = 0
$Form_Main.Controls.Add($Lbl_AccountID)

$Txt_AccountID.Name = "Txt_AccountID"
$Txt_AccountID.Size = New-Object System.Drawing.Size(234,20)
$Txt_AccountID.Location = New-Object System.Drawing.Point(82,12)
$Txt_AccountID.DataBindings.DefaultDataSourceUpdateMode = 0
$Txt_AccountID.Anchor = 13
$Txt_AccountID.TabIndex = 0
$Txt_AccountID.Text = ""
$Form_Main.Controls.Add($Txt_AccountID)

$Btn_GetSPNs.Name = "Btn_GetSPNs"
$Btn_GetSPNs.Text = "Get SPNs"
$Btn_GetSPNs.Size = New-Object System.Drawing.Size(75,23)
$Btn_GetSPNs.Location = New-Object System.Drawing.Point(322,10)
$Btn_GetSPNs.TabIndex = 1
$Btn_GetSPNs.Anchor = 9
$Btn_GetSPNs.UseVisualStyleBackColor = $True
$Btn_GetSPNs.DataBindings.DefaultDataSourceUpdateMode = 0
$Btn_GetSPNs.add_Click($Btn_GetSPNs_OnClick)
$Form_Main.Controls.Add($Btn_GetSPNs)
$Form_Main.AcceptButton = $Btn_GetSPNs

$List_SPNs.Name = "List_SPNs"
$List_SPNs.Size = New-Object System.Drawing.Size(385,240)
$List_SPNs.Location = New-Object System.Drawing.Point(12,40)
$List_SPNs.FormattingEnabled = $True
$List_SPNs.DataBindings.DefaultDataSourceUpdateMode = 0
$List_SPNs.TabIndex = 14
$List_SPNs.Anchor = 15
$List_SPNs.add_Click({ Fill-Form })
$List_SPNs.add_SelectedIndexChanged({ Fill-Form })
$List_SPNs.SelectionMode = "MultiExtended"
$List_SPNs.Sorted = $true
$List_SPNs.ContextMenuStrip = $ContextMenuStrip1
$Form_Main.Controls.Add($List_SPNs)

$ContextMenuStrip1.Items.Add($ToolStripMenuItem1) | Out-Null
$ContextMenuStrip1.Name = "ContextMenuStrip1"
$ContextMenuStrip1.Size = New-Object System.Drawing.Size(182, 48)

$ToolStripMenuItem1.Name = "ToolStripMenuItem1"
$ToolStripMenuItem1.Size = New-Object System.Drawing.Size(181, 22)
$ToolStripMenuItem1.Text = "Copy"
$ToolStripMenuItem1.add_Click({ Copy-SelectedItems })

$Lbl_Filter.Name = "Lbl_Filter"
$Lbl_Filter.Text = "Filter:"
$Lbl_Filter.Size = New-Object System.Drawing.Size(64,13)
$Lbl_Filter.Location = New-Object System.Drawing.Point(($List_SPNs.Right - 182), ($List_SPNs.Bottom + 8))
$Lbl_Filter.TabIndex = 2
$Lbl_Filter.AutoSize = $True
$Lbl_Filter.DataBindings.DefaultDataSourceUpdateMode = 0
$Lbl_Filter.Anchor = 10
$Form_Main.Controls.Add($Lbl_Filter)

$Txt_Filter.Name = "Txt_Filter"
$Txt_Filter.Size = New-Object System.Drawing.Size(120,20)
$Txt_Filter.Location = New-Object System.Drawing.Point(($Lbl_Filter.Right + 4),($Lbl_Filter.Top - 4))
$Txt_Filter.DataBindings.DefaultDataSourceUpdateMode = 0
$Txt_Filter.Anchor = 13
$Txt_Filter.TabIndex = 0
$Txt_Filter.Text = ""
$Txt_Filter.Anchor = 10
$Txt_Filter.add_TextChanged({ Fill-List })
$Form_Main.Controls.Add($Txt_Filter)

$Btn_ClearFilter.Name = "Btn_ClearFilter"
$Btn_ClearFilter.Text = "X"
$Btn_ClearFilter.Size = New-Object System.Drawing.Size(20,20)
$Btn_ClearFilter.Location = New-Object System.Drawing.Point(($Txt_Filter.Right + 4),($Txt_Filter.Top))
$Btn_ClearFilter.TabIndex = 1
$Btn_ClearFilter.Anchor = 9
$Btn_ClearFilter.UseVisualStyleBackColor = $True
$Btn_ClearFilter.DataBindings.DefaultDataSourceUpdateMode = 0
$Btn_ClearFilter.add_Click($Btn_ClearFilter_OnClick)
$Btn_ClearFilter.Anchor = 10
$Form_Main.Controls.Add($Btn_ClearFilter)

$Rdo_Format1.Name = "Rdo_Format1"
$Rdo_Format1.Location = New-Object System.Drawing.Point($List_SPNs.Left,290)
$Rdo_Format1.Size = New-Object System.Drawing.Size(104,24)
$Rdo_Format1.UseVisualStyleBackColor = $True
$Rdo_Format1.Text = "Enter using separate fields:"
$Rdo_Format1.DataBindings.DefaultDataSourceUpdateMode = 0
$Rdo_Format1.TabStop = $True
$Rdo_Format1.AutoSize = $True
$Rdo_Format1.TabIndex = 16
$Rdo_Format1.Checked = $true
$Rdo_Format1.Anchor = 6
$Rdo_Format1.add_Click({ Set-RadioOptions })
$Form_Main.Controls.Add($Rdo_Format1)

$Lbl_Service.Name = "Lbl_Service"
$Lbl_Service.Text = "Service:"
$Lbl_Service.Size = New-Object System.Drawing.Size(49,13)
$Lbl_Service.Location = New-Object System.Drawing.Point(($Rdo_Format1.Left + $formspacingindent),($Rdo_Format1.Top + $formspacingv + 5))
$Lbl_Service.TabIndex = 7
$Lbl_Service.Anchor = 6
$Lbl_Service.AutoSize = $True
$Lbl_Service.DataBindings.DefaultDataSourceUpdateMode = 0
$Form_Main.Controls.Add($Lbl_Service)

$Txt_Service.Name = "Txt_Service"
$Txt_Service.Size = New-Object System.Drawing.Size(225,20)
$Txt_Service.Location = New-Object System.Drawing.Point(($Lbl_Service.Left + $formspacingx),($Lbl_Service.Top - 3))
$Txt_Service.DataBindings.DefaultDataSourceUpdateMode = 0
$Txt_Service.Anchor = 14
$Txt_Service.TabIndex = 4
$Txt_Service.add_LostFocus({ Update-Fields 1 })
$Form_Main.Controls.Add($Txt_Service)

$Lbl_HostName.Name = "Lbl_HostName"
$Lbl_HostName.Text = "Host name:"
$Lbl_HostName.Size = New-Object System.Drawing.Size(70,13)
$Lbl_HostName.Location = New-Object System.Drawing.Point($Lbl_Service.Left,($Lbl_Service.Top + $formspacingv))
$Lbl_HostName.TabIndex = 8
$Lbl_HostName.Anchor = 6
$Lbl_HostName.AutoSize = $True
$Lbl_HostName.DataBindings.DefaultDataSourceUpdateMode = 0
$Form_Main.Controls.Add($Lbl_HostName)

$Txt_HostName.Name = "Txt_HostName"
$Txt_HostName.Size = New-Object System.Drawing.Size(225,20)
$Txt_HostName.Location = New-Object System.Drawing.Point(($Lbl_HostName.Left + $formspacingx),($Txt_Service.Top + $formspacingv))
$Txt_HostName.DataBindings.DefaultDataSourceUpdateMode = 0
$Txt_HostName.Anchor = 14
$Txt_HostName.TabIndex = 5
$Txt_HostName.add_LostFocus({ Update-Fields 1 })
$Form_Main.Controls.Add($Txt_HostName)

$Lbl_Port.Name = "Lbl_Port"
$Lbl_Port.Text = "Port:"
$Lbl_Port.Size = New-Object System.Drawing.Size(67,13)
$Lbl_Port.Location = New-Object System.Drawing.Point($Lbl_Service.Left,($Lbl_HostName.Top + $formspacingv))
$Lbl_Port.TabIndex = 9
$Lbl_Port.Anchor = 6
$Lbl_Port.AutoSize = $True
$Lbl_Port.DataBindings.DefaultDataSourceUpdateMode = 0
$Form_Main.Controls.Add($Lbl_Port)

$Txt_Port.Name = "Txt_Port"
$Txt_Port.Size = New-Object System.Drawing.Size(225,20)
$Txt_Port.Location = New-Object System.Drawing.Point(($Lbl_Port.Left + $formspacingx),($Txt_HostName.Top + $formspacingv))
$Txt_Port.DataBindings.DefaultDataSourceUpdateMode = 0
$Txt_Port.Anchor = 14
$Txt_Port.TabIndex = 6
$Txt_Port.add_LostFocus({ Update-Fields 1 })
$Form_Main.Controls.Add($Txt_Port)

$Rdo_Format2.TabIndex = 16
$Rdo_Format2.Name = "Rdo_Format2"
$Rdo_Format2.Size = New-Object System.Drawing.Size(104,24)
$Rdo_Format2.UseVisualStyleBackColor = $True
$Rdo_Format2.Text = "Enter as text string:"
$Rdo_Format2.Location = New-Object System.Drawing.Point($List_SPNs.Left,390)
$Rdo_Format2.DataBindings.DefaultDataSourceUpdateMode = 0
$Rdo_Format2.TabStop = $True
$Rdo_Format2.AutoSize = $True
$Rdo_Format2.Anchor = 6
$Rdo_Format2.add_Click({ Set-RadioOptions })
$Form_Main.Controls.Add($Rdo_Format2)

$Txt_SPNString.Name = "Txt_SPNString"
$Txt_SPNString.Size = New-Object System.Drawing.Size(300,20)
$Txt_SPNString.Location = New-Object System.Drawing.Point(($Rdo_Format2.Left + $formspacingindent),($Rdo_Format2.Top + $formspacingv))
$Txt_SPNString.DataBindings.DefaultDataSourceUpdateMode = 0
$Txt_SPNString.Anchor = 13
$Txt_SPNString.TabIndex = 18
$Txt_SPNString.Anchor = 14
$Txt_SPNString.add_LostFocus({ Update-Fields 2 })
$Form_Main.Controls.Add($Txt_SPNString)

$Lbl_ServerDN.Name = "Lbl_ServerDN"
$Lbl_ServerDN.Text = "Server DN:"
$Lbl_ServerDN.Size = New-Object System.Drawing.Size(67,13)
$Lbl_ServerDN.Location = New-Object System.Drawing.Point($Rdo_Format2.Left,($Txt_SPNString.Top + $formspacingv * 1.5))
$Lbl_ServerDN.TabIndex = 9
$Lbl_ServerDN.Anchor = 6
$Lbl_ServerDN.AutoSize = $True
$Lbl_ServerDN.DataBindings.DefaultDataSourceUpdateMode = 0
$Form_Main.Controls.Add($Lbl_ServerDN)

$Txt_ServerDN.Name = "Txt_ServerDN"
$Txt_ServerDN.Size = New-Object System.Drawing.Size(325,20)
$Txt_ServerDN.Location = New-Object System.Drawing.Point(($Lbl_ServerDN.Left + 60),($Lbl_ServerDN.Top - 4))
$Txt_ServerDN.DataBindings.DefaultDataSourceUpdateMode = 0
$Txt_ServerDN.ReadOnly = $true
$Txt_ServerDN.Anchor = 14
$Txt_ServerDN.TabIndex = 6
$Form_Main.Controls.Add($Txt_ServerDN)

$Btn_Add.Name = "Btn_Add"
$Btn_Add.Text = "Add"
$Btn_Add.Size = New-Object System.Drawing.Size(75,23)
$Btn_Add.Location = New-Object System.Drawing.Point(12,483)
$Btn_Add.TabIndex = 10
$Btn_Add.Anchor = 6
$Btn_Add.UseVisualStyleBackColor = $True
$Btn_Add.DataBindings.DefaultDataSourceUpdateMode = 0
$Btn_Add.add_Click($Btn_Add_OnClick)
$Form_Main.Controls.Add($Btn_Add)

$Btn_Delete.Name = "Btn_Delete"
$Btn_Delete.Text = "Delete"
$Btn_Delete.Size = New-Object System.Drawing.Size(75,23)
$Btn_Delete.Location = New-Object System.Drawing.Point(($Btn_Add.Right + 10),$Btn_Add.Top)
$Btn_Delete.TabIndex = 12
$Btn_Delete.Anchor = 6
$Btn_Delete.UseVisualStyleBackColor = $True
$Btn_Delete.DataBindings.DefaultDataSourceUpdateMode = 0
$Btn_Delete.add_Click($Btn_Delete_OnClick)
$Form_Main.Controls.Add($Btn_Delete)

$Btn_Close.Name = "Btn_Close"
$Btn_Close.Text = "Close"
$Btn_Close.Size = New-Object System.Drawing.Size(75,23)
$Btn_Close.Location = New-Object System.Drawing.Point(322,$Btn_Add.Top)
$Btn_Close.TabIndex = 13
$Btn_Close.Anchor = 10
$Btn_Close.UseVisualStyleBackColor = $True
$Btn_Close.DataBindings.DefaultDataSourceUpdateMode = 0
$Btn_Close.add_Click($Btn_Close_OnClick)
$Form_Main.Controls.Add($Btn_Close)
$Form_Main.CancelButton = $Btn_Close

#endregion Generated Form Code

#Save the initial state of the form
$InitialFormWindowState = $Form_Main.WindowState
#Init the OnLoad event to correct the initial state of the form
$Form_Main.add_Load($OnLoadForm_StateCorrection)
#Show the Form
$Form_Main.ShowDialog()| Out-Null

} #End Function


function Set-RadioOptions {
    if ($Rdo_Format1.Checked) {
        $formfieldsenabled = $true
        $SPNstringenabled = $false
    } else {
        $formfieldsenabled = $false
        $SPNstringenabled = $true
    }
    $Lbl_Service.Enabled = $formfieldsenabled
    $Txt_Service.Enabled = $formfieldsenabled
    $Lbl_HostName.Enabled = $formfieldsenabled
    $Txt_HostName.Enabled = $formfieldsenabled
    $Lbl_Port.Enabled = $formfieldsenabled
    $Txt_Port.Enabled = $formfieldsenabled
    $Txt_SPNString.Enabled = $SPNstringenabled
}

function Get-SPNs {
    $Txt_AccountID.Text = $global:username.Trim()
    if ($global:username -eq "") { return }
    $search = New-Object DirectoryServices.DirectorySearcher([ADSI]("LDAP://DC=corp,DC=amvescap,DC=net"), "(&(objectclass=user)(objectcategory=person)(samaccountname=$($global:username)))", @("samaccountname"))
    $result = $search.FindOne()
    if ($result) {
        $global:objUser = $result.GetDirectoryEntry()
        $global:allSPNs = @()
        $global:objUser.servicePrincipalName | foreach { $global:allSPNs += $_ }
        Fill-List
    } else {
        $global:objUser = $nothing
        [Windows.Forms.Messagebox]::Show("Account ID not found.")
    }
}

function Fill-List {
param($SPNs)
    $List_SPNs.Items.Clear()
    Clear-Form
    $global:allSPNs | foreach {
        $additem = ($Txt_Filter.Text -eq "")
        if (-not $additem) { $additem = $_.ToLower().Contains($Txt_Filter.Text.ToLower()) }
        if ($additem) { $List_SPNs.Items.Add($_) }
    }
}

function Fill-Form {
    if ($List_SPNs.SelectedItems.Count -eq 1) {
        $selecteditem = $List_SPNs.SelectedItems[0]
        Split-Fields $selecteditem
        $Txt_SPNString.Text = $selecteditem
        Update-ServerDN
    } else {
        Clear-Form
    }
}

function Clear-Form {
    $Txt_Service.Text = ""
    $Txt_HostName.Text = ""
    $Txt_Port.Text = ""
    $Txt_SPNString.Text = ""
    $Txt_ServerDN.Text = ""
}

function Split-Fields {
param($spnvalue)
    $Txt_Service.Text = $spnvalue.Split("/")[0]
    $Txt_HostName.Text = $spnvalue.Replace("$($Txt_Service.Text)/","").Split(":")[0]
    $Txt_Port.Text = ""
    if ($spnvalue.Replace("$service/","").Split(":").Count -gt 1) { $Txt_Port.Text = $spnvalue.Replace("$service/","").Split(":")[1] }
}

function Join-Fields {
    $newspnstring = "$($Txt_Service.Text.Trim())/$($Txt_HostName.Text.Trim())"
    if ($newspnstring -eq "/") { $newspnstring = "" }
    if ($Txt_Port.Text) { $newspnstring += ":$($Txt_Port.Text.Trim())" }
    $Txt_SPNString.Text = $newspnstring
}

function Get-ServerDN {
    $servername = $Txt_HostName.Text.Split(".")[0]
    $search = New-Object DirectoryServices.DirectorySearcher([ADSI]("LDAP://DC=corp,DC=amvescap,DC=net"), "(&(objectclass=computer)(objectcategory=computer)(cn=$servername))", @("cn", "distinguishedname"))
    $result = $search.FindOne()
    if ($result) { $serverdn = $result.properties["distinguishedname"][0] } else { $serverdn = "Server not found." }
    $serverdn
}

function Update-ServerDN {
    $Txt_ServerDN.Text = ""
    if ($Txt_HostName.Text -ne "") { $Txt_ServerDN.Text = Get-ServerDN $Txt_HostName.Text }
}

function Delete-SPN {
    if ($global:objUser) {
        $removeSPNs = @()
        $List_SPNs.SelectedItems | foreach { $removeSPNs += $_.ToString() }
        $removeSPNs | foreach { $global:objUser.servicePrincipalName.Remove($_) }
        $global:objUser.SetInfo()
        Get-SPNs
    }
}

function Add-SPN {
    if ($global:objUser) {
        if ($Rdo_Format1.Checked) {
            $newSPN = "$($Txt_Service.Text.Trim())/$($Txt_HostName.Text.Trim())"
            if ($Txt_Port.Text) { $newSPN += ":$($Txt_Port.Text.Trim())" }
        } else {
            $newSPN = $Txt_SPNString.Text.Trim()
        }
        if ($List_SPNs.Items -contains $newSPN) {
            [Windows.Forms.Messagebox]::Show("This SPN is already in the list.")
        } else {
            $global:objUser.servicePrincipalName.Add($newSPN)
            $global:objUser.SetInfo()
            Get-SPNs
            $List_SPNs.TopIndex = $List_SPNs.FindStringExact($newSPN)
        }
    }
}

function Update-Fields {
param($mode)
    if ($mode -eq 1) {
        Join-Fields
    } else {
        Split-Fields $Txt_SPNString.Text
    }
    Update-ServerDN
}

function Copy-SelectedItems {
    $selecteditems = @()
    $List_SPNs.SelectedItems | foreach { $selecteditems += $_.ToString() }
    if ($selecteditems) {
        [Windows.Forms.Clipboard]::SetText([string]::Join("`r",$selecteditems).Replace("`r", "`r`n"))
    }
}

#Call the Function
GenerateForm
