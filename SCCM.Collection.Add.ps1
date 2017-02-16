<#
Add a computer to multiple SCCM collections
Author:Fabian Wilson
Date:1/3/2017

#>
$server = "SCCMServer"
$FolderID = "Put the folder ID here" # Use the WMI tool to get the FolderID
$site= "Site goes here"

Import-Module -Name "$(split-path $Env:SMS_ADMIN_UI_PATH)\ConfigurationManager.psd1"
Write-Host "Importing SCCM Module"
Write-Host "Connecting to PS-Drive PS1:\"
Set-Location PS1:\

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
$form = New-Object System.Windows.Forms.Form 
$form.Text = "Add a Computer to SCCM Collections"
$form.Size = New-Object System.Drawing.Size(1300,800) 
$form.StartPosition = "CenterScreen"

$form.KeyPreview = $True
$form.Add_KeyDown({if ($_.KeyCode -eq "Enter") 
    {$x=$TextBox.Text;$form.Close()}})
$form.Add_KeyDown({if ($_.KeyCode -eq "Escape") 
    {$form.Close()}})


$statusBar = New-Object -TypeName System.Windows.Forms.StatusBar
$statusBar.Text = ‘Hold the CTRL key to select multiple items’
$form.Controls.Add($statusBar)

$label = New-Object System.Windows.Forms.Label
$Label.Location = New-Object System.Drawing.Size(40,80) 
$Label.Size = New-Object System.Drawing.Size(380,20) 
$label.Text = "Please make a selection from the list below:"
$form.Controls.Add($label) 

$listBox = New-Object System.Windows.Forms.Listbox 
$listBox.Location = New-Object System.Drawing.Point(40,100)
$listBox.Size = New-Object System.Drawing.Size(600,600) 
$listBox.SelectionMode = "MultiExtended"
$form.Controls.Add($listBox) 

$outputBox = New-Object System.Windows.Forms.TextBox 
$outputBox.Location = New-Object System.Drawing.Size(660,100) 
$outputBox.Size = New-Object System.Drawing.Size(600,590) 
$outputBox.ForeColor = "LimeGreen"
$outputBox.BackColor = "Black"
$outputBox.MultiLine = $True 
$outputBox.font = "Impact Regular,8"
$outputBox.ScrollBars = "Vertical" 
$form.Controls.Add($outputBox) 

#Get the list of collections from SCCM folder and loads them into the list box.
$CollectionsInSpecficFolder = Get-WmiObject -ComputerName $server -Namespace "ROOT\SMS\Site_$Site" `
-Query "select * from SMS_Collection where CollectionID is in(select InstanceKey from SMS_ObjectContainerItem where ObjectType='5000' and ContainerNodeID='$FolderID') and CollectionType='2'"
$CollectionsInSpecficFolder.Name|Sort| ForEach-Object {[void] $listBox.Items.Add($_)}

Function closeButton_OnClick  
{ 
    $form.Close() 
} 

$form.Topmost = $True
$Label = New-Object System.Windows.Forms.Label
$Label.Location = New-Object System.Drawing.Size(40,20) 
$Label.Size = New-Object System.Drawing.Size(280,20) 
$Label.Text = "Please enter the computername in the space below:"
$form.Controls.Add($Label) 

$TextBox = New-Object System.Windows.Forms.TextBox 
$TextBox.Location = New-Object System.Drawing.Size(40,40) 
$TextBox.Size = New-Object System.Drawing.Size(260,20) 
$TextBox.add_TextChanged({ $OKButton.Enabled = $true })
$x = @()
$y = @()
$form.Controls.Add($TextBox) 
$form.Add_Shown({$form.Activate()})

    Function AddSoftware{
   	$y = $listBox.SelectedItems
	$x=$TextBox.Text
    foreach($i in $y){ 
         Add-CMDeviceCollectionDirectMembershipRule  -CollectionName $i -ResourceId $(get-cmdevice -Name $x).ResourceID -ErrorAction Inquire  
		 }
     
    $outputBox.AppendText($x +" "+"has been added to the"+" "+ $i +" "+"collection" + "`n" + "`n");
	}
     
$OKButton = New-Object System.Windows.Forms.Button
$OKButton.Location = New-Object System.Drawing.Size(40,700)
$OKButton.Size = New-Object System.Drawing.Size(75,23)
$OKButton.Text = "OK"
$OKButton.Enabled = $false
$OKButton.Add_Click({AddSoftware}) 
$form.AcceptButton = $OKButton
$form.Controls.Add($OKButton)


$CancelButton = New-Object System.Windows.Forms.Button
$CancelButton.Location = New-Object System.Drawing.Size(150,700)
$CancelButton.Size = New-Object System.Drawing.Size(75,23)
$CancelButton.Text = "Cancel"
$CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$form.CancelButton = $CancelButton
$form.Controls.Add($CancelButton)

$CloseButton = New-Object System.Windows.Forms.Button
$CloseButton.Location = New-Object System.Drawing.Size(250,700)
$CloseButton.Size = New-Object System.Drawing.Size(75,23)
$CloseButton.Text = "Close"
$CloseButton.Add_Click({CloseButton_OnClick}) 
$form.AcceptButton = $CloseButton
$form.Controls.Add($CloseButton)


$form.Add_Shown({$form.Activate()})
$form.ShowDialog()
# SIG # Begin signature block
# MIIEBQYJKoZIhvcNAQcCoIID9jCCA/ICAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUl/3XK0lfnVGlsh9toSKEVgoq
# 1l+gggIfMIICGzCCAYigAwIBAgIQS1x2DIz7V5BNoqLcf8tV0DAJBgUrDgMCHQUA
# MBwxGjAYBgNVBAMTEVNDQ00uU29mdHdhcmUuQWRkMB4XDTE3MDExMzE0MzYxNloX
# DTM5MTIzMTIzNTk1OVowHDEaMBgGA1UEAxMRU0NDTS5Tb2Z0d2FyZS5BZGQwgZ8w
# DQYJKoZIhvcNAQEBBQADgY0AMIGJAoGBAKBupyPYbD8DOLSGJwhGDkQ2MPKzuO+G
# FZw0X5egm70RNtMxaMXJrfr7cgvRWSy34GGZ8lqpHVtDI84fh6f17Gs1LYJDM+MM
# yGQ04bXCgweY865nFlgexkdo9K6jh0FCbO+JIKsYRkImO7aUYmCuDdOJijr3bNyF
# GejUkZn24t9NAgMBAAGjZjBkMBMGA1UdJQQMMAoGCCsGAQUFBwMDME0GA1UdAQRG
# MESAECXgMkwtBmW4G+nzTf3vuWWhHjAcMRowGAYDVQQDExFTQ0NNLlNvZnR3YXJl
# LkFkZIIQS1x2DIz7V5BNoqLcf8tV0DAJBgUrDgMCHQUAA4GBAHvHphpynkERVaP5
# cemCN/J01tvfzNQvtqx0Hu8QgsibguGUHJayjdbaK+Hy8T8yQBaJ/MMCG04sxc8L
# tSR0hRrkEuJPw5NoPR5fWQ8kcBYnsY5N81R9yLzcaZKsuXaUCjr1T9yDPnFJK+DA
# KylSlkxm+siAs683xzX/XTWnWebvMYIBUDCCAUwCAQEwMDAcMRowGAYDVQQDExFT
# Q0NNLlNvZnR3YXJlLkFkZAIQS1x2DIz7V5BNoqLcf8tV0DAJBgUrDgMCGgUAoHgw
# GAYKKwYBBAGCNwIBDDEKMAigAoAAoQKAADAZBgkqhkiG9w0BCQMxDAYKKwYBBAGC
# NwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQx
# FgQU+TEeKillPIcMpQUABXgHAMOn8pQwDQYJKoZIhvcNAQEBBQAEgYBrxvcLoKrR
# wZiRdWe+XBYZxPA+AFmZXAeuHJmRTOASqZT/kq3l+nnZeU+SaS886veMWGl+CY0s
# k0PkXPzirwcnoEmmz2+3Ns1Pp9K+GAyYiwf9p7fn1HFmtb2xkkrM/SL0EO93dFyG
# AzuQ189dsaEtPIo/NlDSzcEx9WOzvxMl0A==
# SIG # End signature block