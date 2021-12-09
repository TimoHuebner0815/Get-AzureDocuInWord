# Connect to Azure
IF (Get-InstalledModule -Name AZ) {
    
    Connect-AzAccount

} else {

    Install-Module AZ -Force
    Import-Module AZ

    Connect-AzAccount
}

# Creating the Word Object, set Word to visual and add document
$Word = New-Object -ComObject Word.Application
$Word.Visible = $True
$Document = $Word.Documents.Add("$PSScriptRoot\Azure_Documentation_Template.docx")
$Selection = $Word.Selection

## Go to End of Document to start
$Selection.EndKey(6,0)

###
### VIRTUAL MACHINES
###

## Add some text
$Selection.Style = 'Überschrift 1'
$Selection.TypeText("Virtual Machines")
$Selection.TypeParagraph()

## Get all VMs from Azure
$VMs = Get-AzVM

## Add a table for VMs
$TableBehavior = [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior

# Tables.Add(Range, RowCoun, ColumnsCount, Behavior)
$VMTable = $Selection.Tables.add($Word.Selection.Range, $VMs.Count + 2, 5, $TableBehavior)

$VMTable.Style = "Default"
$VMTable.Cell(1,1).Range.Text = "Name"
$VMTable.Cell(1,2).Range.Text = "Computer Name"
$VMTable.Cell(1,3).Range.Text = "VM Size"
$VMTable.Cell(1,4).Range.Text = "Resource Group Name"
$VMTable.Cell(1,5).Range.Text = "Network Interface"

## Values
$i=0
Foreach ($VM in $VMs) {

        $VMName = $VM.NetworkProfile.NetworkInterfaces.id
        $Parts = $VMName.Split("/")
        $NICLabel = $Parts[8]

    $VMTable.cell(($i+2),1).range.Bold = 0
    $VMTable.cell(($i+2),1).range.text = [string]$VM.Name
    $VMTable.cell(($i+2),2).range.Bold = 0
    $VMTable.cell(($i+2),2).range.text = [string]$VM.OSProfile.ComputerName
    $VMTable.cell(($i+2),3).range.Bold = 0
    $VMTable.cell(($i+2),3).range.text = [string]$VM.HardwareProfile.VmSize
    $VMTable.cell(($i+2),4).range.Bold = 0
    $VMTable.cell(($i+2),4).range.text = [string]$VM.ResourceGroupName
    $VMTable.cell(($i+2),5).range.Bold = 0
    $VMTable.cell(($i+2),5).range.text = [string]$NICLabel

$i++
}


$Word.Selection.Start= $Document.Content.End
$Selection.TypeParagraph()

########
######## NETWORK INTERFACE
########



$Selection.Style = 'Überschrift 1'
$Selection.TypeText("Network Interfaces")
$Selection.TypeParagraph()

$NICs = Get-AzNetworkInterface

## Add a table for NICs
$TableBehavior = [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior

# Tables.Add(Range, RowCoun, ColumnsCount, Behavior)
$NICTable = $Selection.Tables.add($Word.Selection.Range, $NICs.Count + 2, 7, $TableBehavior)


$NICTable.Style = "Default"
$NICTable.Cell(1,1).Range.Text = "Virtual Machine"
$NICTable.Cell(1,2).Range.Text = "Network Card Name"
$NICTable.Cell(1,3).Range.Text = "Resource Group Name"
$NICTable.Cell(1,4).Range.Text = "VNET"
$NICTable.Cell(1,5).Range.Text = "Subnet"
$NICTable.Cell(1,6).Range.Text = "Private IP Address"
$NICTable.Cell(1,7).Range.Text = "Private IP Allocation Method"



## Write NICs to NIC table 
$i=0
Foreach ($NIC in $NICs) {

## Get connected VM, if there is one connected to the network interface
If (!$NIC.VirtualMachine.id) 
    { $VMLabel = " "}
Else
    {
        $VMName = $NIC.VirtualMachine.id
        $Parts = $VMName.Split("/")
        $VMLabel = $PArts[8]
    }

## GET VNET and SUBNET

        $NETCONF = $NIC.IPconfigurations.subnet.id
        $Parts = $NETCONF.Split("/")
        $VNETNAME = $Parts[8]
        $SUBNETNAME = $Parts[10]

    $NICTable.cell(($i+2),1).range.Bold = 0
    $NICTable.cell(($i+2),1).range.text = [string]$VMLabel
    $NICTable.cell(($i+2),2).range.Bold = 0
    $NICTable.cell(($i+2),2).range.text = [string]$NIC.Name
    $NICTable.cell(($i+2),3).range.Bold = 0
    $NICTable.cell(($i+2),3).range.text = [string]$NIC.ResourceGroupName
    $NICTable.cell(($i+2),4).range.Bold = 0
    $NICTable.cell(($i+2),4).range.text = [string]$VNETNAME 
    $NICTable.cell(($i+2),5).range.Bold = 0
    $NICTable.cell(($i+2),5).range.text = [string]$SUBNETNAME
    $NICTable.cell(($i+2),6).range.Bold = 0   
    $NICTable.cell(($i+2),6).range.text = [string]$NIC.IPconfigurations.PrivateIpAddress
    $NICTable.cell(($i+2),7).range.Bold = 0
    $NICTable.cell(($i+2),7).range.text = [string]$NIC.IPconfigurations.PrivateIpAllocationMethod


$i++
}

$Word.Selection.Start= $Document.Content.End
$Selection.TypeParagraph()

########
######## Create a table for NSG
########

## Add some text
$Selection.Style = 'Überschrift 1'
$Selection.TypeText("Network Security Groups")
$Selection.TypeParagraph()

$NSGs = Get-AzNetworkSecurityGroup

## Add a table for NSGs
$TableBehavior = [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior

# Tables.Add(Range, RowCoun, ColumnsCount, Behavior)
$NSGTable = $Selection.Tables.add($Word.Selection.Range, $NSGs.Count + 2, 4, $TableBehavior)

$NSGTable.Style = "Default"
$NSGTable.Cell(1,1).Range.Text = "NSG Name"
$NSGTable.Cell(1,2).Range.Text = "Resource Group Name"
$NSGTable.Cell(1,3).Range.Text = "Network Interfaces"
$NSGTable.Cell(1,4).Range.Text = "Subnets"

## Write NICs to NIC table 
$i=0

Foreach ($NSG in $NSGs) {

    $NICLabel = $null
    $SubnetLabel = $null

## Get connected NIC, if there is one connected 
If (!$NSG.NetworkInterfaces.Id) 
    { $NICLabel = " "}
Else
    {
        Foreach ($NICID in $NSG.NetworkInterfaces.Id) {

            $NICLabel += ("{0},`r`n" -f ($NICID.split("/"))[-1])

        }
        $NICLabel = $NICLabel.Substring(0, ($Niclabel.Length -3))
    }



## Get connected SUBNET, if there is one connected 
If (!$NSG.Subnets.Id) 
    { $SubnetLabel = " "}
Else
    {
        Foreach ($SubnetID in $NSG.Subnets.Id) {

            $SubnetLabel += ("{0},`r`n" -f ($SubnetID.split("/"))[-1])

        }
        $SubnetLabel = $SubnetLabel.Substring(0, ($SubnetLabel.Length -3))

      }


    $NSGTable.cell(($i+2),1).range.Bold = 0
    $NSGTable.cell(($i+2),1).range.text = [string]$NSG.Name
    $NSGTable.cell(($i+2),2).range.Bold = 0
    $NSGTable.cell(($i+2),2).range.text = [string]$NSG.ResourceGroupName
    $NSGTable.cell(($i+2),3).range.Bold = 0
    $NSGTable.cell(($i+2),3).range.text = $NICLabel
    $NSGTable.cell(($i+2),4).range.Bold = 0
    $NSGTable.cell(($i+2),4).range.text = $SubnetLabel



$i++
}

$Word.Selection.Start= $Document.Content.End
$Selection.TypeParagraph()

########
######## Create a table for each NSG
########

### Get all NSGs
$NSGs = Get-AzNetworkSecurityGroup

ForEach ($NSG in $NSGs) {

    ## Add Heading for each NSG
    $Selection.Style = 'Überschrift 2'
    $Selection.TypeText($NSG.Name)
    $Selection.TypeParagraph()

    
        ### Add a table for each NSG, the NSG has custom rules
        $TableBehavior = [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior

        ## Tables.Add(Range, RowCoun, ColumnsCount, Behavior)
        $NSGRuleTable = $Selection.Tables.add($Word.Selection.Range, ($NSG.SecurityRules.Count + $NSG.DefaultSecurityRules.Count) + 2, 9, $TableBehavior)


        $NSGRuleTable.Style = "Default"
        $NSGRuleTable.Cell(1,1).Range.Text = "Rule Name"
        $NSGRuleTable.Cell(1,2).Range.Text = "Protocol"
        $NSGRuleTable.Cell(1,3).Range.Text = "Source Port Range"
        $NSGRuleTable.Cell(1,4).Range.Text = "Destination Port Range"
        $NSGRuleTable.Cell(1,5).Range.Text = "Source Address Prefix"
        $NSGRuleTable.Cell(1,6).Range.Text = "Destination Address Prefix"
        $NSGRuleTable.Cell(1,7).Range.Text = "Access"
        $NSGRuleTable.Cell(1,8).Range.Text = "Priority"
        $NSGRuleTable.Cell(1,9).Range.Text = "Direction"


        ### Get all custom Security Rules in the NSG
        $NSGRules = (Get-AzNetworkSecurityGroup -Name $NSG.Name).SecurityRules
        $i = 0

        ForEach ($NSGRule in $NSGRules) {
                $NSGRuleTable.cell(($i+2),1).range.Bold = 0
                $NSGRuleTable.cell(($i+2),1).range.text = [string]$NSGRule.Name

                $NSGRuleTable.cell(($i+2),2).range.Bold = 0
                $NSGRuleTable.cell(($i+2),2).range.text = [string]$NSGRule.Protocol

                $NSGRuleTable.cell(($i+2),3).range.Bold = 0
                $NSGRuleTable.cell(($i+2),3).range.text = [string]$NSGRule.SourcePortRange

                $NSGRuleTable.cell(($i+2),4).range.Bold = 0
                $NSGRuleTable.cell(($i+2),4).range.text = [string]$NSGRule.DestinationPortRange

                $NSGRuleTable.cell(($i+2),5).range.Bold = 0
                $NSGRuleTable.cell(($i+2),5).range.text = [string]$NSGRule.SourceAddressPrefix

                $NSGRuleTable.cell(($i+2),6).range.Bold = 0
                $NSGRuleTable.cell(($i+2),6).range.text = [string]$NSGRule.DestinationAddressPrefix

                $NSGRuleTable.cell(($i+2),7).range.Bold = 0
                $NSGRuleTable.cell(($i+2),7).range.text = [string]$NSGRule.Access

                $NSGRuleTable.cell(($i+2),8).range.Bold = 0
                $NSGRuleTable.cell(($i+2),8).range.text = [string]$NSGRule.Priority

                $NSGRuleTable.cell(($i+2),9).range.Bold = 0
                $NSGRuleTable.cell(($i+2),9).range.text = [string]$NSGRule.Direction

                $i++
            }

             ### Get all default Security Rules in the NSG
        $NSGRules = (Get-AzNetworkSecurityGroup -Name $NSG.Name).DefaultSecurityRules

        ForEach ($NSGRule in $NSGRules) {
                $NSGRuleTable.cell(($i+2),1).range.Bold = 0
                $NSGRuleTable.cell(($i+2),1).range.text = [string]$NSGRule.Name

                $NSGRuleTable.cell(($i+2),2).range.Bold = 0
                $NSGRuleTable.cell(($i+2),2).range.text = [string]$NSGRule.Protocol

                $NSGRuleTable.cell(($i+2),3).range.Bold = 0
                $NSGRuleTable.cell(($i+2),3).range.text = [string]$NSGRule.SourcePortRange

                $NSGRuleTable.cell(($i+2),4).range.Bold = 0
                $NSGRuleTable.cell(($i+2),4).range.text = [string]$NSGRule.DestinationPortRange

                $NSGRuleTable.cell(($i+2),5).range.Bold = 0
                $NSGRuleTable.cell(($i+2),5).range.text = [string]$NSGRule.SourceAddressPrefix

                $NSGRuleTable.cell(($i+2),6).range.Bold = 0
                $NSGRuleTable.cell(($i+2),6).range.text = [string]$NSGRule.DestinationAddressPrefix

                $NSGRuleTable.cell(($i+2),7).range.Bold = 0
                $NSGRuleTable.cell(($i+2),7).range.text = [string]$NSGRule.Access

                $NSGRuleTable.cell(($i+2),8).range.Bold = 0
                $NSGRuleTable.cell(($i+2),8).range.text = [string]$NSGDRule.Priority

                $NSGRuleTable.cell(($i+2),9).range.Bold = 0
                $NSGRuleTable.cell(($i+2),9).range.text = [string]$NSGDRule.Direction

                $i++
            }

            ### Close the NSG table
            $Word.Selection.Start= $Document.Content.End
            $Selection.TypeParagraph()

        }





### Update the TOC now when all data has been written to the document 
$Document.TablesOfContents(1).Update()

# Save the document
$Report = $PSScriptRoot + "\Azure_Documentation_" + (Get-Date -format "dd-MM-yyyy") + ".docx"
$Document.SaveAs([ref]$Report,[ref]$SaveFormat::wdFormatDocument)
$word.Quit()

# Free up memory
$null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$word)
[gc]::Collect()
[gc]::WaitForPendingFinalizers()
Remove-Variable word 
