<#
.SYNOPSIS
  Merged all WEC subscription into one

.DESCRIPTION

.INPUTS
  SkeletonPath
  Subscription XML skeleton 
  
  SubscriptionOutputPath
  Subscription XML output
  
.OUTPUT
  Write a subscription in XML format and load it to registry

.NOTES
  Version:        1.1
  Author:         SOUMILL
  Creation Date:  May 10th 2019
  Purpose/Change: Fix query merging
  
.EXAMPLE
 
#>

#---------------------------------------------------------[Initialisations]--------------------------------------------------------

Param(
     [string]$SkeletonPath = ".\Subscription_Skeleton.xml",
     [string]$SubscriptionOutputPath = ".\Merged_Subscription.xml"
)

#----------------------------------------------------------[Declarations]----------------------------------------------------------

#Script Version
$Script_Version = "1.1"

#Path to Subscriptions in Registry
$Subscripton_Base_Path = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\EventCollector\Subscriptions"

#Flag to check if merged subscription is already in production
$Already_Created = $false

#-----------------------------------------------------------[Functions]------------------------------------------------------------

#Check skeleton path
function Check-Skeleton-Path 
{
    if (-not (Test-Path -Path $SkeletonPath) -or [IO.Path]::GetExtension($SkeletonPath) -ne ".xml")
    {
        Write-Error "Invalid path for XML skeleton"
        exit(1)
    }
}

#Check output path
function Check-Output-Path 
{
    if ([IO.Path]::GetExtension($SubscriptionOutputPath) -ne ".xml")
    {
        Write-Error "Invalid path for Output XML"
        exit(1)
    }
}

#Check if Windows Event collector is running
function Check-WEC-Status
{
    if ((Get-Service -Name "Wecsvc").Status -ne "Running")
    {
        Write-Error "WEC server is stopped"
        exit(1)
    }
}

#Load Skeleton Object from XML file
function Load-Skeleton 
{
    [xml]$Subscription_Object = Get-Content $SkeletonPath
    return $Subscription_Object
}


#Return a list with all subscription names
function Create-QueryList ($Subscription_Object)
{
    $Select_String = ""
    $Cpt = 0    
    foreach ($Key in Get-ChildItem -Path $Subscripton_Base_Path)
    {
        $Subscription_Name = Split-Path $Key.Name -Leaf
        #If the key if the generated subscription then ignore the query
        if ($Subscription_Name -eq $Subscription_Object.Subscription.SubscriptionId)
        {
            $global:Already_Created = $true
            continue
        }
        $QueryList = ([xml](Get-ItemProperty -Path "$($Subscripton_Base_Path)\$($Subscription_Name)").Query).QueryList.Query
        foreach ($Query in $QueryList)
        {
            #$Query.Id = $Cpt
            $Query.SetAttribute("Id", $Cpt)
            $Cpt = $Cpt + 1
        }
        #$QueryList_For_Key = ([xml](Get-ItemProperty -Path "$($Subscripton_Base_Path)\$($Subscription_Name)").Query).QueryList.Query.Select
        # Add Select to others
        $Select_String = "$($Select_String)$($QueryList.OuterXml)"
    }
    return "<QueryList>$($Select_String)</QueryList>"
}

#Add freshly create Querylist to Query 
function Update-Skeleton-With-QueryList ($Subscription_Object, $Query_List_Str)
{
    #Add Querylist as CDATA
    $Subscription_Object.Subscription.Query.'#cdata-section' = $Query_List_Str
    return $Subscription_Object
}

#Write merged Subscription in File
function Dump-Subscription-Xml ($Complete_Subscription)
{
    if ($global:Already_Created)
    {
        # Thanks to Microsoft https://support.microsoft.com/en-za/help/4491324/0x57-error-wecutil-command-update-event-forwarding-subscription
        $Node_To_Remove = $Subscription_Object.Subscription.Item("SubscriptionType")
        $Subscription_Object.Subscription.RemoveChild($Node_To_Remove) | Out-Null
    }
    if (Test-Path -Path $SubscriptionOutputPath)
    {
        Move-Item $SubscriptionOutputPath -Destination "$($SubscriptionOutputPath).bak" -Force
    }
    echo "" > $SubscriptionOutputPath
    $Complete_Subscription.Save((Resolve-Path $SubscriptionOutputPath))
}

#Load File to collector
function Load-To-WEC ($Subscription_Object)
{
    #Update Merged_Subscription based on previously created XML
    if ($global:Already_Created)
    {
        wecutil.exe ss /c:$SubscriptionOutputPath
    }
    else
    {
        wecutil.exe cs $SubscriptionOutputPath
    }
}

#-----------------------------------------------------------[Execution]------------------------------------------------------------

Write-Host "[*] -- Start Script -- [*]"

Check-WEC-Status
Check-Output-Path
Check-Skeleton-Path

$Subscription_Object = Load-Skeleton

$Query_List_Str = Create-QueryList $Subscription_Object

$Complete_Subscription = Update-Skeleton-With-QueryList $Subscription_Object $Query_List_Str

Dump-Subscription-Xml $Complete_Subscription

Load-To-WEC $Subscription_Object

Write-Host "[*] -- End Script -- [*]"