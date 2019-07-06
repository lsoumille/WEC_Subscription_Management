#Script Version
$Script_Version = "1.0"

#Path to Subscriptions in Registry
$Subscription_Base_Path = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\EventCollector\Subscriptions"
$Targets_Base_Path = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\EventCollector\Subscriptions\Account_Logon\EventSources"
$Merged_Subscription = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\EventCollector\Subscriptions\Merged_Subscription\EventSources"

$Database = @{}

function Update-Database ($Target_Name, $Channel, $Record_Id)
{
    if (-not $Database.ContainsKey($Target_Name))
    {
        $Database.Add($Target_Name, @{$Channel.Trim() = $Record_Id})
        return
    }
    if (-not $Database.$Target_Name.ContainsKey($Channel.Trim()))
    {
        $Database.$Target_Name.Add($Channel.Trim(), $Record_Id)
        return
    }
    if (-not ([uint64]$Database.$Target_Name.$Channel -gt [uint64]$Record_Id))
    {
        $Database.$Target_Name.$Channel.Trim() = $Record_Id
        return
    }
}

function Push-DB-To_Registry ($Database)
{
    foreach ($Target in $Database.Keys)
    {
        $Bookmarks_Str = ""
        $First = $true
        foreach ($Channel in $Database.$Target.keys)
        {
            if ($First)
            {
                $Bookmarks_Str = "$($Bookmarks_Str)<Bookmark Channel=`"$($Channel)`" RecordId=`"$($Database.$Target.$Channel)`" IsCurrent=`"true`"/>"
                $First = $false
            }
            else
            {
                $Bookmarks_Str = "$($Bookmarks_Str)<Bookmark Channel=`"$($Channel)`" RecordId=`"$($Database.$Target.$Channel)`" />"
            }
        }
        $Bookmark_List_Str = "<BookmarkList>$($Bookmarks_Str)</BookmarkList>"
        if (-not (Test-Path -Path "$($Merged_Subscription)\$($Target)"))
        {
            New-Item -Path "$($Merged_Subscription)\$($Target)"
        }
        New-ItemProperty -Path "$($Merged_Subscription)\$($Target)" -Name "Bookmark" -Value $Bookmark_List_Str -PropertyType String -Force
    }
}

#For each server
foreach ($Target_Key in Get-ChildItem -Path $Targets_Base_Path)
{
    $Target_Name = Split-Path $Target_Key.Name -Leaf
    #For each subscription
    foreach ($Subscription_Key in Get-ChildItem -Path $Subscription_Base_Path)
    {
        $Subscription_Name = Split-Path $Subscription_Key.Name -Leaf
        if (Test-Path -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\EventCollector\Subscriptions\$($Subscription_Name)\EventSources\$($Target_Name)")
        {
            #Retrieve Bookmark value
            $Bookmark_List = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\EventCollector\Subscriptions\$($Subscription_Name)\EventSources\$($Target_Name)").Bookmark
            if ($Bookmark_List -and (-not $Bookmark_List.StartsWith("<Book")))
            {
                continue
            }
            $Bookmark_XML = [xml]$Bookmark_List
            #For each bookmark
            foreach ($Bookmark in $Bookmark_XML.BookmarkList.Bookmark)
            {
                Update-Database $Target_Name $Bookmark.Channel $Bookmark.RecordId
            }
        }
    }
    Write-Host $Target_Key
}

Push-DB-To_Registry $Database
