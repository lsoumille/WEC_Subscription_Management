#Script Version
$Script_Version = "1.0"

#Path to Subscriptions in Registry
$Subscription_Base_Path = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\EventCollector\Subscriptions"

foreach ($Subscription_Key in Get-ChildItem -Path $Subscription_Base_Path)
{
    $Subscription_Name = Split-Path $Subscription_Key.Name -Leaf
    if ((Test-Path -Path "$($Subscription_Base_Path)\$($Subscription_Name)") -and (-not ($Subscription_Name -eq "Merged_Subscription")))
    {
        Remove-Item -Path "$($Subscription_Base_Path)\$($Subscription_Name)\EventSources" -Confirm:$false -Force -Recurse
    }
}
    
