#Script Version
$Script_Version = "1.0"

#Path to Subscriptions in Registry
$Subscription_Base_Path = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\EventCollector\Subscriptions"

foreach ($Subscription_Key in Get-ChildItem -Path $Subscription_Base_Path)
{
    $Subscription_Name = Split-Path $Subscription_Key.Name -Leaf
    if (Test-Path -Path "$($Subscription_Base_Path)\$($Subscription_Name)")
    {
        New-ItemProperty -Path "$($Subscription_Base_Path)\$($Subscription_Name)" -Name "Enabled" -Value 0 -PropertyType DWORD -Force
    }
}
    
