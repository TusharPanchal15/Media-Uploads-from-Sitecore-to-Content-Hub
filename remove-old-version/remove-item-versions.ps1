Import-Function -Name ConvertTo-Xlsx
$item = Get-Item -Path "master:\content"
$dialogProps = @{
    Parameters = @(
        @{ Name = "item"; Title="Branch to analyse"; Root="/sitecore/content"},
        @{ Name = "count"; Value=10; Title="Max number of versions";  Editor="number"},
        @{ Name = "remove"; Value=$False; Title="Do you wish to remove items?"; Editor="check"}
    )
    Title = "Limit item version count"
    Description = "Sitecore recommends keeping 10 or fewer versions on any item, but policy may dictate this to be a higher number."
    Width = 500
    Height = 280
    OkButtonName = "Proceed"
    CancelButtonName = "Abort"
}
$result = Read-Variable @dialogProps 
if($result -ne "ok") {
    Close-Window
    Exit
}
$items = @()
Get-Item -Path master: -ID $item.ID -Language * | ForEach-Object { $items += @($_) + @(($_.Axes.GetDescendants())) | Where-Object { $_.Versions.Count -gt $count } | Initialize-Item }
$ritems = @()
$items | ForEach-Object {
    $citem = Get-Item -Path master: -ID $_.ID -Language $_.Language
    if ($citem) {
        $latestVersionItem = $citem.Versions.GetLatestVersion()
        $minVersion = $latestVersionItem.Version.Number - $count
        $ritems += Get-Item -Path master: -ID $_.ID -Language $_.Language -Version * | Where-Object { $_.Version.Number -le $minVersion }
    }
}
if ($remove) {
    $toRemove = $ritems.Count
    $ritems | ForEach-Object {
        $_ | Remove-ItemVersion
    }
    Show-Alert "Removed $toRemove versions"
} else {
    $reportProps = @{
        Property = @(
            "DisplayName",
            @{Name="Version"; Expression={$_.Version}},
            @{Name="Path"; Expression={$_.ItemPath}},
            @{Name="Language"; Expression={$_.Language}}
        )
        Title = "Versions proposed to remove"
        InfoTitle = "Sitecore recommendation: Limit the number of versions of any item to the fewest possible."
        InfoDescription = "The report shows all items that have more than <b>$count versions</b>."
    }
    [byte[]]$outobject = $ritems | Show-ListView @reportProps | ConvertTo-Xlsx
    Out-Download -Name "report-$datetime.xlsx" -InputObject $outobject
}
Close-Window
