$olApp = New-Object -ComObject Outlook.Application
$olApp.Quit()
Remove-Variable olApp

Get-ChildItem $env:LOCALAPPDATA\Microsoft\Outlook\* -Include dirigenten_rbartists_at*, solisten_rbartists_at*, tournee1_rbartists_at* | Remove-Item

$Profiles = Get-ChildItem 'HKCU:\SOFTWARE\Microsoft\Office\16.0\Outlook\Profiles\Outlook' |foreach {
    Get-ChildItem $_.Name |foreach {
        Get-ChildItem -Path $_.Name
    } 
}
 
foreach($Profile in $Profiles){
    try{
        $AccountName = Get-ItemPropertyValue -Path $Profile.Name -Name 'Account Name' -ErrorAction Stop
        if($AccountName -like 'dirigenten_rbartists_at*'){
            'HKCU:\' + ($Profile.Name.Split('\')[1..7] -join '\') | Remove-Item -Recurse
        }
		if($AccountName -like 'solisten_rbartists_at*'){
            'HKCU:\' + ($Profile.Name.Split('\')[1..7] -join '\') | Remove-Item -Recurse
        }
		if($AccountName -like 'tournee1_rbartists_at*'){
            'HKCU:\' + ($Profile.Name.Split('\')[1..7] -join '\') | Remove-Item -Recurse
        }
		
    }catch{
        Continue
    }
}