$olApp = New-Object -ComObject Outlook.Application
$olApp.Quit()
Remove-Variable olApp