$confirm = Read-Host "Upozorneni: Skript smaze vsechny ulozene prihlaseni do M365 aktualne prihlaseneho uzivatele. Vytvori novy Outlook profil O365 a nastavi jej jako vychozi. Otevre Outlook. Spustit? (y/n)"
if ($confirm -notmatch "^[Yy]$") { exit }
Stop-Process -Name "outlook" -Force -ErrorAction SilentlyContinue
Stop-Process -Name "ms-teams" -Force -ErrorAction SilentlyContinue
Remove-Item -Path "HKCU:\Software\Microsoft\Office\16.0\Common\Identity" -Recurse -Force
Remove-Item -Path "HKCU:\Software\Microsoft\Windows NT\CurrentVersion\WorkplaceJoin\JoinInfo" -Recurse -Force
Remove-Item -Path "HKCU:\Software\Microsoft\Windows NT\CurrentVersion\WorkplaceJoin\TenantInfo" -Recurse -Force
remove-item -Path "$($env:LOCALAPPDATA)\Microsoft\Office\16.0\Licensing" -Recurse
remove-item -Path "$($env:LOCALAPPDATA)\Microsoft\IdentityCache" -Recurse
remove-item -Path "$($env:LOCALAPPDATA)\Microsoft\OneAuth" -Recurse
New-Item -Path HKCU:\SOFTWARE\Microsoft\Office\16.0\Outlook\Profiles -Name O365 -force
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Outlook" -Name "DefaultProfile" -Value O365
start-process -name outlook
