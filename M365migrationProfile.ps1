[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$confirm = Read-Host "Upozornění: 1) Skript smaze vsechny ulozene prihlaseni do M365 aktualne prihlaseneho uzivatele. 2) Vytvori novy Outlook profil O365 a nastavi jej jako vychozi. 3) Otevre Outlook. Pokracovat? (y/n)"
if ($confirm -notmatch "^[Yy]$") { 
    Write-Host "Operace zrušena uživatelem." -ForegroundColor Yellow
    exit 
}

# Funkce pro bezpečné provedení příkazů a barevnou zpětnou vazbu
function Try-Execute {
    param (
        [string]$Description,
        [scriptblock]$Command
    )
    try {
        & $Command
        Write-Host "$Description - Úspěšně provedeno." -ForegroundColor Green
    } catch {
        Write-Host "$Description - Chyba: $_" -ForegroundColor Red
    }
}

# Ukončení procesů
Try-Execute "Ukončení Outlooku" { Stop-Process -Name "outlook" -Force -ErrorAction SilentlyContinue }
Try-Execute "Ukončení Microsoft Teams" { Stop-Process -Name "ms-teams" -Force -ErrorAction SilentlyContinue }

# Odstranění klíčů registru a souborů
Try-Execute "Odstranění Identity klíčů" { Remove-Item -Path "HKCU:\Software\Microsoft\Office\16.0\Common\Identity" -Recurse -Force -ErrorAction SilentlyContinue }
Try-Execute "Odstranění WorkplaceJoin JoinInfo" { Remove-Item -Path "HKCU:\Software\Microsoft\Windows NT\CurrentVersion\WorkplaceJoin\JoinInfo" -Recurse -Force -ErrorAction SilentlyContinue }
Try-Execute "Odstranění WorkplaceJoin TenantInfo" { Remove-Item -Path "HKCU:\Software\Microsoft\Windows NT\CurrentVersion\WorkplaceJoin\TenantInfo" -Recurse -Force -ErrorAction SilentlyContinue }
Try-Execute "Odstranění Office Licensing" { Remove-Item -Path "$($env:LOCALAPPDATA)\Microsoft\Office\16.0\Licensing" -Recurse -ErrorAction SilentlyContinue }
Try-Execute "Odstranění IdentityCache" { Remove-Item -Path "$($env:LOCALAPPDATA)\Microsoft\IdentityCache" -Recurse -ErrorAction SilentlyContinue }
Try-Execute "Odstranění OneAuth" { Remove-Item -Path "$($env:LOCALAPPDATA)\Microsoft\OneAuth" -Recurse -ErrorAction SilentlyContinue }

# Vytvoření nového profilu Outlooku a jeho nastavení jako výchozí
Try-Execute "Vytvoření nového Outlook profilu O365" { New-Item -Path HKCU:\SOFTWARE\Microsoft\Office\16.0\Outlook\Profiles -Name O365 -Force -ErrorAction SilentlyContinue }
Try-Execute "Nastavení O365 jako výchozího profilu" { Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Outlook" -Name "DefaultProfile" -Value O365 -ErrorAction SilentlyContinue }

# Spuštění Outlooku
Try-Execute "Spuštění Outlooku" { Start-Process -Name outlook }

Write-Host "Skript dokončen!" -ForegroundColor Cyan
