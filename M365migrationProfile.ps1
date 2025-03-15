$confirm = Read-Host "Warning: 1) This script will remove all stored M365 sign-ins for the current user. 2) It will create a new Outlook profile 'O365' and set it as the default. 3) It will open Outlook. Continue? (y/n)"
if ($confirm -notmatch "^[Yy]$") { 
    Write-Host "Operation canceled by user." -ForegroundColor Yellow
    exit 
}

# Function to execute commands safely and display colored status messages
function Try-Execute {
    param (
        [string]$Description,
        [scriptblock]$Command
    )
    try {
        & $Command
        Write-Host "$Description - Success." -ForegroundColor Green
    } catch {
        Write-Host "$Description - Error: $_" -ForegroundColor Red
    }
}

# Kill processes
Try-Execute "Stopping Outlook" { Stop-Process -Name "outlook" -Force -ErrorAction SilentlyContinue }
Try-Execute "Stopping Microsoft Teams" { Stop-Process -Name "ms-teams" -Force -ErrorAction SilentlyContinue }

# Remove registry keys and local files
Try-Execute "Removing Identity registry keys" { Remove-Item -Path "HKCU:\Software\Microsoft\Office\16.0\Common\Identity" -Recurse -Force -ErrorAction SilentlyContinue }
Try-Execute "Removing WorkplaceJoin JoinInfo" { Remove-Item -Path "HKCU:\Software\Microsoft\Windows NT\CurrentVersion\WorkplaceJoin\JoinInfo" -Recurse -Force -ErrorAction SilentlyContinue }
Try-Execute "Removing WorkplaceJoin TenantInfo" { Remove-Item -Path "HKCU:\Software\Microsoft\Windows NT\CurrentVersion\WorkplaceJoin\TenantInfo" -Recurse -Force -ErrorAction SilentlyContinue }
Try-Execute "Removing Office Licensing data" { Remove-Item -Path "$($env:LOCALAPPDATA)\Microsoft\Office\16.0\Licensing" -Recurse -ErrorAction SilentlyContinue }
Try-Execute "Removing IdentityCache" { Remove-Item -Path "$($env:LOCALAPPDATA)\Microsoft\IdentityCache" -Recurse -ErrorAction SilentlyContinue }
Try-Execute "Removing OneAuth data" { Remove-Item -Path "$($env:LOCALAPPDATA)\Microsoft\OneAuth" -Recurse -ErrorAction SilentlyContinue }

# Create new Outlook profile and set as default
Try-Execute "Creating new Outlook profile 'O365'" { New-Item -Path HKCU:\SOFTWARE\Microsoft\Office\16.0\Outlook\Profiles -Name O365 -Force -ErrorAction SilentlyContinue }
Try-Execute "Setting 'O365' as default profile" { Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Outlook" -Name "DefaultProfile" -Value O365 -ErrorAction SilentlyContinue }

# Start Outlook
Try-Execute "Starting Outlook" { Start-Process outlook }

Write-Host "Script completed!" -ForegroundColor Cyan
