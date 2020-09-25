#Requires -RunAsAdministrator

<#
Verion 1.0

By: Arjan Sturing

PowerShell fork of the Windows OpenVPN 2.4.x Easy-RSA2 scripts.

Automate the World! #PowerShell
#>

# Script variables
$env:ovpndir=(New-Object -ComObject WScript.Shell).RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\OpenVPN\InstallLocation") | %{$_.Substring(0, $_.length - 1) }
$env:powerrsadir="$env:ovpndir\powerrsa"


# Function for creating a new PKI
Function initpki {
Remove-Item $env:powerrsadir\pki -Force -Recurse -ErrorAction SilentlyContinue
md $env:powerrsadir -Force -ErrorAction SilentlyContinue
md "$env:powerrsadir\pki" -Force -ErrorAction SilentlyContinue
New-Item $env:powerrsadir\variables.ps1 -Force -ErrorAction SilentlyContinue
New-Item $env:powerrsadir\pki\index.txt -Force -ErrorAction SilentlyContinue
New-Item $env:powerrsadir\pki\serial -Force -ErrorAction SilentlyContinue
add-content $env:powerrsadir\pki\serial "01" -Force -ErrorAction SilentlyContinue
$OVPNDIR=$env:ovpndir
banner
Write-Host ""
Write-Host "Config init started..." -ForegroundColor Green
Write-Host ""
$COUNTRY=Read-Host "Enter Country"
$PROVINCE= Read-Host "Enter State of Province"
$CITY=Read-Host "Enter City"
$ORG=Read-Host "Enter Organization Name"
$EMAIL=Read-Host "Enter E-Mail Address"
$OU=Read-Host "Enter Department Name"
add-content $env:powerrsadir\variables.ps1 ('$env:ovpndir="'+($OVPNDIR)+('"')) -Force -ErrorAction SilentlyContinue
add-content $env:powerrsadir\variables.ps1  '$env:powerrsadir="$env:ovpndir\powerrsa"' -Force -ErrorAction SilentlyContinue
add-content $env:powerrsadir\variables.ps1  '$env:PATH="$env:ovpndir\bin"' -Force -ErrorAction SilentlyContinue
add-content $env:powerrsadir\variables.ps1  '$env:HOME=$env:powerrsadir' -Force -ErrorAction SilentlyContinue
add-content $env:powerrsadir\variables.ps1  '$env:KEY_CONFIG="$env:powerrsadir\config\openssl-1.0.0.cnf"' -Force -ErrorAction SilentlyContinue
add-content $env:powerrsadir\variables.ps1  '$env:KEY_DIR="pki"'-Force -ErrorAction SilentlyContinue
add-content $env:powerrsadir\variables.ps1  '$env:DH_KEY_SIZE="2048"' -Force -ErrorAction SilentlyContinue
add-content $env:powerrsadir\variables.ps1  '$env:KEY_SIZE="4096"' -Force -ErrorAction SilentlyContinue
add-content $env:powerrsadir\variables.ps1  ('$env:KEY_COUNTRY="'+($COUNTRY)+('"')) -Force -ErrorAction SilentlyContinue
add-content $env:powerrsadir\variables.ps1  ('$env:KEY_PROVINCE="'+($PROVINCE)+('"'))-Force -ErrorAction SilentlyContinue
add-content $env:powerrsadir\variables.ps1  ('$env:KEY_CITY="'+($CITY)+('"')) -Force -ErrorAction SilentlyContinue
add-content $env:powerrsadir\variables.ps1  ('$env:KEY_ORG="'+($ORG)+('"')) -Force -ErrorAction SilentlyContinue
add-content $env:powerrsadir\variables.ps1  ('$env:KEY_EMAIL="'+($EMAIL)+('"')) -Force -ErrorAction SilentlyContinue
add-content $env:powerrsadir\variables.ps1  ('$env:KEY_OU="'+($OU)+('"')) -Force -ErrorAction SilentlyContinue
add-content $env:powerrsadir\variables.ps1  '$env:PKCS11_MODULE_PATH="changeme"' -Force -ErrorAction SilentlyContinue
add-content $env:powerrsadir\variables.ps1  '$env:PKCS11_PIN="1234"' -Force -ErrorAction SilentlyContinue
cd $env:powerrsadir
. .\Variables.ps1
cls
Write-Host "Config init completed" -ForeGroundColor Green
Start-Sleep 5
}

# Function for creating CA certificate
Function createca {
cls
banner
$env:KEY_CN=Read-Host "Enter CA Name"
$env:KEY_NAME=$env:KEY_CN
openssl req -days 3650 -nodes -new -x509 -keyout $env:KEY_DIR\ca.key -out $env:KEY_DIR\ca.crt -config $env:KEY_CONFIG -batch 
Start-Sleep 2
cls
Write-Host "CA Certificate created" -ForeGroundColor Green
Write-Host "You can find the files in the following directory: $env:HOME\pki" -ForeGroundColor Green
Start-Sleep 5 
}

# Function for creating server key
Function createserver {
cls
banner
$env:KEY_CN=Read-Host "Enter Servername"
$env:KEY_NAME=$env:KEY_CN
cd $env:HOME
openssl req -days 3650 -nodes -new -keyout $env:KEY_DIR\$env:KEY_CN.key -out $env:KEY_DIR\$env:KEY_CN.csr -config $env:KEY_CONFIG -batch
openssl ca -days 3650 -out $env:KEY_DIR\$env:KEY_CN.crt -in $env:KEY_DIR\$env:KEY_CN.csr -extensions server -config $env:KEY_CONFIG -batch
cd $env:HOME
cd $env:KEY_DIR
Get-ChildItem *.old | foreach { Remove-Item -Path $_.FullName }
Start-Sleep 2
cls
Write-Host "Server Certificate created" -ForeGroundColor Green
Write-Host "You can find the files in the following directory: $env:HOME\pki" -ForeGroundColor Green
Start-Sleep 5
}

# Function for creatina Diffie-Hellman key
Function createdh {
cls
banner
cd $env:HOME
openssl dhparam -out $env:KEY_DIR\DH$env:DH_KEY_SIZE.pem $env:DH_KEY_SIZE   
Start-Sleep 2
cls
Write-Host "Diffie-Hellman key created" -ForeGroundColor Green
Write-Host "You can find the files in the following directory: $env:HOME\pki" -ForeGroundColor Green
Start-Sleep 5
}

# Function for creatinga TLS-Auth key
Function createta {
cls
banner
openvpn --genkey --secret $env:HOME\$env:KEY_DIR\ta.key   
Start-Sleep 2
cls
Write-Host "TLS-Auth key created" -ForeGroundColor Green
Write-Host "You can find the files in the following directory: $env:HOME\pki" -ForeGroundColor Green
Start-Sleep 5
}

# Function for creating client certificate without password
Function client {
cls
banner
$env:KEY_CN=Read-Host "Enter Clientname"
cd $env:HOME
openssl req -days 3650 -nodes -new -keyout $env:KEY_DIR\$env:KEY_CN.key -out $env:KEY_DIR\$env:KEY_CN.csr -config $env:KEY_CONFIG -batch
openssl ca -days 3650 -out $env:KEY_DIR\$env:KEY_CN.crt -in $env:KEY_DIR\$env:KEY_CN.csr -config $env:KEY_CONFIG -batch
cd $env:HOME
cd $env:KEY_DIR
Get-ChildItem *.old | foreach { Remove-Item -Path $_.FullName }
Start-Sleep 2
cls
Write-Host "Client certificate created" -ForeGroundColor Green
Write-Host "You can find the files in the following directory: $env:HOME\pki" -ForeGroundColor Green
Start-Sleep 5
}

# Function for creating client certificate without password
Function clientpwd {
cls
banner
$env:KEY_CN=Read-Host "Enter Clientname"
cd $env:HOME
openssl req -days 3650 -new -keyout $env:KEY_DIR\$env:KEY_CN.key -out $env:KEY_DIR\$env:KEY_CN.csr -config $env:KEY_CONFIG -batch
openssl ca -days 3650 -out $env:KEY_DIR\$env:KEY_CN.key -in $env:KEY_DIR\$env:KEY_CN.csr -config $env:KEY_CONFIG -batch
cd $env:HOME
cd $env:KEY_DIR
Get-ChildItem *.old | foreach { Remove-Item -Path $_.FullName }
Start-Sleep 2
cls
Write-Host "Client certificate created" -ForeGroundColor Green
Write-Host "You can find the files in the following directory: $env:HOME\pki" -ForeGroundColor Green
Start-Sleep 5
}


# Script banner
Function Banner {
Clear-Host
Write-Host " ######  ####### #     # ####### ######        ######   #####     #   " -ForeGroundColor Red 
Write-Host " #     # #     # #  #  # #       #     #       #     # #     #   # #   " -ForeGroundColor Red 
Write-Host " #     # #     # #  #  # #       #     #       #     # #        #   #  " -ForeGroundColor Red 
Write-Host " ######  #     # #  #  # #####   ######  ##### ######   #####  #     # " -ForeGroundColor Red 
Write-Host " #       #     # #  #  # #       #   #         #   #         # ####### " -ForeGroundColor Red 
Write-Host " #       #     # #  #  # #       #    #        #    #  #     # #     # " -ForeGroundColor Red 
Write-Host " #       #######  ## ##  ####### #     #       #     #  #####  #     # " -ForeGroundColor Red 
Write-Host ""
Write-Host " Version 1.0 " -ForeGroundColor Red 
Write-Host ""
Write-Host "By: Arjan Sturing" -ForeGroundColor Green
Write-Host ""
}

Function MainMenu {
do {
    do {
Banner                                                                       
Write-Host "Select option:"
Write-Host ""
Write-Host "1: Create CA Certificate"
Write-Host "2: Create Server Certificate"
Write-Host "3: Create Diffie-Hellman Certificate"
Write-Host "4: Create TLS-Auth key"
Write-Host "5: Create Create Client Certificate without password"
Write-Host "6: Create Create Client Certificate with password"
Write-Host "N: Create new PKI"
Write-Host "Q: Quit"
Write-Host ""
Write-Host "Automate the world! #PowerShell" -ForegroundColor Yellow
Write-Host ""
        write-host -nonewline "Enter choice and press Enter: "
        
        $choice = read-host
        
        write-host ""
        
        $ok = $choice -match '^[123456nq]+$'
        if ( -not $ok) {
        cls 
        Write-Host "Wrong Choice!" -ForegroundColor Red
        Start-Sleep 5
        cls
        }
    } until ( $ok )
    
    switch -Regex ( $choice ) {
        "N"
        {
            initpki
        }
        
        "1"
        {
            createca 
           
        }

        "2"
        {
           createserver
        }

        "3"
        {
           createdh 
        }
        "4"
        {
           createta
        }
        "5"
        {
           client
          
        }
        "6"
        {
           clientpwd
          
        }
    }
} until ( $choice -match "Q" )
} 

Banner
If (Test-Path $env:powerrsadir\variables.ps1) {cd $env:powerrsadir
. .\Variables.ps1
}
Else{
initpki
}
banner
mainmenu
