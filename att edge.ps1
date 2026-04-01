param (
    [string]$driveLetter,
    [switch]$UpdateComputerPolicy,
    [switch]$RestartExplorer
)

$EdgeProgId = "MSEdgeHTM"
$EdgeExecutable = "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
$FirefoxUninstallerPaths = @(
    "C:\Program Files (x86)\Mozilla Firefox\uninstall\helper.exe",
    "C:\Program Files\Mozilla Firefox\uninstall\helper.exe"
)
$ExcludedUserProfileNames = @("All Users", "Default", "Default User", "Public")
$AdobeExtensionId = "elhekieabhbkpmcefcoobjddigjcaadp"
$AdobeReaderClsId = "{CA8A9780-280D-11CF-A24D-444553540000}"
$script:UserProfilesCache = $null

if ([string]::IsNullOrWhiteSpace($driveLetter)) {
    Write-Error "A letra da unidade mapeada nao foi informada."
    exit 1
}

function Join-DrivePath {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ChildPath
    )

    return Join-Path -Path ("{0}:\" -f $driveLetter) -ChildPath $ChildPath
}

function Get-UserProfiles {
    if ($null -ne $script:UserProfilesCache) {
        return $script:UserProfilesCache
    }

    $script:UserProfilesCache = @(
        Get-ChildItem -Path "C:\Users" -Directory -Force -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -notin $ExcludedUserProfileNames } |
            ForEach-Object {
                [PSCustomObject]@{
                    ProfilePath = $_.FullName
                    DesktopPath = Join-Path -Path $_.FullName -ChildPath "Desktop"
                }
            }
    )

    return $script:UserProfilesCache
}

function Get-UserProfileRegistryEntries {
    return @(
        Get-ChildItem -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList" -ErrorAction SilentlyContinue |
            ForEach-Object {
                try {
                    $profileData = Get-ItemProperty -Path $_.PSPath -ErrorAction Stop
                    $profilePath = $profileData.ProfileImagePath
                    $profileName = Split-Path -Path $profilePath -Leaf
                    $hiveFile = Join-Path -Path $profilePath -ChildPath "NTUSER.DAT"

                    if (
                        $_.PSChildName -match "^S-1-5-21-" -and
                        $profilePath -like "C:\Users\*" -and
                        $profileName -notin $ExcludedUserProfileNames -and
                        (Test-Path -Path $hiveFile)
                    ) {
                        [PSCustomObject]@{
                            Sid = $_.PSChildName
                            ProfilePath = $profilePath
                            HiveFile = $hiveFile
                        }
                    }
                } catch {
                    Write-Host "Falha ao ler o perfil de registro $($_.PSChildName): $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
    )
}

function Remove-FirefoxRemnants {
    Write-Host "Removendo sobras do Firefox..."

    try {
        Get-Process -Name "firefox*" -ErrorAction SilentlyContinue | Stop-Process -Force
        Start-Sleep -Seconds 2
    } catch {
        Write-Host "Nenhum processo do Firefox em execucao."
    }

    $firefoxDirs = @(
        "C:\Program Files (x86)\Mozilla Firefox",
        "C:\Program Files\Mozilla Firefox",
        "$env:APPDATA\Mozilla",
        "$env:LOCALAPPDATA\Mozilla"
    )

    foreach ($dir in $firefoxDirs) {
        if (Test-Path -Path $dir) {
            Write-Host "Removendo diretorio: $dir"
            Remove-Item -Path $dir -Recurse -Force -ErrorAction SilentlyContinue
        }
    }

    $regPaths = @(
        "HKLM:\SOFTWARE\Mozilla",
        "HKLM:\SOFTWARE\WOW6432Node\Mozilla",
        "HKCU:\SOFTWARE\Mozilla",
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Mozilla Firefox",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\Mozilla Firefox"
    )

    foreach ($regPath in $regPaths) {
        if (Test-Path -Path $regPath) {
            Write-Host "Removendo chave de registro: $regPath"
            Remove-Item -Path $regPath -Recurse -Force -ErrorAction SilentlyContinue
        }
    }

    $startMenuPaths = @(
        "$env:ProgramData\Microsoft\Windows\Start Menu\Programs\Mozilla Firefox.lnk",
        "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Mozilla Firefox.lnk"
    )

    foreach ($path in $startMenuPaths) {
        if (Test-Path -Path $path) {
            Remove-Item -Path $path -Force -ErrorAction SilentlyContinue
        }
    }

    Write-Host "Limpeza de sobras do Firefox concluida."
}

function Invoke-FirefoxUninstallCommand {
    param (
        [Parameter(Mandatory = $true)]
        [string]$CommandLine
    )

    if ($CommandLine -match "(?i)msiexec(\.exe)?") {
        $productCodeMatch = [regex]::Match($CommandLine, "{[A-Z0-9\-]+}", "IgnoreCase")
        if ($productCodeMatch.Success) {
            Start-Process -FilePath "msiexec.exe" -ArgumentList "/x $($productCodeMatch.Value) /qn" -Wait -ErrorAction Stop
            return $true
        }
    }

    $exePath = $null
    $arguments = ""

    if ($CommandLine -match '^\s*"([^"]+)"\s*(.*)$') {
        $exePath = $matches[1]
        $arguments = $matches[2]
    } elseif ($CommandLine -match '^\s*([^\s]+)\s*(.*)$') {
        $exePath = $matches[1]
        $arguments = $matches[2]
    }

    if ([string]::IsNullOrWhiteSpace($exePath) -or -not (Test-Path -Path $exePath)) {
        Write-Host "Nao foi possivel localizar o desinstalador do Firefox em: $CommandLine" -ForegroundColor Yellow
        return $false
    }

    if ($arguments -notmatch "(?i)(^|\s)(/S|/silent|/quiet|/qn)(\s|$)") {
        $arguments = ($arguments + " /S").Trim()
    }

    Start-Process -FilePath $exePath -ArgumentList $arguments -Wait -ErrorAction Stop
    return $true
}

function Set-EdgeAsDefaultHttpHandler {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ProgId
    )

    Write-Host "Ajustando associacoes HTTP/HTTPS para o Edge nos perfis locais..."

    foreach ($profile in Get-UserProfileRegistryEntries) {
        $hiveRoot = "Registry::HKEY_USERS\$($profile.Sid)"
        $loadedHere = $false

        if (-not (Test-Path -Path $hiveRoot)) {
            & reg.exe load "HKU\$($profile.Sid)" "$($profile.HiveFile)" | Out-Null
            if ($LASTEXITCODE -ne 0) {
                Write-Host "Nao foi possivel carregar o hive de $($profile.ProfilePath)." -ForegroundColor Yellow
                continue
            }
            $loadedHere = $true
        }

        try {
            foreach ($protocol in @("http", "https")) {
                $assocPath = "$hiveRoot\Software\Microsoft\Windows\Shell\Associations\UrlAssociations\$protocol\UserChoice"
                if (Test-Path -Path $assocPath) {
                    Set-ItemProperty -Path $assocPath -Name ProgId -Value $ProgId -ErrorAction Stop
                }
            }
        } catch {
            Write-Host "Falha ao ajustar navegador padrao em $($profile.ProfilePath): $($_.Exception.Message)" -ForegroundColor Yellow
        } finally {
            if ($loadedHere) {
                & reg.exe unload "HKU\$($profile.Sid)" | Out-Null
            }
        }
    }
}

function Add-UrlsToRegistry {
    param (
        [Parameter(Mandatory = $true)]
        [string]$RegistryPath,
        [Parameter(Mandatory = $true)]
        [string[]]$Urls
    )

    if (-not (Test-Path -Path $RegistryPath)) {
        New-Item -Path $RegistryPath -Force | Out-Null
    }

    Remove-ItemProperty -Path $RegistryPath -Name * -ErrorAction SilentlyContinue

    for ($i = 0; $i -lt $Urls.Length; $i++) {
        $name = ($i + 1).ToString()
        Set-ItemProperty -Path $RegistryPath -Name $name -Value $Urls[$i] -ErrorAction SilentlyContinue
    }
}

function Update-Shortcuts {
    param (
        [Parameter(Mandatory = $true)]
        [string]$FolderPath
    )

    if (-not (Test-Path -Path $FolderPath)) {
        return
    }

    $shell = New-Object -ComObject WScript.Shell

    Get-ChildItem -Path $FolderPath -Filter *.lnk -Force -ErrorAction SilentlyContinue | ForEach-Object {
        try {
            $shortcutObj = $shell.CreateShortcut($_.FullName)
            if ($shortcutObj.TargetPath -like "*firefox.exe*") {
                $shortcutObj.TargetPath = $EdgeExecutable
                $shortcutObj.Save()
            }
        } catch {
            Write-Host "Falha ao atualizar o atalho $($_.FullName): $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }
}

function Restart-ExplorerSafely {
    Write-Host "Reiniciando explorer.exe sob demanda..."

    try {
        Get-Process -Name explorer -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction Stop
        Start-Sleep -Seconds 2
        $newExplorer = Start-Process -FilePath "explorer.exe" -PassThru -ErrorAction Stop
        Start-Sleep -Seconds 2

        if (Get-Process -Id $newExplorer.Id -ErrorAction SilentlyContinue) {
            Write-Host "explorer.exe reiniciado com sucesso."
        } else {
            Write-Host "O explorer.exe nao confirmou reinicio automatico. Reinicie manualmente se necessario." -ForegroundColor Yellow
        }
    } catch {
        Write-Host "Falha ao reiniciar o explorer.exe: $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

Write-Host "Iniciando atualizacao do Edge..."
Write-Host "Obs: reinicio do explorer.exe e atualizacao de politica do computador estao desativados por padrao."

$userProfiles = Get-UserProfiles

# Desinstala firefox

$firefoxUninstalled = $false

foreach ($uninstallerPath in $FirefoxUninstallerPaths) {
    if (Test-Path -Path $uninstallerPath) {
        Write-Host "Desinstalando Mozilla Firefox..."
        try {
            Start-Process -FilePath $uninstallerPath -ArgumentList "/S" -Wait -ErrorAction Stop
            $firefoxUninstalled = $true
            Write-Host "Mozilla Firefox foi desinstalado."
            Remove-FirefoxRemnants
        } catch {
            Write-Host "Falha ao executar o desinstalador do Firefox: $($_.Exception.Message)" -ForegroundColor Red
        }
        break
    }
}

if (-not $firefoxUninstalled) {
    Write-Host "Desinstalador padrao do Firefox nao encontrado. Buscando via registro..."

    $uninstallKeys = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )

    $firefoxFound = $false
    foreach ($keyPath in $uninstallKeys) {
        Get-ChildItem -Path $keyPath -ErrorAction SilentlyContinue | ForEach-Object {
            $displayName = $_.GetValue("DisplayName")
            if ($displayName -like "*Firefox*") {
                $firefoxFound = $true
                $uninstallString = $_.GetValue("UninstallString")
                if ($uninstallString) {
                    Write-Host "Firefox encontrado no registro: $displayName"
                    try {
                        if (Invoke-FirefoxUninstallCommand -CommandLine $uninstallString) {
                            $firefoxUninstalled = $true
                            Remove-FirefoxRemnants
                        }
                    } catch {
                        Write-Host "Falha ao desinstalar via registro: $($_.Exception.Message)" -ForegroundColor Red
                    }
                }
            }
        }
    }

    if (-not $firefoxFound) {
        Write-Host "Nenhuma instalacao do Firefox encontrada no registro."
    }
}

if ($firefoxUninstalled) {
    Write-Host "Verificando sobras do Firefox..."
    $remaining = @()

    foreach ($dir in @(
        "C:\Program Files (x86)\Mozilla Firefox",
        "C:\Program Files\Mozilla Firefox",
        "$env:APPDATA\Mozilla",
        "$env:LOCALAPPDATA\Mozilla"
    )) {
        if (Test-Path -Path $dir) {
            $remaining += "Directory: $dir"
        }
    }

    foreach ($regPath in @(
        "HKLM:\SOFTWARE\Mozilla",
        "HKLM:\SOFTWARE\WOW6432Node\Mozilla",
        "HKCU:\SOFTWARE\Mozilla"
    )) {
        if (Test-Path -Path $regPath) {
            $remaining += "Registry: $regPath"
        }
    }

    if ($remaining.Count -gt 0) {
        Write-Host "Aviso: algumas sobras do Firefox ainda foram encontradas:" -ForegroundColor Yellow
        $remaining | ForEach-Object { Write-Host "- $_" }
    } else {
        Write-Host "Verificacao do Firefox concluida sem sobras aparentes."
    }
} else {
    Write-Host "Nenhuma desinstalacao de Firefox foi necessaria."
}

Write-Host "Processo do Firefox concluido."

# Muda navegador padrão

Set-EdgeAsDefaultHttpHandler -ProgId $EdgeProgId

# Ajusta exceções de segurança do Java

$sourcePath = Join-DrivePath -ChildPath "arquivos\exception.sites"
$targetDir = "AppData\LocalLow\Sun\Java\Deployment\security"

if (Test-Path -Path $sourcePath) {
    foreach ($user in $userProfiles) {
        $targetPath = Join-Path -Path $user.ProfilePath -ChildPath (Join-Path -Path $targetDir -ChildPath "exception.sites")
        $targetDirFull = Split-Path -Path $targetPath -Parent

        if (-not (Test-Path -Path $targetDirFull)) {
            try {
                New-Item -ItemType Directory -Path $targetDirFull -Force | Out-Null
            } catch {
                Write-Host "Erro ao criar diretorio ${targetDirFull}: $($_.Exception.Message)" -ForegroundColor Red
                continue
            }
        }

        try {
            Copy-Item -Path $sourcePath -Destination $targetPath -Force -ErrorAction Stop
        } catch {
            Write-Host "Erro ao copiar exception.sites para ${targetPath}: $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    Write-Host "Processamento do exception.sites concluido para os perfis locais."
} else {
    Write-Host "Arquivo de origem nao encontrado: $sourcePath" -ForegroundColor Red
}

# Deleta Atalhos antigos

$shortcuts = @(
    "EFFORT.lnk",
    "SAM - WEB.lnk",
    "SAMWEB.lnk",
    "Sistema Hospitalar.lnk",
    "Sistemas Hospitalar.lnk",
    "HOSPITALAR.lnk",
    "HOSPITALAR*.lnk",
    "Sistema Hapvida.lnk",
    "Sistemas Hapvida.lnk",
    "SAVI.lnk",
    "Ponto.lnk",
    "PEP.lnk",
    "Chat hapvida.lnk",
    "Chat.lnk",
    "Auditoria de digitais.lnk",
    "Siga Clinicas - Hapclin.lnk",
    "Siga.lnk",
    "Hapclin.lnk",
    "Zimbra - Webmail.lnk",
    "ServiceDesk.lnk",
    "Ambulancia.lnk",
    "Chat Hapvida.URL",
    "Mozila Firefox.lnk",
    "Siga Clinicas.lnk",
    "SIGO HML Hapvida.lnk",
    "SIGO PRD Hapvida.lnk"
)

foreach ($user in $userProfiles) {
    if (-not (Test-Path -Path $user.DesktopPath)) {
        continue
    }

    foreach ($shortcut in $shortcuts) {
        try {
            Get-ChildItem -Path $user.DesktopPath -Filter $shortcut -Force -ErrorAction SilentlyContinue | ForEach-Object {
                Remove-Item -Path $_.FullName -Force -ErrorAction Stop
            }
        } catch {
            Write-Host "Erro ao remover $($user.DesktopPath)\${shortcut}: $($_.Exception.Message)" -ForegroundColor Red
        }
    }
}

# Copia arquivos

$sourceInitial = Join-DrivePath -ChildPath "arquivos\Instaladores"
$destinationInitial = "C:\"

if (-not (Test-Path -Path $sourceInitial)) {
    Write-Host "A pasta de origem '$sourceInitial' nao existe." -ForegroundColor Red
    exit 1
}

Write-Host "Copiando pasta Instaladores para C:\ ..."
try {
    Copy-Item -Path $sourceInitial -Destination $destinationInitial -Recurse -Force -ErrorAction Stop
} catch {
    Write-Host "Falha ao copiar a pasta 'Instaladores': $($_.Exception.Message)" -ForegroundColor Red
}

foreach ($xmlFile in @("apppadrao.xml", "ie_compat_list.xml")) {
    $sourceXml = Join-DrivePath -ChildPath ("arquivos\" + $xmlFile)
    try {
        Copy-Item -Path $sourceXml -Destination $destinationInitial -Force -ErrorAction Stop
    } catch {
        Write-Host "Falha ao copiar '$xmlFile': $($_.Exception.Message)" -ForegroundColor Red
    }
}

$sourceShortcuts = Join-DrivePath -ChildPath "arquivos\ATALHOS"
$destinationShortcuts = "C:\Users\Public\Desktop"

if (-not (Test-Path -Path $sourceShortcuts)) {
    Write-Host "A pasta de atalhos '$sourceShortcuts' nao existe." -ForegroundColor Red
    exit 1
}

if (-not (Test-Path -Path $destinationShortcuts)) {
    Write-Host "A pasta de destino '$destinationShortcuts' nao existe." -ForegroundColor Red
    exit 1
}

Write-Host "Copiando atalhos atualizados..."
try {
    Copy-Item -Path "$sourceShortcuts\*" -Destination $destinationShortcuts -Recurse -Force -ErrorAction Stop
} catch {
    Write-Host "Falha ao copiar atalhos: $($_.Exception.Message)" -ForegroundColor Red
}

$sourceStartup = Join-DrivePath -ChildPath "arquivos\AJUSTE-IE11-OPCAODOWNLOADPDF.bat"
$targetStartup = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\StartUp"

try {
    Copy-Item -Path $sourceStartup -Destination $targetStartup -Force -ErrorAction Stop
} catch {
    Write-Host "Falha ao copiar script de inicializacao: $($_.Exception.Message)" -ForegroundColor Red
}

# Adiciona exceções para conteudo seguro

$allUrls = Get-Content -Path (Join-DrivePath -ChildPath "arquivos\exception.txt") -ErrorAction Stop |
    Where-Object { $_.Trim() -ne "" } |
    ForEach-Object { $_.Trim() } |
    Sort-Object -Unique

$insecureContentPath = "HKLM:\SOFTWARE\Policies\Microsoft\Edge\InsecureContentAllowedForUrls"
$popupsAllowedPath = "HKLM:\SOFTWARE\Policies\Microsoft\Edge\PopupsAllowedForUrls"
$autoDownloadsPath = "HKLM:\SOFTWARE\Policies\Microsoft\Edge\AutomaticDownloadsAllowedForUrls"

Add-UrlsToRegistry -RegistryPath $insecureContentPath -Urls $allUrls
Add-UrlsToRegistry -RegistryPath $popupsAllowedPath -Urls $allUrls
Add-UrlsToRegistry -RegistryPath $autoDownloadsPath -Urls $allUrls

Write-Host "Configuracao concluida: $($allUrls.Count) URLs adicionadas as politicas de Edge."

# Importa Registros

try {
    & reg.exe import (Join-DrivePath -ChildPath "arquivos\edge-ajuste.reg") | Out-Null
    if ($LASTEXITCODE -eq 0) {
        Write-Host "REGISTRO IMPORTADO COM SUCESSO" -ForegroundColor Green
    } else {
        Write-Host "Falha ao importar edge-ajuste.reg." -ForegroundColor Red
    }
} catch {
    Write-Host "Erro ao importar ajustes de registro: $($_.Exception.Message)" -ForegroundColor Red
}

# Adiciona usuários como admin local

try {
    $adminGroups = Get-LocalGroup | Where-Object { $_.Name -like "Administradores*" -or $_.Name -like "Administrators*" }
    if (-not $adminGroups) {
        Write-Host "Grupo local de administradores nao encontrado." -ForegroundColor Red
    } else {
        $localAdmins = $adminGroups[0].Name
        Add-LocalGroupMember -Group $localAdmins -Member "HAPVIDA\Domain Users" 2>$null
        Write-Host "Grupo HAPVIDA\\Domain Users validado em $localAdmins." -ForegroundColor Green

        $currentMembers = net localgroup "$localAdmins" | Select-String -Pattern "HAPVIDA\\Domain Users"
        if ($currentMembers) {
            Write-Host "Domain Users confirmado no grupo local de administradores." -ForegroundColor Green
        } else {
            Write-Host "Nao foi possivel confirmar Domain Users no grupo local." -ForegroundColor Yellow
        }

        $invalidSIDs = net localgroup "$localAdmins" | Select-String -Pattern "S-1-5-21"
        foreach ($invalidSID in $invalidSIDs) {
            $sid = ($invalidSID.Line -replace "^\s+", "").Trim()
            try {
                Remove-LocalGroupMember -Group $localAdmins -Member $sid -ErrorAction Stop
                Write-Host "SID invalido removido de ${localAdmins}: $sid" -ForegroundColor Yellow
            } catch {
                Write-Host "Falha ao remover SID invalido ${sid}: $($_.Exception.Message)" -ForegroundColor Yellow
            }
        }
    }
} catch {
    Write-Host "Erro ao ajustar administradores locais: $($_.Exception.Message)" -ForegroundColor Red
}

# App Padrão

$xmlPath = "C:\apppadrao.xml"

if (-not (Test-Path -Path $xmlPath)) {
    Write-Error "O arquivo XML de associacoes padrao nao existe em $xmlPath."
    exit 1
}

Write-Host "Importando associacoes padrao de aplicativos..."
Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\System" -Name "DefaultAssociationsConfiguration" -Value $xmlPath
Dism.exe /Online /Import-DefaultAppAssociations:C:\Apppadrao.xml

# Altera quaquer atalho para o edge

$shortcutPaths = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
$null = $shortcutPaths.Add("C:\Users\Public\Desktop")
$null = $shortcutPaths.Add("C:\Instaladores\Atalhos")
$null = $shortcutPaths.Add("C:\Users\Default\Desktop")

foreach ($user in $userProfiles) {
    if (Test-Path -Path $user.DesktopPath) {
        $null = $shortcutPaths.Add($user.DesktopPath)
    }
}

foreach ($path in $shortcutPaths) {
    Update-Shortcuts -FolderPath $path
}

& reg.exe add "HKCU\Software\Microsoft\Windows\CurrentVersion\Ext\Settings\$AdobeReaderClsId" /v Flags /t REG_DWORD /d 1 /f | Out-Null
& reg.exe add "HKCU\Software\Microsoft\Windows\CurrentVersion\Ext\Settings\$AdobeReaderClsId" /v Version /d "*" /f | Out-Null

# Instala Adobe no edge como politica

$regPath = "HKLM:\Software\Policies\Microsoft\Edge\ExtensionInstallForcelist"

if (-not (Test-Path -Path $regPath)) {
    New-Item -Path $regPath -Force | Out-Null
}

Set-ItemProperty -Path $regPath -Name "1" -Value "$AdobeExtensionId;https://edge.microsoft.com/extensionwebstorebase/v1/crx"
Write-Host "ADDON DO ADOBE CONFIGURADO COMO POLITICA PARA TODOS OS USUARIOS."

# Restart explorer.exe

if ($RestartExplorer) {
    Restart-ExplorerSafely
} else {
    Write-Host "Reinicio do explorer.exe ignorado para evitar tela preta e instabilidade."
}

# Força politica na maquina

if ($UpdateComputerPolicy) {
    Write-Host "Solicitando atualizacao assicrona da politica do computador..."
    gpupdate /target:computer /wait:0 | Out-Null
    Write-Host "Atualizacao de politica solicitada. O processamento pode continuar em segundo plano."
} else {
    Write-Host "Atualizacao de politica ignorada para reduzir o tempo final. Se necessario, rode gpupdate /target:computer manualmente."
}

Write-Host "############################
#### Script finalizado ####
############################"
