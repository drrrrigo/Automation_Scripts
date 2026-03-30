$ErrorActionPreference = 'SilentlyContinue'

$logFolderName = 'pixviewerlogs'
$usersRoot = 'C:\Users'

if (-not (Test-Path -LiteralPath $usersRoot)) {
    exit 0
}

$excludedProfiles = @(
    'All Users',
    'Default',
    'Default User',
    'Public',
    'defaultuser0',
    'WDAGUtilityAccount'
)

$profiles = Get-ChildItem -LiteralPath $usersRoot -Directory -Force | Where-Object {
    $excludedProfiles -notcontains $_.Name
}

foreach ($profile in $profiles) {
    try {
        $targetFolders = @()

        $directFolder = Join-Path -Path $profile.FullName -ChildPath $logFolderName
        if (Test-Path -LiteralPath $directFolder) {
            $targetFolders += Get-Item -LiteralPath $directFolder -Force
        }

        $targetFolders += Get-ChildItem -LiteralPath $profile.FullName -Directory -Recurse -Force -Filter $logFolderName

        foreach ($folder in $targetFolders) {
            try {
                $items = Get-ChildItem -LiteralPath $folder.FullName -Force

                foreach ($item in $items) {
                    try {
                        Remove-Item -LiteralPath $item.FullName -Recurse -Force -ErrorAction Stop
                    }
                    catch {
                        # Ignora apenas o item que estiver em uso, bloqueado ou sem acesso.
                    }
                }
            }
            catch {
                # Ignora pastas momentaneamente inacessiveis.
            }
        }
    }
    catch {
        # Ignora perfis com restricao de leitura.
    }
}

exit 0
