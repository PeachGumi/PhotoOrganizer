param(
    [string]$Version = "0.0.0-ci",
    [string]$OutputRoot = ""
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

if ([string]::IsNullOrWhiteSpace($OutputRoot)) {
    $OutputRoot = Join-Path ([IO.Path]::GetTempPath()) ("PhotoOrganizer-package-smoke-" + [Guid]::NewGuid().ToString('N'))
}

function Get-PeMachine([string]$Path) {
    $stream = [IO.File]::OpenRead($Path)
    try {
        $reader = [IO.BinaryReader]::new($stream)
        if ($reader.ReadUInt16() -ne 0x5A4D) { throw "Not an MZ executable: $Path" }
        $stream.Position = 0x3c
        $peOffset = $reader.ReadInt32()
        if ($peOffset -lt 0 -or $peOffset + 6 -gt $stream.Length) { throw "Invalid PE offset: $Path" }
        $stream.Position = $peOffset
        if ($reader.ReadUInt32() -ne 0x00004550) { throw "Missing PE signature: $Path" }
        return $reader.ReadUInt16()
    }
    finally {
        $stream.Dispose()
    }
}

$expectedMachine = @{
    'x64' = 0x8664
    'arm64' = 0xAA64
}

try {
    New-Item -ItemType Directory -Force -Path $OutputRoot | Out-Null

    foreach ($arch in @('x64', 'arm64')) {
        $rid = "win-$arch"
        $publish = Join-Path $OutputRoot "$rid\publish"
        $packageDir = Join-Path $OutputRoot "$rid\package"
        New-Item -ItemType Directory -Force -Path $publish, $packageDir | Out-Null

        dotnet publish src/PhotoOrganizer.App/PhotoOrganizer.App.csproj `
            --configuration Release `
            --runtime $rid `
            --self-contained true `
            -p:Version=$Version `
            -p:ContinuousIntegrationBuild=true `
            --output $publish
        if ($LASTEXITCODE -ne 0) { throw "dotnet publish failed for $rid" }

        $required = @(
            'PhotoOrganizer.exe',
            'PhotoOrganizer.dll',
            'PhotoOrganizer.Core.dll',
            'PhotoOrganizer.deps.json',
            'PhotoOrganizer.runtimeconfig.json'
        )
        foreach ($name in $required) {
            $file = Join-Path $publish $name
            if (-not (Test-Path -LiteralPath $file -PathType Leaf)) { throw "Required publish output missing: $file" }
            if ((Get-Item -LiteralPath $file).Length -le 0) { throw "Required publish output is empty: $file" }
        }

        $exe = Join-Path $publish 'PhotoOrganizer.exe'
        $machine = Get-PeMachine $exe
        if ($machine -ne $expectedMachine[$arch]) {
            throw ("Unexpected PE machine for {0}: 0x{1:X4}" -f $rid, $machine)
        }

        Copy-Item LICENSE (Join-Path $publish 'LICENSE.txt')
        Copy-Item -Recurse -Force (Join-Path $publish '*') $packageDir

        $zip = Join-Path $OutputRoot "PhotoOrganizer-Windows-$arch-$Version.zip"
        if (Test-Path $zip) { Remove-Item $zip -Force }
        Compress-Archive -Path (Join-Path $packageDir '*') -DestinationPath $zip -CompressionLevel Optimal
        if (-not (Test-Path -LiteralPath $zip -PathType Leaf) -or (Get-Item $zip).Length -le 0) {
            throw "Smoke package was not created: $zip"
        }

        $extract = Join-Path $OutputRoot "$rid\extract"
        Expand-Archive -LiteralPath $zip -DestinationPath $extract
        foreach ($name in @('PhotoOrganizer.exe', 'PhotoOrganizer.dll', 'PhotoOrganizer.Core.dll', 'LICENSE.txt')) {
            if (-not (Test-Path -LiteralPath (Join-Path $extract $name) -PathType Leaf)) {
                throw "Packaged file missing after ZIP round trip: $name"
            }
        }

        $hash = (Get-FileHash -Algorithm SHA256 -LiteralPath $zip).Hash.ToLowerInvariant()
        if ($hash -notmatch '^[0-9a-f]{64}$') { throw "Invalid package SHA-256: $zip" }
        Write-Host "Packaging smoke passed for $rid: $zip ($hash)"
    }
}
finally {
    if (Test-Path $OutputRoot) {
        Remove-Item $OutputRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
}
