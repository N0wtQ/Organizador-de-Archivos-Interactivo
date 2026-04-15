#requires -Version 5.1
<#
Script: Organizador de Archivos Interactivo
Descripción:
- Escanea recursivamente una carpeta origen.
- Clasifica archivos por temática usando palabras clave en nombre/contenido.
- Renombra con formato YYYY-MM-DD_[NombreOriginal].ext.
- Mueve a D:\Organizado\[Temática]\.
- Incluye preview, estadísticas, log, conflictos y deshacer.
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
function Initialize-ConsoleEncoding {
    try {
        [Console]::InputEncoding = [System.Text.UTF8Encoding]::new($false)
        [Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
        $OutputEncoding = [System.Text.UTF8Encoding]::new($false)
    } catch {
        # Si falla el ajuste de codificacion, continuar sin interrumpir.
    }
}

# =========================
# Configuración global
# =========================
$script:State = [ordered]@{
    SourcePath                 = $null
    BaseDestination            = "D:\Organizado"
    AutoExcludedExtensions     = @(".exe", ".dll", ".sys", ".msi", ".bat", ".cmd", ".ps1", ".vbs")
    AutoExcludedDirectories    = @("Windows", "Program Files", "Program Files (x86)", "AppData", "ProgramData", '$Recycle.Bin', "System Volume Information")
    CustomExcludedExtensions   = @()
    CustomExcludedDirectories  = @()
    LastPreviewPlan            = @()
    LastRunStats               = $null
    LastRunStartedAt           = $null
}

$script:WritableDirCache = @{}
$script:LogFilePath = Join-Path $script:State.BaseDestination "organizacion_log.txt"
$script:UndoMapPath = Join-Path $script:State.BaseDestination "ultimo_mapeo_movimientos.json"

# =========================
# Utilidades
# =========================
function Ensure-BaseDestination {
    if (-not (Test-Path -LiteralPath $script:State.BaseDestination)) {
        New-Item -Path $script:State.BaseDestination -ItemType Directory -Force | Out-Null
    }
}

function Write-Log {
    param(
        [Parameter(Mandatory = $true)][string]$Message,
        [ValidateSet("INFO", "WARN", "ERROR")][string]$Level = "INFO"
    )
    try {
        Ensure-BaseDestination
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $line = "[{0}] [{1}] {2}" -f $timestamp, $Level, $Message
        Add-Content -LiteralPath $script:LogFilePath -Value $line -Encoding UTF8
    } catch {
        Write-Warning "No se pudo escribir en el log: $($_.Exception.Message)"
    }
}

function Pause-Continue {
    [void](Read-Host "Presiona Enter para continuar")
}

function Normalize-FileName {
    param([Parameter(Mandatory = $true)][string]$InputName)

    $invalid = [System.IO.Path]::GetInvalidFileNameChars()
    $safe = $InputName
    foreach ($char in $invalid) {
        $safe = $safe.Replace($char, "_")
    }
    $safe = $safe -replace "\s+", " "
    $safe = $safe.Trim()
    if ([string]::IsNullOrWhiteSpace($safe)) {
        $safe = "SinNombre"
    }
    return $safe
}

function Is-ExcludedDirectory {
    param([Parameter(Mandatory = $true)][string]$FullPath)

    # Compara por segmentos de carpeta para detectar exclusiones en cualquier nivel.
    $segments = $FullPath -split "[\\/]" | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    $excludedDirs = @($script:State.AutoExcludedDirectories + $script:State.CustomExcludedDirectories)
    foreach ($segment in $segments) {
        foreach ($excluded in $excludedDirs) {
            if ($segment.Equals($excluded, [System.StringComparison]::OrdinalIgnoreCase)) {
                return $true
            }
        }
    }
    return $false
}

function Is-ExcludedExtension {
    param([Parameter(Mandatory = $true)][string]$Extension)

    $allExcluded = @($script:State.AutoExcludedExtensions + $script:State.CustomExcludedExtensions) |
        ForEach-Object { $_.ToLowerInvariant() }
    return $allExcluded -contains $Extension.ToLowerInvariant()
}

function Get-DocumentText {
    param([Parameter(Mandatory = $true)][System.IO.FileInfo]$File)

    # Se limita lectura para evitar consumo excesivo.
    $maxChars = 6000
    $ext = $File.Extension.ToLowerInvariant()

    try {
        switch ($ext) {
            ".txt" { return (Get-Content -LiteralPath $File.FullName -Raw -ErrorAction Stop).Substring(0, [Math]::Min($maxChars, (Get-Content -LiteralPath $File.FullName -Raw).Length)) }
            ".md"  { return (Get-Content -LiteralPath $File.FullName -Raw -ErrorAction Stop).Substring(0, [Math]::Min($maxChars, (Get-Content -LiteralPath $File.FullName -Raw).Length)) }
            ".csv" { return (Get-Content -LiteralPath $File.FullName -Raw -ErrorAction Stop).Substring(0, [Math]::Min($maxChars, (Get-Content -LiteralPath $File.FullName -Raw).Length)) }
            ".json"{ return (Get-Content -LiteralPath $File.FullName -Raw -ErrorAction Stop).Substring(0, [Math]::Min($maxChars, (Get-Content -LiteralPath $File.FullName -Raw).Length)) }
            ".xml" { return (Get-Content -LiteralPath $File.FullName -Raw -ErrorAction Stop).Substring(0, [Math]::Min($maxChars, (Get-Content -LiteralPath $File.FullName -Raw).Length)) }
            ".log" { return (Get-Content -LiteralPath $File.FullName -Raw -ErrorAction Stop).Substring(0, [Math]::Min($maxChars, (Get-Content -LiteralPath $File.FullName -Raw).Length)) }
            ".ini" { return (Get-Content -LiteralPath $File.FullName -Raw -ErrorAction Stop).Substring(0, [Math]::Min($maxChars, (Get-Content -LiteralPath $File.FullName -Raw).Length)) }
            ".docx" {
                Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue
                $zip = [System.IO.Compression.ZipFile]::OpenRead($File.FullName)
                try {
                    $entry = $zip.Entries | Where-Object { $_.FullName -eq "word/document.xml" } | Select-Object -First 1
                    if ($null -eq $entry) { return "" }
                    $reader = New-Object System.IO.StreamReader($entry.Open())
                    try {
                        $xml = $reader.ReadToEnd()
                    } finally {
                        $reader.Dispose()
                    }
                    $plain = ($xml -replace "<[^>]+>", " ")
                    if ($plain.Length -gt $maxChars) { $plain = $plain.Substring(0, $maxChars) }
                    return $plain
                } finally {
                    $zip.Dispose()
                }
            }
            default {
                # Para otros tipos, no intenta parseo profundo.
                return ""
            }
        }
    } catch {
        Write-Log -Level "WARN" -Message ("No se pudo leer contenido de {0}: {1}" -f $File.FullName, $_.Exception.Message)
        return ""
    }
}

function Get-ThemeFromDocument {
    param(
        [Parameter(Mandatory = $true)][System.IO.FileInfo]$File,
        [Parameter(Mandatory = $true)][string]$Text
    )

    # Diccionario de palabras clave por temática.
    $themes = [ordered]@{
        "Facturas"  = @("factura", "invoice", "iva", "rfc", "subtotal", "total", "pago", "cliente", "proveedor", "cfdi")
        "Contratos" = @("contrato", "acuerdo", "cláusula", "clausula", "firmado", "vigencia", "anexo", "arrendamiento", "servicio")
        "Personal"  = @("curriculum", "cv", "dni", "pasaporte", "nómina", "nomina", "empleado", "personal", "vacaciones", "salario")
        "Trabajo"   = @("reunión", "reunion", "meeting", "minuta", "tarea", "entrega", "kpi", "objetivo", "empresa", "oficina")
        "Proyectos" = @("proyecto", "roadmap", "milestone", "sprint", "backlog", "arquitectura", "diseño", "diseno", "release", "deploy")
        "Finanzas"  = @("balance", "estado de cuenta", "presupuesto", "ingresos", "egresos", "gasto", "financiero", "contabilidad")
        "Legal"     = @("legal", "demanda", "jurídico", "juridico", "ley", "reglamento", "licencia")
        "Academico" = @("universidad", "escuela", "curso", "certificado", "tesis", "investigación", "investigacion", "clase")
    }

    $haystack = ("{0} {1}" -f $File.Name, $Text).ToLowerInvariant()
    $bestTheme = "SinClasificar"
    $bestScore = 0

    foreach ($theme in $themes.Keys) {
        $score = 0
        foreach ($keyword in $themes[$theme]) {
            if ($haystack.Contains($keyword.ToLowerInvariant())) {
                $score++
            }
        }
        if ($score -gt $bestScore) {
            $bestScore = $score
            $bestTheme = $theme
        }
    }

    return $bestTheme
}

function Ensure-DirectoryWritable {
    param([Parameter(Mandatory = $true)][string]$DirectoryPath)

    if ($script:WritableDirCache.ContainsKey($DirectoryPath)) {
        return [bool]$script:WritableDirCache[$DirectoryPath]
    }

    try {
        if (-not (Test-Path -LiteralPath $DirectoryPath)) {
            New-Item -Path $DirectoryPath -ItemType Directory -Force | Out-Null
        }
        $tempFile = Join-Path $DirectoryPath ("__perm_test_{0}.tmp" -f [guid]::NewGuid().ToString("N"))
        Set-Content -LiteralPath $tempFile -Value "perm-test" -Encoding UTF8 -ErrorAction Stop
        Remove-Item -LiteralPath $tempFile -Force -ErrorAction Stop
        $script:WritableDirCache[$DirectoryPath] = $true
        return $true
    } catch {
        $script:WritableDirCache[$DirectoryPath] = $false
        Write-Log -Level "ERROR" -Message ("Sin permisos de escritura en {0}: {1}" -f $DirectoryPath, $_.Exception.Message)
        return $false
    }
}

function Ensure-FileReadable {
    param([Parameter(Mandatory = $true)][string]$FilePath)

    try {
        $stream = [System.IO.File]::Open($FilePath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
        $stream.Dispose()
        return $true
    } catch {
        Write-Log -Level "WARN" -Message ("Sin permisos/lectura en {0}: {1}" -f $FilePath, $_.Exception.Message)
        return $false
    }
}

function Get-UniqueDestinationPath {
    param([Parameter(Mandatory = $true)][string]$DestinationPath)
    if (-not (Test-Path -LiteralPath $DestinationPath)) {
        return $DestinationPath
    }
    $directory = Split-Path -LiteralPath $DestinationPath -Parent
    $fileName = [System.IO.Path]::GetFileNameWithoutExtension($DestinationPath)
    $ext = [System.IO.Path]::GetExtension($DestinationPath)
    $index = 1
    do {
        $candidate = Join-Path $directory ("{0} ({1}){2}" -f $fileName, $index, $ext)
        $index++
    } while (Test-Path -LiteralPath $candidate)
    return $candidate
}

function Resolve-Conflict {
    param(
        [Parameter(Mandatory = $true)][string]$SourcePath,
        [Parameter(Mandatory = $true)][string]$DestinationPath
    )

    Write-Host ""
    Write-Host "Conflicto detectado:" -ForegroundColor Yellow
    Write-Host ("- Origen : {0}" -f $SourcePath)
    Write-Host ("- Destino: {0}" -f $DestinationPath)
    Write-Host "Opciones: [S]obrescribir / [R]enombrar / [O]mitir"

    while ($true) {
        $option = (Read-Host "Selecciona una opcion (S/R/O)").Trim().ToUpperInvariant()
        switch ($option) {
            "S" { return [pscustomobject]@{ Action = "Overwrite"; FinalPath = $DestinationPath } }
            "R" {
                $newPath = Get-UniqueDestinationPath -DestinationPath $DestinationPath
                Write-Host ("Se usará nombre alternativo: {0}" -f $newPath) -ForegroundColor Cyan
                return [pscustomobject]@{ Action = "Rename"; FinalPath = $newPath }
            }
            "O" { return [pscustomobject]@{ Action = "Skip"; FinalPath = $null } }
            default { Write-Host "Opcion invalida. Escribe S, R u O." -ForegroundColor Red }
        }
    }
}

# =========================
# Núcleo de escaneo/preview
# =========================
function Build-ClassificationPreview {
    param([Parameter(Mandatory = $true)][string]$SourcePath)

    if (-not (Test-Path -LiteralPath $SourcePath)) {
        throw "La ruta origen no existe."
    }

    Write-Host "Escaneando archivos, por favor espera..." -ForegroundColor Cyan
    Write-Log -Message ("Inicio de escaneo en: {0}" -f $SourcePath)

    $allFiles = Get-ChildItem -LiteralPath $SourcePath -Recurse -File -Force -ErrorAction SilentlyContinue
    $total = @($allFiles).Count
    $current = 0
    $plan = New-Object System.Collections.Generic.List[object]

    foreach ($file in $allFiles) {
        $current++
        $percent = if ($total -gt 0) { [int](($current / $total) * 100) } else { 100 }
        Write-Progress -Id 1 -Activity "Escaneo y clasificación" -Status ("Procesando {0}/{1}: {2}" -f $current, $total, $file.Name) -PercentComplete $percent

        if (Is-ExcludedDirectory -FullPath $file.FullName) { continue }
        if (Is-ExcludedExtension -Extension $file.Extension) { continue }
        if (-not (Ensure-FileReadable -FilePath $file.FullName)) { continue }

        $rawText = Get-DocumentText -File $file
        $theme = Get-ThemeFromDocument -File $file -Text $rawText

        $datePart = $file.LastWriteTime.ToString("yyyy-MM-dd")
        $safeOriginal = Normalize-FileName -InputName $file.BaseName
        $newName = "{0}_[{1}]{2}" -f $datePart, $safeOriginal, $file.Extension

        $themeFolder = Join-Path $script:State.BaseDestination $theme
        $destination = Join-Path $themeFolder $newName

        $plan.Add([pscustomobject]@{
                SourcePath      = $file.FullName
                OriginalName    = $file.Name
                Theme           = $theme
                DestinationPath = $destination
                DestinationDir  = $themeFolder
            })
    }

    Write-Progress -Id 1 -Activity "Escaneo y clasificación" -Completed
    Write-Log -Message ("Fin de escaneo. Archivos candidatos: {0}" -f $plan.Count)
    return $plan
}

function Show-Preview {
    param([Parameter(Mandatory = $true)][array]$Plan)

    if ($Plan.Count -eq 0) {
        Write-Host "No hay archivos para organizar con la configuración actual." -ForegroundColor Yellow
        return
    }

    $byTheme = $Plan | Group-Object -Property Theme | Sort-Object -Property Count -Descending
    Write-Host ""
    Write-Host "=== Resumen por temática ===" -ForegroundColor Cyan
    foreach ($group in $byTheme) {
        Write-Host ("- {0}: {1}" -f $group.Name, $group.Count)
    }

    Write-Host ""
    Write-Host "=== Muestra de previsualización (hasta 20) ===" -ForegroundColor Cyan
    $Plan | Select-Object -First 20 | ForEach-Object {
        Write-Host ("[{0}] {1}" -f $_.Theme, $_.OriginalName)
        Write-Host ("  -> {0}" -f $_.DestinationPath) -ForegroundColor DarkGray
    }

    Write-Host ""
    Write-Host ("Total de archivos en previsualización completa: {0}" -f $Plan.Count) -ForegroundColor Green
}

# =========================
# Ejecución de movimientos
# =========================
function Execute-Organization {
    param([Parameter(Mandatory = $true)][array]$Plan)

    if ($Plan.Count -eq 0) {
        Write-Host "No hay nada para organizar. Genera una previsualización primero." -ForegroundColor Yellow
        return
    }

    Ensure-BaseDestination
    $startTime = Get-Date
    $movements = New-Object System.Collections.Generic.List[object]
    $createdThemes = New-Object System.Collections.Generic.HashSet[string]

    $stats = [ordered]@{
        StartTime         = $startTime
        EndTime           = $null
        TotalPlanned      = $Plan.Count
        Moved             = 0
        Skipped           = 0
        Errors            = 0
        ThemesCreated     = 0
        ThemesTouched     = 0
        LastLogPath       = $script:LogFilePath
        LastUndoMapPath   = $script:UndoMapPath
    }

    Write-Host ""
    Write-Host "Se van a mover $($Plan.Count) archivos." -ForegroundColor Cyan
    $confirm = (Read-Host "Deseas continuar? (S/N)").Trim().ToUpperInvariant()
    if ($confirm -ne "S") {
        Write-Host "Operación cancelada por el usuario."
        Write-Log -Level "WARN" -Message "Organización cancelada por el usuario antes de ejecutar movimientos."
        return
    }

    $i = 0
    foreach ($item in $Plan) {
        $i++
        $percent = [int](($i / $Plan.Count) * 100)
        Write-Progress -Id 2 -Activity "Moviendo archivos" -Status ("Procesando {0}/{1}: {2}" -f $i, $Plan.Count, $item.OriginalName) -PercentComplete $percent

        try {
            if (-not (Test-Path -LiteralPath $item.SourcePath)) {
                $stats.Skipped++
                Write-Log -Level "WARN" -Message ("Omitido (no existe): {0}" -f $item.SourcePath)
                continue
            }

            if (-not (Ensure-FileReadable -FilePath $item.SourcePath)) {
                $stats.Skipped++
                continue
            }

            if (-not (Ensure-DirectoryWritable -DirectoryPath $item.DestinationDir)) {
                $stats.Errors++
                continue
            }

            if (-not (Test-Path -LiteralPath $item.DestinationDir)) {
                New-Item -Path $item.DestinationDir -ItemType Directory -Force | Out-Null
                [void]$createdThemes.Add($item.Theme)
            } elseif (-not $createdThemes.Contains($item.Theme)) {
                # Cuenta temática tocada aunque ya exista.
                [void]$createdThemes.Add($item.Theme)
            }

            $finalDestination = $item.DestinationPath
            if (Test-Path -LiteralPath $finalDestination) {
                $decision = Resolve-Conflict -SourcePath $item.SourcePath -DestinationPath $finalDestination
                switch ($decision.Action) {
                    "Skip" {
                        $stats.Skipped++
                        Write-Log -Level "WARN" -Message ("Omitido por conflicto: {0}" -f $item.SourcePath)
                        continue
                    }
                    "Rename" {
                        $finalDestination = $decision.FinalPath
                    }
                    "Overwrite" {
                        # Se continuará con -Force.
                    }
                }
            }

            Move-Item -LiteralPath $item.SourcePath -Destination $finalDestination -Force -ErrorAction Stop
            $stats.Moved++
            $movements.Add([pscustomobject]@{
                    MovedFrom = $item.SourcePath
                    MovedTo   = $finalDestination
                    Theme     = $item.Theme
                    Timestamp = (Get-Date).ToString("o")
                })

            Write-Log -Message ("Movido: {0} -> {1}" -f $item.SourcePath, $finalDestination)
        } catch {
            $stats.Errors++
            Write-Log -Level "ERROR" -Message ("Error moviendo {0}: {1}" -f $item.SourcePath, $_.Exception.Message)
        }
    }

    Write-Progress -Id 2 -Activity "Moviendo archivos" -Completed

    $stats.EndTime = Get-Date
    $stats.ThemesTouched = ($Plan | Select-Object -ExpandProperty Theme -Unique).Count
    $stats.ThemesCreated = $createdThemes.Count

    # Guarda mapeo de última ejecución para deshacer.
    try {
        $movements | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $script:UndoMapPath -Encoding UTF8
        Write-Log -Message ("Mapeo de deshacer guardado en: {0}" -f $script:UndoMapPath)
    } catch {
        Write-Log -Level "ERROR" -Message ("No se pudo guardar mapeo de deshacer: {0}" -f $_.Exception.Message)
    }

    $script:State.LastRunStats = [pscustomobject]$stats
    $script:State.LastRunStartedAt = $startTime

    Write-Host ""
    Write-Host "Organización finalizada." -ForegroundColor Green
    Show-Stats
}

function Undo-LastOrganization {
    if (-not (Test-Path -LiteralPath $script:UndoMapPath)) {
        Write-Host "No hay mapeo de la última organización para deshacer." -ForegroundColor Yellow
        return
    }

    try {
        $raw = Get-Content -LiteralPath $script:UndoMapPath -Raw -ErrorAction Stop
        $entries = $raw | ConvertFrom-Json -ErrorAction Stop
        if ($null -eq $entries) {
            Write-Host "El mapeo de deshacer está vacío." -ForegroundColor Yellow
            return
        }
    } catch {
        Write-Host "No se pudo leer el archivo de deshacer: $($_.Exception.Message)" -ForegroundColor Red
        Write-Log -Level "ERROR" -Message ("Error leyendo undo map: {0}" -f $_.Exception.Message)
        return
    }

    if ($entries -isnot [System.Collections.IEnumerable]) {
        $entries = @($entries)
    }

    Write-Host ("Se intentará deshacer {0} movimientos." -f @($entries).Count) -ForegroundColor Cyan
    $confirm = (Read-Host "¿Confirmas deshacer la última organización? (S/N)").Trim().ToUpperInvariant()
    if ($confirm -ne "S") {
        Write-Host "Deshacer cancelado."
        return
    }

    $restored = 0
    $skipped = 0
    $errors = 0
    $total = @($entries).Count
    $index = 0

    # Revertimos en orden inverso para minimizar conflictos.
    foreach ($entry in @($entries)[($total - 1)..0]) {
        $index++
        $percent = [int](($index / $total) * 100)
        Write-Progress -Id 3 -Activity "Deshaciendo organización" -Status ("Revirtiendo {0}/{1}" -f $index, $total) -PercentComplete $percent

        try {
            $from = $entry.MovedTo
            $to = $entry.MovedFrom

            if (-not (Test-Path -LiteralPath $from)) {
                $skipped++
                Write-Log -Level "WARN" -Message ("Undo omitido (origen no existe): {0}" -f $from)
                continue
            }

            $targetDir = Split-Path -LiteralPath $to -Parent
            if (-not (Ensure-DirectoryWritable -DirectoryPath $targetDir)) {
                $errors++
                continue
            }

            $finalTo = $to
            if (Test-Path -LiteralPath $finalTo) {
                $decision = Resolve-Conflict -SourcePath $from -DestinationPath $finalTo
                switch ($decision.Action) {
                    "Skip" {
                        $skipped++
                        continue
                    }
                    "Rename" { $finalTo = $decision.FinalPath }
                    "Overwrite" { }
                }
            }

            Move-Item -LiteralPath $from -Destination $finalTo -Force -ErrorAction Stop
            $restored++
            Write-Log -Message ("Undo: {0} -> {1}" -f $from, $finalTo)
        } catch {
            $errors++
            Write-Log -Level "ERROR" -Message ("Error en undo para {0}: {1}" -f $entry.MovedTo, $_.Exception.Message)
        }
    }

    Write-Progress -Id 3 -Activity "Deshaciendo organización" -Completed
    Write-Host ""
    Write-Host "Deshacer finalizado." -ForegroundColor Green
    Write-Host ("Restaurados: {0} | Omitidos: {1} | Errores: {2}" -f $restored, $skipped, $errors)
}

function Show-Stats {
    if ($null -eq $script:State.LastRunStats) {
        Write-Host "No hay estadísticas aún. Ejecuta una organización primero." -ForegroundColor Yellow
        return
    }

    $s = $script:State.LastRunStats
    Write-Host ""
        Write-Host "=== Estadisticas de la ultima ejecucion ===" -ForegroundColor Cyan
    Write-Host ("Inicio               : {0}" -f $s.StartTime)
    Write-Host ("Fin                  : {0}" -f $s.EndTime)
    Write-Host ("Archivos planificados: {0}" -f $s.TotalPlanned)
    Write-Host ("Archivos movidos     : {0}" -f $s.Moved)
    Write-Host ("Omitidos             : {0}" -f $s.Skipped)
    Write-Host ("Errores              : {0}" -f $s.Errors)
    Write-Host ("Tematicas tocadas    : {0}" -f $s.ThemesTouched)
    Write-Host ("Tematicas creadas    : {0}" -f $s.ThemesCreated)
    Write-Host ("Log                  : {0}" -f $s.LastLogPath)
    Write-Host ("Undo map             : {0}" -f $s.LastUndoMapPath)
}

# =========================
# Configuración de exclusiones
# =========================
function Show-CurrentExclusions {
    Write-Host ""
    Write-Host "Extensiones excluidas (auto + personalizadas):" -ForegroundColor Cyan
    ($script:State.AutoExcludedExtensions + $script:State.CustomExcludedExtensions | Select-Object -Unique) |
        ForEach-Object { Write-Host ("- {0}" -f $_) }

    Write-Host ""
    Write-Host "Carpetas excluidas (auto + personalizadas):" -ForegroundColor Cyan
    ($script:State.AutoExcludedDirectories + $script:State.CustomExcludedDirectories | Select-Object -Unique) |
        ForEach-Object { Write-Host ("- {0}" -f $_) }
}

function Configure-Exclusions {
    while ($true) {
        Clear-Host
        Write-Host "=== Configurar exclusiones ===" -ForegroundColor Green
        Write-Host "1) Ver exclusiones actuales"
        Write-Host "2) Agregar extension personalizada"
        Write-Host "3) Quitar extension personalizada"
        Write-Host "4) Agregar carpeta personalizada"
        Write-Host "5) Quitar carpeta personalizada"
        Write-Host "6) Volver al menu principal"

        $choice = (Read-Host "Elige una opcion").Trim()
        switch ($choice) {
            "1" {
                Show-CurrentExclusions
                Pause-Continue
            }
            "2" {
                $ext = (Read-Host "Ingresa extension (ej: .tmp)").Trim().ToLowerInvariant()
                if (-not $ext.StartsWith(".")) { $ext = "." + $ext }
                if (-not ($script:State.CustomExcludedExtensions -contains $ext)) {
                    $script:State.CustomExcludedExtensions += $ext
                    Write-Host "Extension agregada: $ext" -ForegroundColor Green
                } else {
                    Write-Host "Ya existe en exclusiones personalizadas." -ForegroundColor Yellow
                }
                Pause-Continue
            }
            "3" {
                $ext = (Read-Host "Ingresa extension a quitar").Trim().ToLowerInvariant()
                if (-not $ext.StartsWith(".")) { $ext = "." + $ext }
                $script:State.CustomExcludedExtensions = @($script:State.CustomExcludedExtensions | Where-Object { $_ -ne $ext })
                Write-Host "Extension eliminada (si existia): $ext" -ForegroundColor Green
                Pause-Continue
            }
            "4" {
                $dir = (Read-Host "Ingresa nombre de carpeta a excluir (ej: Backup)").Trim()
                if (-not [string]::IsNullOrWhiteSpace($dir)) {
                    if (-not ($script:State.CustomExcludedDirectories -contains $dir)) {
                        $script:State.CustomExcludedDirectories += $dir
                        Write-Host "Carpeta excluida agregada: $dir" -ForegroundColor Green
                    } else {
                        Write-Host "Ya existe en exclusiones personalizadas." -ForegroundColor Yellow
                    }
                }
                Pause-Continue
            }
            "5" {
                $dir = (Read-Host "Ingresa nombre de carpeta a quitar").Trim()
                $script:State.CustomExcludedDirectories = @($script:State.CustomExcludedDirectories | Where-Object { $_ -ne $dir })
                Write-Host "Carpeta eliminada (si existia): $dir" -ForegroundColor Green
                Pause-Continue
            }
            "6" { return }
            default {
                Write-Host "Opcion invalida." -ForegroundColor Red
                Pause-Continue
            }
        }
    }
}

function Select-SourceFolder {
    $path = (Read-Host "Ingresa ruta origen a escanear").Trim()
    if ([string]::IsNullOrWhiteSpace($path)) {
        Write-Host "Ruta vacía. Operación cancelada." -ForegroundColor Yellow
        return
    }
    if (-not (Test-Path -LiteralPath $path)) {
        Write-Host "La ruta no existe: $path" -ForegroundColor Red
        return
    }
    $script:State.SourcePath = $path
    $script:State.LastPreviewPlan = @()
    Write-Host "Ruta origen establecida: $path" -ForegroundColor Green
}

function Build-And-ShowPreview {
    if ([string]::IsNullOrWhiteSpace($script:State.SourcePath)) {
        Write-Host "Primero selecciona una carpeta origen." -ForegroundColor Yellow
        return
    }
    try {
        $plan = Build-ClassificationPreview -SourcePath $script:State.SourcePath
        $script:State.LastPreviewPlan = @($plan)
        Show-Preview -Plan $script:State.LastPreviewPlan
    } catch {
        Write-Host "Error en previsualizacion: $($_.Exception.Message)" -ForegroundColor Red
        Write-Log -Level "ERROR" -Message ("Error en preview: {0}" -f $_.Exception.Message)
    }
}

function Execute-FromPreview {
    if ($script:State.LastPreviewPlan.Count -eq 0) {
        Write-Host "No hay preview cargada. Se generará una nueva ahora..." -ForegroundColor Yellow
        Build-And-ShowPreview
    }
    if ($script:State.LastPreviewPlan.Count -eq 0) {
        return
    }
    Execute-Organization -Plan $script:State.LastPreviewPlan
}

function Show-MainMenu {
    while ($true) {
        Clear-Host
        Write-Host "=== Organizador de Archivos Interactivo ===" -ForegroundColor Green
        Write-Host ("Origen actual         : {0}" -f ($(if ($script:State.SourcePath) { $script:State.SourcePath } else { "(no definido)" })))
        Write-Host ("Carpeta base destino  : {0}" -f $script:State.BaseDestination)
        Write-Host ""
        Write-Host "1) Seleccionar carpeta origen a escanear"
        Write-Host "2) Previsualizar clasificacion antes de mover"
        Write-Host "3) Ejecutar organizacion"
        Write-Host "4) Ver estadisticas"
        Write-Host "5) Configurar exclusiones personalizadas"
        Write-Host "6) Deshacer ultima organizacion"
        Write-Host "7) Salir"

        $option = (Read-Host "Selecciona una opcion").Trim()
        switch ($option) {
            "1" {
                Select-SourceFolder
                Pause-Continue
            }
            "2" {
                Build-And-ShowPreview
                Pause-Continue
            }
            "3" {
                Execute-FromPreview
                Pause-Continue
            }
            "4" {
                Show-Stats
                Pause-Continue
            }
            "5" {
                Configure-Exclusions
            }
            "6" {
                Undo-LastOrganization
                Pause-Continue
            }
            "7" {
                Write-Host "Saliendo..." -ForegroundColor Cyan
                break
            }
            default {
                Write-Host "Opcion invalida." -ForegroundColor Red
                Pause-Continue
            }
        }
    }
}

# =========================
# Inicio
# =========================
try {
    Initialize-ConsoleEncoding
    Ensure-BaseDestination
    Write-Log -Message "Inicio del script de organización interactiva."
    Show-MainMenu
    Write-Log -Message "Fin de sesión del script."
} catch {
    Write-Host "Error crítico: $($_.Exception.Message)" -ForegroundColor Red
    Write-Log -Level "ERROR" -Message ("Error crítico: {0}" -f $_.Exception.Message)
}
