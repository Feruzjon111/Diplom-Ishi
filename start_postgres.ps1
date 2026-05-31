param(
    [string]$Database = "amaliyotdocx",
    [string]$User = "postgres",
    [string]$HostName = "127.0.0.1",
    [string]$Port = "5432",
    [switch]$ImportSqliteData
)

$ErrorActionPreference = "Stop"

Push-Location $PSScriptRoot
try {
    $python = Join-Path $PSScriptRoot ".venv\Scripts\python.exe"
    if (-not (Test-Path $python)) {
        throw "Virtual muhit topilmadi: $python"
    }

    $psql = "C:\Program Files\PostgreSQL\17\bin\psql.exe"
    $createdb = "C:\Program Files\PostgreSQL\17\bin\createdb.exe"
    if (-not (Test-Path $psql)) {
        $psql = "psql"
    }
    if (-not (Test-Path $createdb)) {
        $createdb = "createdb"
    }

    $securePassword = Read-Host "PostgreSQL '$User' parolini kiriting" -AsSecureString
    $passwordPointer = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($securePassword)
    $password = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($passwordPointer)
    [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($passwordPointer)

    $env:PGPASSWORD = $password
    $dbExists = & $psql -h $HostName -p $Port -U $User -d postgres -tAc "SELECT 1 FROM pg_database WHERE datname = '$Database';"
    if (($dbExists | Out-String).Trim() -ne "1") {
        & $createdb -h $HostName -p $Port -U $User $Database
    }

    if ($ImportSqliteData) {
        $fixturePath = Join-Path $env:TEMP "amaliyotdocx-sqlite-export.json"
        $env:DATABASE_URL = ""
        $env:DB_ENGINE = "sqlite"
        $env:SQLITE_DB_PATH = Join-Path $PSScriptRoot "amaliyotdocx.sqlite3"
        & $python manage.py dumpdata --natural-foreign --natural-primary --exclude contenttypes --exclude auth.Permission --indent 2 -o $fixturePath
    }

    $env:DATABASE_URL = ""
    $env:DB_ENGINE = "postgresql"
    $env:POSTGRES_DB = $Database
    $env:POSTGRES_USER = $User
    $env:POSTGRES_PASSWORD = $password
    $env:POSTGRES_HOST = $HostName
    $env:POSTGRES_PORT = $Port

    & $python manage.py migrate

    if ($ImportSqliteData) {
        & $python manage.py loaddata $fixturePath
    }

    & $python manage.py create_admin

    $runserverProcesses = Get-CimInstance Win32_Process |
        Where-Object { $_.Name -like "python*" -and $_.CommandLine -like "*manage.py runserver 127.0.0.1:8000*" }
    foreach ($process in $runserverProcesses) {
        Stop-Process -Id $process.ProcessId -Force -ErrorAction SilentlyContinue
    }

    $outLog = Join-Path $PSScriptRoot "server.current.8000.out.log"
    $errLog = Join-Path $PSScriptRoot "server.current.8000.err.log"
    $server = Start-Process -FilePath $python `
        -ArgumentList @("manage.py", "runserver", "127.0.0.1:8000", "--noreload") `
        -WorkingDirectory $PSScriptRoot `
        -WindowStyle Hidden `
        -RedirectStandardOutput $outLog `
        -RedirectStandardError $errLog `
        -PassThru

    Write-Host "Server PostgreSQL bilan ishga tushdi: http://127.0.0.1:8000/"
    Write-Host "Process ID: $($server.Id)"
}
finally {
    Pop-Location
}
