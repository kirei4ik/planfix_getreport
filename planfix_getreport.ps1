# ===================================================================
# SETTINGS
# ===================================================================
$VerboseMode        = $false      
$UpdateExisting     = $false      # ставь true если хочешь обновлять какой-то файл / лист в текущей папке
$ExistingFileName   = "data.xlsx" 
$SheetName          = "Лист1"     

$Token      = "Bearer INSERT_TOKEN" #апи-токен с нужными правами
$Account    = "ACCOUNT_NAME" #имя аккаунта
$ReportId   = "Report_ID" #идентификатор отчета, можно взять из параметров ссылки
# ===================================================================

[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

function Write-VerboseLog {
    param([string]$Message, [string]$Color = "Gray")
    if ($VerboseMode) { Write-Host $Message -ForegroundColor $Color }
}

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
if (-not $ScriptDir) { $ScriptDir = $PWD.Path }

if ($UpdateExisting) {
    $ExcelFile = Join-Path $ScriptDir $ExistingFileName
    Write-Host "MODE: UPDATE existing file" -ForegroundColor Magenta
} else {
    $ExcelFile = Join-Path $ScriptDir "report_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx"
    Write-Host "MODE: CREATE new file" -ForegroundColor Magenta
}

$Headers = @{
    "Authorization" = $Token
    "Accept"        = "application/json"
    "Content-Type"  = "application/json"
}

$BaseUrl = "https://$Account.planfix.ru/rest"
$SaveId = $null
$SkipGeneration = $false

# Step 1: Generate (Invoke-RestMethod - работает стабильно)
Write-Host "`n[1/3] Generating report..." -ForegroundColor Cyan
try {
    $Gen = Invoke-RestMethod -Uri "$BaseUrl/report/$ReportId/generate" -Method Post -Headers $Headers -Body "" -ErrorAction Stop
    
    if ($Gen.result -eq "success") {
        $RequestId = $Gen.requestId
        Write-Host "      Request ID: $RequestId" -ForegroundColor Green
    } else {
        throw "API returned fail"
    }
} catch {
    $errDetails = $_.ErrorDetails.Message | ConvertFrom-Json -ErrorAction SilentlyContinue
    
    if ($errDetails -and $errDetails.code -eq 9004) {
        Write-Host "`n[!] Rate limit: 1 request per 10 minutes" -ForegroundColor Yellow
        Write-Host "    Enter SaveID (number): " -ForegroundColor Cyan -NoNewline
        $inputId = Read-Host
        
        if ([string]::IsNullOrWhiteSpace($inputId)) { exit }
        
        if ($inputId -match '^\d+$') {
            $SaveId = $inputId
            $SkipGeneration = $true
            Write-Host "      Using SaveID: $SaveId" -ForegroundColor Green
        } else {
            Write-Host "ERROR: Must be a number" -ForegroundColor Red
            Read-Host; exit
        }
    } else {
        Write-Host "ERROR: $_" -ForegroundColor Red
        Read-Host; exit
    }
}

# Step 2: Wait (Invoke-RestMethod - работает стабильно)
if (-not $SkipGeneration) {
    Write-Host "[2/3] Waiting..." -ForegroundColor Cyan
    $attempts = 0
    
    do {
        try {
            $Status = Invoke-RestMethod -Uri "$BaseUrl/report/status/$RequestId" -Headers $Headers
            
            if ($Status.result -eq "success" -and $Status.status -eq "ready") {
                $SaveId = $Status.save.id
                Write-Host "      Ready! SaveID: $SaveId" -ForegroundColor Green
                break
            } elseif ($Status.status -eq "error") {
                Write-Host "      Failed" -ForegroundColor Red; Read-Host; exit
            }
        } catch {}
        
        if ($attempts -eq 0) { Write-Host "      Processing..." -NoNewline -ForegroundColor Yellow } 
        else { Write-Host "." -NoNewline -ForegroundColor Yellow }
        
        Start-Sleep 2; $attempts++
    } while ($attempts -lt 60)
    
    if (-not $SaveId) { Write-Host "`nERROR: Timeout" -ForegroundColor Red; Read-Host; exit }
}

# Step 3: Download with WebClient (UTF-8 guaranteed)
Write-Host "`n[3/3] Downloading..." -ForegroundColor Cyan
$AllRows = @()
$chunk = 0

# Create WebClient with UTF-8 encoding
$wc = New-Object System.Net.WebClient
$wc.Headers.Add("Authorization", $Token)
$wc.Headers.Add("Accept", "application/json")
$wc.Encoding = [System.Text.Encoding]::UTF8

while ($chunk -lt 100) {
    $url = "$BaseUrl/report/$ReportId/save/$SaveId/data?chunk=$chunk"
    Write-VerboseLog "      [GET] $url" "DarkGray"
    
    try {
        # WebClient.UploadString with UTF8 encoding
        $JsonText = $wc.UploadString($url, "POST", "")
        $Resp = $JsonText | ConvertFrom-Json
        
        Write-VerboseLog "      [RESP] Rows: $($Resp.data.rows.Count)" "DarkCyan"
        
        if ($Resp.result -eq "success" -and $Resp.data.rows.Count -gt 0) {
            $AllRows += $Resp.data.rows
            Write-Host "      Chunk $chunk : $($Resp.data.rows.Count) rows" -ForegroundColor Gray
            $chunk++
        } else {
            break
        }
    } catch {
        Write-Host "      Stopped at chunk $chunk" -ForegroundColor Gray
        Write-VerboseLog "      [ERROR] $_" "Red"
        break
    }
}

if ($AllRows.Count -eq 0) {
    Write-Host "ERROR: No data downloaded" -ForegroundColor Red
    Read-Host; exit
}

Write-Host "      Total: $($AllRows.Count) rows" -ForegroundColor Green

# Parse
Write-Host "      Processing..." -ForegroundColor Gray
$HeaderRow = $AllRows | Where-Object { $_.type -eq "Header" } | Select-Object -First 1
$DataRows  = $AllRows | Where-Object { $_.type -eq "Normal" }

if (-not $HeaderRow) { Write-Host "ERROR: No header" -ForegroundColor Red; Read-Host; exit }

$ColumnNames = $HeaderRow.items | ForEach-Object { $_.text }

$DataArray = $DataRows | ForEach-Object {
    $row = $_
    $obj = [ordered]@{}
    for ($i = 0; $i -lt $ColumnNames.Count; $i++) {
        $val = if ($row.items[$i]) { $row.items[$i].text } else { "" }
        $val = $val -replace "`r`n", " " -replace "`n", " " -replace "`r", " "
        $obj[$ColumnNames[$i]] = $val
    }
    [PSCustomObject]$obj
}

# Save to Excel (русские буквы будут корректны благодаря UTF8 в WebClient)
Write-Host "      Saving to Excel..." -ForegroundColor Gray
try {
    $Excel = New-Object -ComObject Excel.Application
    $Excel.Visible = $false
    $Excel.DisplayAlerts = $false
    
    if ($UpdateExisting -and (Test-Path $ExcelFile)) {
        $Workbook = $Excel.Workbooks.Open($ExcelFile)
    } else {
        $Workbook = $Excel.Workbooks.Add()
    }
    
    $Sheet = $null
    if ($UpdateExisting) {
        foreach ($s in $Workbook.Sheets) { 
            if ($s.Name -eq $SheetName) { $Sheet = $s; $Sheet.Cells.Clear(); break } 
        }
    }
    
    if (-not $Sheet) {
        if ($Workbook.Sheets.Count -eq 1 -and -not $UpdateExisting) {
            $Sheet = $Workbook.Sheets.Item(1)
            $Sheet.Name = $SheetName
        } else {
            $Sheet = $Workbook.Sheets.Add()
            $Sheet.Name = $SheetName
        }
    }
    
    for ($c = 0; $c -lt $ColumnNames.Count; $c++) {
        $Cell = $Sheet.Cells(1, $c + 1)
        $Cell.Value = $ColumnNames[$c]
        $Cell.Font.Bold = $true
        $Cell.Font.Size = 11
    }
    
    for ($r = 0; $r -lt $DataArray.Count; $r++) {
        for ($c = 0; $c -lt $ColumnNames.Count; $c++) {
            $Sheet.Cells($r + 2, $c + 1) = $DataArray[$r].($ColumnNames[$c])
        }
    }
    
    $Sheet.UsedRange.Columns.AutoFit()
    
    if (Test-Path $ExcelFile) { $Workbook.Save() } else { $Workbook.SaveAs($ExcelFile) }
    
    $Workbook.Close()
    $Excel.Quit()
    
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Sheet) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Workbook) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) | Out-Null
    
    Write-Host "`nSUCCESS!" -ForegroundColor Green
    Write-Host "File: $ExcelFile" -ForegroundColor Cyan
    Write-Host "Rows: $($DataArray.Count)" -ForegroundColor Cyan
    
    Invoke-Item $ExcelFile
    
} catch {
    Write-Host "`nERROR Excel: $_" -ForegroundColor Red
}

Read-Host "`nPress Enter to exit"