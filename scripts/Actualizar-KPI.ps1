<#
.SYNOPSIS
  Actualiza el KPI "Tasa de Instalacion BAF" con los archivos fuente de un nuevo periodo,
  acumula el resultado en el historico mensual, regenera el dashboard y el reporte Excel,
  y opcionalmente hace commit/push a GitHub.

.PARAMETER ArchivoTOA
  Ruta al archivo pXX_TOA_ELECNOR_AAAA-MM.xlsx del periodo a cargar.

.PARAMETER ArchivoMaestroVentas
  Ruta al archivo Maestro_Ventas_Elecnor_AAAA-MM.xlsx del periodo a cargar.

.PARAMETER Periodo
  Etiqueta del periodo en formato AAAA-MM, por ejemplo "2026-09".

.PARAMETER Meta
  Meta de tasa de instalacion en porcentaje. Por defecto 80.7.

.PARAMETER Push
  Si se incluye, hace "git push origin master" despues del commit.

.EXAMPLE
  .\Actualizar-KPI.ps1 -ArchivoTOA "C:\Users\ximel\Downloads\p13_TOA_ELECNOR_2026-09.xlsx" `
                        -ArchivoMaestroVentas "C:\Users\ximel\Downloads\Maestro_Ventas_Elecnor_2026-09.xlsx" `
                        -Periodo "2026-09" -Push
#>
param(
    [Parameter(Mandatory = $true)][string]$ArchivoTOA,
    [Parameter(Mandatory = $true)][string]$ArchivoMaestroVentas,
    [Parameter(Mandatory = $true)][ValidatePattern('^\d{4}-\d{2}$')][string]$Periodo,
    [double]$Meta = 80.7,
    [switch]$Push
)

$ErrorActionPreference = "Stop"

$repoRoot     = Split-Path -Parent $PSScriptRoot
$historicoPath = Join-Path $repoRoot "historico\kpi_historico.csv"
$periodoDir    = Join-Path $repoRoot "periodos\$Periodo"
$templatePath  = Join-Path $PSScriptRoot "dashboard_template.html"

New-Item -ItemType Directory -Force -Path (Join-Path $periodoDir "datos")     | Out-Null
New-Item -ItemType Directory -Force -Path (Join-Path $periodoDir "reportes") | Out-Null
New-Item -ItemType Directory -Force -Path (Join-Path $periodoDir "graficos") | Out-Null
New-Item -ItemType Directory -Force -Path (Split-Path $historicoPath)       | Out-Null

function Stop-StrayExcel {
    Get-Process -Name EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
}

Write-Output "== 1/6 Exportando archivos fuente a CSV =="
Stop-StrayExcel
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$tmpToa = Join-Path $env:TEMP "toa_$Periodo.csv"
$wbToa = $excel.Workbooks.Open($ArchivoTOA, $null, $true)
$wbToa.Sheets.Item(1).SaveAs($tmpToa, 6)   # 6 = xlCSV
$wbToa.Close($false)

$tmpMv = Join-Path $env:TEMP "mv_$Periodo.csv"
$wbMv = $excel.Workbooks.Open($ArchivoMaestroVentas, $null, $true)
$wbMv.Sheets.Item(1).SaveAs($tmpMv, 6)
$wbMv.Close($false)

$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

Write-Output "== 2/6 Calculando numerador (TOA: complete + XA_WORK_TYPE_ID que inicia con AT) =="
$toaRows = Import-Csv -Path $tmpToa -Delimiter ';'
$numerador = @{}
$numeradorDetalle = New-Object System.Collections.Generic.List[object]
foreach ($row in $toaRows) {
    $status = ("$($row.status)").Trim().ToLower()
    $wt = "$($row.xa_work_type)"
    $agencia = ("$($row.xa_original_agency)").Trim().ToUpper()
    if (-not $agencia) { continue }
    if ($status -eq 'complete' -and $wt.Length -ge 2 -and $wt.Substring(0, 2).ToUpper() -eq 'AT') {
        if (-not $numerador.ContainsKey($agencia)) { $numerador[$agencia] = 0 }
        $numerador[$agencia]++
        $numeradorDetalle.Add([PSCustomObject]@{
            appt_number = $row.appt_number; xa_original_agency = $agencia
            xa_work_type = $wt; status = $status; fecha = $row.fecha
        })
    }
}

Write-Output "== 3/6 Calculando denominador (Maestro_Ventas: Emision - Cancelada) =="
$mvRows = Import-Csv -Path $tmpMv -Delimiter ';'
$emision = @{}
$cancelada = @{}
foreach ($row in $mvRows) {
    $tipo = ("$($row.TIPO)").Trim().ToLower()
    $agencia = ("$($row.AGENCIA_ZONAL_INSTALACION)").Trim().ToUpper()
    if (-not $agencia) { continue }
    if ($tipo -eq 'emision') {
        if (-not $emision.ContainsKey($agencia)) { $emision[$agencia] = 0 }
        $emision[$agencia]++
    } elseif ($tipo -eq 'cancelada') {
        if (-not $cancelada.ContainsKey($agencia)) { $cancelada[$agencia] = 0 }
        $cancelada[$agencia]++
    }
}

$agencias = @($numerador.Keys) + @($emision.Keys) + @($cancelada.Keys) | Select-Object -Unique | Sort-Object
$resultados = foreach ($ag in $agencias) {
    $num = if ($numerador.ContainsKey($ag)) { $numerador[$ag] } else { 0 }
    $em  = if ($emision.ContainsKey($ag))   { $emision[$ag] }   else { 0 }
    $can = if ($cancelada.ContainsKey($ag)) { $cancelada[$ag] } else { 0 }
    $den = $em - $can
    $tasa = if ($den -ne 0) { [math]::Round(($num / $den) * 100, 1) } else { 0 }
    $brecha = [math]::Round($tasa - $Meta, 1)
    $estado =
        if ($den -eq 0) { 'sin_datos' }
        elseif ($tasa -ge 100) { 'verificar_dato' }
        elseif ($tasa -ge $Meta) { 'cumple' }
        elseif ($tasa -ge ($Meta - 5)) { 'cerca_meta' }
        else { 'critico' }
    [PSCustomObject]@{
        periodo = $Periodo; agencia = $ag; numerador = $num; emision = $em; cancelada = $can
        denominador = $den; tasa = $tasa; meta = $Meta; brecha = $brecha; estado = $estado
    }
}

Write-Output ($resultados | Format-Table -AutoSize | Out-String)

Write-Output "== 4/6 Guardando detalle del periodo y actualizando historico =="
$numeradorDetalle | Export-Csv -Path (Join-Path $periodoDir "datos\numerador_filtrado.csv") -NoTypeInformation -Delimiter ';' -Encoding UTF8
$mvRows | Select-Object TIPO, AGENCIA_ZONAL_INSTALACION, ORDEN, FECHA_EMISION |
    Export-Csv -Path (Join-Path $periodoDir "datos\denominador_maestro_ventas.csv") -NoTypeInformation -Delimiter ';' -Encoding UTF8

if (Test-Path $historicoPath) {
    $existing = @(Import-Csv $historicoPath | Where-Object { $_.periodo -ne $Periodo })
} else {
    $existing = @()
}
$historicoAll = @($existing) + @($resultados) | Sort-Object periodo, agencia
$historicoAll | Export-Csv -Path $historicoPath -NoTypeInformation -Encoding UTF8

Write-Output "== 5/6 Regenerando dashboard.html / index.html con el historico completo =="
$jsonRows = $historicoAll | ForEach-Object {
    '{"periodo":"' + $_.periodo + '","agencia":"' + $_.agencia + '","numerador":' + $_.numerador +
    ',"emision":' + $_.emision + ',"cancelada":' + $_.cancelada + ',"denominador":' + $_.denominador +
    ',"tasa":' + $_.tasa + ',"meta":' + $_.meta + ',"brecha":' + $_.brecha + ',"estado":"' + $_.estado + '"}'
}
$json = "[`n" + ($jsonRows -join ",`n") + "`n]"
$template = Get-Content -Path $templatePath -Raw -Encoding UTF8
$dashboardHtml = $template.Replace('__KPI_DATA_JSON__', $json)
Set-Content -Path (Join-Path $repoRoot "dashboard.html") -Value $dashboardHtml -Encoding UTF8 -NoNewline
Copy-Item (Join-Path $repoRoot "dashboard.html") (Join-Path $repoRoot "index.html") -Force

Write-Output "== 6/6 Regenerando reporte Excel del periodo =="
Stop-StrayExcel
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$wb = $excel.Workbooks.Add()
$ws1 = $wb.Sheets.Item(1)
$ws1.Name = "Resumen"
$ws1.Cells.Item(1,1) = "Tasa de Instalacion BAF - ELECNOR - Periodo $Periodo"
$ws1.Cells.Item(1,1).Font.Bold = $true
$ws1.Cells.Item(1,1).Font.Size = 14
$ws1.Cells.Item(2,1) = "Numerador: TOA, STATUS_CD=complete y XA_WORK_TYPE_ID que inicia con AT (BAF). Denominador: Maestro_Ventas, Emision - Cancelada por AGENCIA_ZONAL_INSTALACION."

$headers = @("Agencia","Numerador (Complete+AT)","Emision","Cancelada","Denominador (Neto)","Tasa Instalacion (%)","Meta (%)","Brecha (%)","Cumple Meta")
for ($c=0; $c -lt $headers.Length; $c++) {
    $cell = $ws1.Cells.Item(4, $c+1)
    $cell.Value2 = $headers[$c]
    $cell.Font.Bold = $true
    $cell.Interior.Color = 15773696
    $cell.Font.Color = 16777215
}

$r = 5
foreach ($row in $resultados) {
    $ws1.Cells.Item($r,1) = $row.agencia
    $ws1.Cells.Item($r,2) = $row.numerador
    $ws1.Cells.Item($r,3) = $row.emision
    $ws1.Cells.Item($r,4) = $row.cancelada
    $ws1.Cells.Item($r,5) = "=C$r-D$r"
    $ws1.Cells.Item($r,6) = "=B$r/E$r"
    $ws1.Cells.Item($r,7) = $Meta / 100
    $ws1.Cells.Item($r,8) = "=F$r-G$r"
    $ws1.Cells.Item($r,9) = "=IF(F$r>=G$r,""Si"",""No"")"
    $r++
}
$totalRow = $r
$ws1.Cells.Item($totalRow,1) = "TOTAL"
$ws1.Cells.Item($totalRow,1).Font.Bold = $true
$ws1.Cells.Item($totalRow,2) = "=SUM(B5:B$($totalRow-1))"
$ws1.Cells.Item($totalRow,3) = "=SUM(C5:C$($totalRow-1))"
$ws1.Cells.Item($totalRow,4) = "=SUM(D5:D$($totalRow-1))"
$ws1.Cells.Item($totalRow,5) = "=C$totalRow-D$totalRow"
$ws1.Cells.Item($totalRow,6) = "=B$totalRow/E$totalRow"
$ws1.Cells.Item($totalRow,7) = $Meta / 100
$ws1.Cells.Item($totalRow,8) = "=F$totalRow-G$totalRow"
$ws1.Cells.Item($totalRow,9) = "=IF(F$totalRow>=G$totalRow,""Si"",""No"")"
for ($c=1; $c -le 9; $c++) { $ws1.Cells.Item($totalRow,$c).Font.Bold = $true }

$ws1.Range("F5:H$totalRow").NumberFormat = "0,0%"
$ws1.Cells.Item($totalRow+2,1) = "Generado automaticamente por scripts\Actualizar-KPI.ps1"
$ws1.Columns.Item("A:I").AutoFit() | Out-Null

$chartObjs = $ws1.ChartObjects()
$co = $chartObjs.Add(50, ($totalRow+4)*15, 520, 300)
$chart = $co.Chart
$chart.ChartType = 51
$s1 = $chart.SeriesCollection().NewSeries()
$s1.Values = $ws1.Range("F5:F$($totalRow-1)")
$s1.XValues = $ws1.Range("A5:A$($totalRow-1)")
$s1.Name = "Tasa Instalacion (%)"
$s2 = $chart.SeriesCollection().NewSeries()
$s2.Values = $ws1.Range("G5:G$($totalRow-1)")
$s2.Name = "Meta (%)"
$chart.HasTitle = $true
$chart.ChartTitle.Text = "Tasa de Instalacion BAF vs Meta - $Periodo"
$chart.HasLegend = $true
$chart.Axes(2).TickLabels.NumberFormat = "0%"

$ws2 = $wb.Sheets.Add([System.Reflection.Missing]::Value, $wb.Sheets.Item($wb.Sheets.Count))
$ws2.Name = "Numerador_Detalle"
$qt = $ws2.QueryTables.Add("TEXT;" + (Join-Path $periodoDir "datos\numerador_filtrado.csv"), $ws2.Range("A1"))
$qt.TextFileParseType = 1
$qt.TextFileSemicolonDelimiter = $true
$qt.Refresh($false) | Out-Null
$ws2.Rows.Item(1).Font.Bold = $true
$ws2.Columns.Item("A:E").AutoFit() | Out-Null

$ws3 = $wb.Sheets.Add([System.Reflection.Missing]::Value, $wb.Sheets.Item($wb.Sheets.Count))
$ws3.Name = "Denominador_Detalle"
$qt2 = $ws3.QueryTables.Add("TEXT;" + (Join-Path $periodoDir "datos\denominador_maestro_ventas.csv"), $ws3.Range("A1"))
$qt2.TextFileParseType = 1
$qt2.TextFileSemicolonDelimiter = $true
$qt2.Refresh($false) | Out-Null
$ws3.Rows.Item(1).Font.Bold = $true
$ws3.Columns.Item("A:D").AutoFit() | Out-Null

$ws1.Select()
$ws1.Range("A1").Select()

$outPath = Join-Path $periodoDir "reportes\Reporte_Tasa_Instalacion_BAF_$Periodo.xlsx"
if (Test-Path $outPath) { Remove-Item $outPath -Force }
$wb.SaveAs($outPath, 51)
$chart.Export((Join-Path $periodoDir "graficos\tasa_instalacion_vs_meta.png"), "PNG") | Out-Null
$wb.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

Write-Output "`nListo. Periodo $Periodo actualizado:"
Write-Output "  - $historicoPath"
Write-Output "  - $outPath"
Write-Output "  - $(Join-Path $repoRoot 'dashboard.html') / index.html"

Push-Location $repoRoot
git add -A
$commitMsg = "Actualiza KPI Tasa de Instalacion BAF - periodo $Periodo"
git commit -m $commitMsg
if ($Push) {
    git push origin master
    Write-Output "`nPublicado en GitHub. GitHub Pages se actualiza automaticamente en ~1 minuto."
} else {
    Write-Output "`nCommit creado localmente. Ejecuta 'git push origin master' (o vuelve a correr con -Push) para publicarlo."
}
Pop-Location
