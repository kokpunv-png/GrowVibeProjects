<#
  Шаг 2 из 3: МОДЕРАЦИЯ.
  PRD, п.7 (нефункциональные требования): "Возможность ручной модерации записей перед отправкой
  отчёта — на начальном этапе обязательна, т.к. NLP-извлечение может ошибаться".

  Берёт output\pending_review_YYYY-MM.csv (создаётся collect_report.ps1), в котором человек
  вручную заполнил колонку "Статус" значением "да" или "нет" для каждой строки, и формирует
  финальный отчёт только из подтверждённых строк: output\report_YYYY-MM_final.csv

  Запуск:
    powershell -File .\approve_report.ps1
    powershell -File .\approve_report.ps1 -Month 2026-08
#>

param(
  [string]$Month = (Get-Date -Format 'yyyy-MM')
)

$ErrorActionPreference = 'Stop'

$outDir = Join-Path $PSScriptRoot '..\output'
$pendingFile = Join-Path $outDir "pending_review_$Month.csv"
$finalFile = Join-Path $outDir "report_${Month}_final.csv"

if (-not (Test-Path $pendingFile)) {
  Write-Error "Файл на модерации не найден: $pendingFile`nСначала запустите collect_report.ps1."
  exit 1
}

$rows = Import-Csv -Path $pendingFile -Delimiter ';' -Encoding UTF8

$approved = New-Object System.Collections.Generic.List[object]
$rejected = 0
$pending = 0

foreach ($row in $rows) {
  $status = ($row.'Статус' -as [string]).Trim().ToLower()
  if ($status -in @('да', 'yes', '1', 'y', 'true')) {
    $approved.Add($row)
  } elseif ($status -in @('нет', 'no', '0', 'n', 'false')) {
    $rejected++
  } else {
    $pending++
  }
}

if ($pending -gt 0) {
  Write-Warning "$pending запись(ей) со статусом ещё не заполнены (не 'да'/'нет') — они НЕ попали в финальный отчёт."
}

$approved | Export-Csv -Path $finalFile -NoTypeInformation -Encoding UTF8 -Delimiter ';'

Write-Host ''
Write-Host "Всего строк на модерации: $($rows.Count)"
Write-Host "Подтверждено (да): $($approved.Count)"
Write-Host "Отклонено (нет): $rejected"
Write-Host "Не промодерировано: $pending"
Write-Host ''
Write-Host "Финальный отчёт сохранён: $finalFile"
Write-Host 'Дальше: send_report.ps1 — доставка на email/Telegram.'
