<#
  Шаг 3 из 3: ДОСТАВКА.
  PRD, п.4/п.6: доставка готового отчёта пользователю на email и/или в Telegram.

  Настройка (один раз):
    1. Скопируйте config.example.json -> config.json (в этой же папке scripts\).
    2. Заполните в config.json свои реальные данные:
         - email: SMTP-сервер, логин/пароль (для Gmail — пароль приложения, не основной пароль),
           адрес(а) получателя(ей);
         - telegram: токен бота (получить у @BotFather) и chatId получателя/канала (например,
           узнать через https://api.telegram.org/bot<токен>/getUpdates после /start боту).
       config.json НЕ передавайте никому и не публикуйте — там секреты.
    3. У каждого блока ("email"/"telegram") выставьте "enabled": true, если хотите его использовать.

  ВАЖНО: этот скрипт не запрашивает и не хранит ваши секреты за вас — их нужно вписать в
  config.json самостоятельно, локально на этой машине.

  По умолчанию скрипт запускается в режиме DRY RUN — только показывает, что было бы отправлено,
  и ничего не отправляет. Чтобы реально отправить — добавьте флаг -Execute.

  Запуск:
    powershell -File .\send_report.ps1                 # dry run
    powershell -File .\send_report.ps1 -Execute         # реальная отправка
    powershell -File .\send_report.ps1 -Month 2026-08 -Execute
#>

param(
  [string]$Month = (Get-Date -Format 'yyyy-MM'),
  [switch]$Execute
)

$ErrorActionPreference = 'Stop'

$scriptDir = $PSScriptRoot
$outDir = Join-Path $scriptDir '..\output'
$finalFile = Join-Path $outDir "report_${Month}_final.csv"
$configFile = Join-Path $scriptDir 'config.json'
$exampleFile = Join-Path $scriptDir 'config.example.json'

if (-not (Test-Path $finalFile)) {
  Write-Error "Финальный отчёт не найден: $finalFile`nСначала запустите collect_report.ps1, затем approve_report.ps1."
  exit 1
}

if (-not (Test-Path $configFile)) {
  Write-Warning "Файл config.json не найден."
  Write-Host "Скопируйте $exampleFile в $configFile и заполните SMTP/Telegram настройки."
  exit 1
}

$config = Get-Content -Path $configFile -Raw -Encoding UTF8 | ConvertFrom-Json

if (-not $Execute) {
  Write-Host '=== DRY RUN (ничего не отправляется) ==='
  Write-Host "Файл отчёта: $finalFile"
}

# ---- Email ----

function Send-EmailReport {
  param($EmailConfig, [string]$AttachmentPath, [bool]$DryRunOnly)

  if (-not $EmailConfig.enabled) {
    Write-Host 'Email: пропущено (enabled=false в config.json).'
    return
  }
  $toList = @($EmailConfig.to)
  Write-Host "Email: получатели -> $($toList -join ', ')"
  Write-Host "Email: тема -> $($EmailConfig.subject)"

  if ($DryRunOnly) {
    Write-Host 'Email: (dry run) письмо не отправлено. Запустите с -Execute для реальной отправки.'
    return
  }

  if (-not $EmailConfig.username -or -not $EmailConfig.password) {
    Write-Warning 'Email: не заполнены username/password в config.json — отправка невозможна.'
    return
  }

  $msg = New-Object System.Net.Mail.MailMessage
  $msg.From = $(if ($EmailConfig.from) { $EmailConfig.from } else { $EmailConfig.username })
  foreach ($t in $toList) { $msg.To.Add($t) }
  $msg.Subject = $EmailConfig.subject
  $msg.Body = "Автоматический отчёт о новых точках торговли и услуг в Акмолинской области за $Month. См. вложение (CSV)."
  $attachment = New-Object System.Net.Mail.Attachment($AttachmentPath)
  $msg.Attachments.Add($attachment)

  $smtp = New-Object System.Net.Mail.SmtpClient($EmailConfig.smtpHost, [int]$EmailConfig.smtpPort)
  $smtp.EnableSsl = [bool]$EmailConfig.useSsl
  $smtp.Credentials = New-Object System.Net.NetworkCredential($EmailConfig.username, $EmailConfig.password)

  try {
    $smtp.Send($msg)
    Write-Host 'Email: отправлено успешно.'
  } catch {
    Write-Warning "Email: ошибка отправки: $($_.Exception.Message)"
  } finally {
    $attachment.Dispose()
    $msg.Dispose()
  }
}

# ---- Telegram ----

function Send-TelegramReport {
  param($TelegramConfig, [string]$AttachmentPath, [bool]$DryRunOnly)

  if (-not $TelegramConfig.enabled) {
    Write-Host 'Telegram: пропущено (enabled=false в config.json).'
    return
  }
  Write-Host "Telegram: chatId -> $($TelegramConfig.chatId)"

  if ($DryRunOnly) {
    Write-Host 'Telegram: (dry run) документ не отправлен. Запустите с -Execute для реальной отправки.'
    return
  }

  if (-not $TelegramConfig.botToken -or -not $TelegramConfig.chatId) {
    Write-Warning 'Telegram: не заполнены botToken/chatId в config.json — отправка невозможна.'
    return
  }

  Add-Type -AssemblyName System.Net.Http

  $httpClient = New-Object System.Net.Http.HttpClient
  try {
    $url = "https://api.telegram.org/bot$($TelegramConfig.botToken)/sendDocument"
    $content = New-Object System.Net.Http.MultipartFormDataContent

    $chatIdContent = New-Object System.Net.Http.StringContent([string]$TelegramConfig.chatId)
    $content.Add($chatIdContent, 'chat_id')

    $captionContent = New-Object System.Net.Http.StringContent("Отчёт о новых точках торговли и услуг в Акмолинской области за $Month")
    $content.Add($captionContent, 'caption')

    $fileBytes = [System.IO.File]::ReadAllBytes($AttachmentPath)
    $fileContent = New-Object System.Net.Http.ByteArrayContent($fileBytes)
    $content.Add($fileContent, 'document', (Split-Path $AttachmentPath -Leaf))

    $response = $httpClient.PostAsync($url, $content).GetAwaiter().GetResult()
    $body = $response.Content.ReadAsStringAsync().GetAwaiter().GetResult()

    if ($response.IsSuccessStatusCode) {
      Write-Host 'Telegram: отправлено успешно.'
    } else {
      Write-Warning "Telegram: ошибка отправки (HTTP $($response.StatusCode)): $body"
    }
  } catch {
    Write-Warning "Telegram: ошибка отправки: $($_.Exception.Message)"
  } finally {
    $httpClient.Dispose()
  }
}

Send-EmailReport -EmailConfig $config.email -AttachmentPath $finalFile -DryRunOnly (-not $Execute)
Send-TelegramReport -TelegramConfig $config.telegram -AttachmentPath $finalFile -DryRunOnly (-not $Execute)
