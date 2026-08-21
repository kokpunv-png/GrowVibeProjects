<#
  Прототип MVP по PRD "Мониторинг новых торговых точек и точек услуг в Акмолинской области"
  Шаг 1 из 3: СБОР.

  Источники:
    - Региональные СМИ (реальные, через их встроенный поиск в RSS-фиде WordPress):
        kokshetoday.kz, akmolinform.kz, apgazeta.kz
    - Google News RSS — как дополнительный широкий источник (бесплатный, без ключей API).
      ВАЖНО: фид Google News по своим условиям предназначен для личного некоммерческого
      использования в feed-ридере (см. <copyright> в самом фиде) — для продакшн-рассылки
      внешним получателям нужно либо убрать этот источник, либо согласовать использование.

  Что делает:
    1. Собирает свежие статьи по набору поисковых запросов про открытие точек торговли/услуг.
    2. Извлекает эвристикой (regex, без ML/LLM): категорию, населённый пункт, дату, источник,
       уровень достоверности.
    3. Дедуплицирует записи (по ссылке и по нормализованному заголовку).
    4. Сохраняет CSV в ..\output\pending_review_YYYY-MM.csv с пустой колонкой "Статус" —
       её нужно заполнить вручную (да/нет) перед шагом 2 (approve_report.ps1).

  Ограничения (для первой итерации прототипа):
    - Госреестры и официальный сайт акимата (akmo.gov.kz) пока не подключены: у akmo.gov.kz
      при обращении рвётся TLS-цепочка сертификата — просто отключать проверку сертификата
      для гос.сайта показалось небезопасным решением "по-тихому", нужно разбираться отдельно
      (возможно, проблема на стороне цепочки корневых сертификатов на этой машине).
    - Адрес (улица/дом) не извлекается — колонка оставлена пустой.
    - "Название" — черновое, не выделено NLP из заголовка.
    - Ручная модерация (заполнение "Статус") обязательна — см. следующий скрипт approve_report.ps1.

  Запуск:
    powershell -File .\collect_report.ps1
#>

param(
  [int]$MaxResultsPerQuery = 30
)

$ErrorActionPreference = 'Stop'

# ---- Конфигурация ----

$UserAgent = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36'

# Населённые пункты/районы Акмолинской области для распознавания в тексте
$Settlements = @(
  'Кокшетау', 'Степногорск', 'Атбасар', 'Ерейментау', 'Щучинск', 'Макинск', 'Косшы', 'Акколь',
  'Державинск', 'Балкашино', 'Аршалы', 'Жаксы', 'Есиль', 'Шортанды', 'Бурабай', 'Боровое',
  'Астраханка', 'Целиноград', 'Акмолинск'
)

# Категория => regex ключевых слов
$Categories = [ordered]@{
  'Торговля' = 'магазин|супермаркет|гипермаркет|торгов\w*\s*(дом|центр)|\bТЦ\b|\bТРЦ\b|бутик|минимаркет'
  'Общепит'  = 'кафе|ресторан|кофейн\w*|пекарн\w*|кондитерск\w*|фастфуд|столов\w*'
  'Услуги'   = 'салон красоты|парикмахерск\w*|\bСТО\b|автосервис|автомойк\w*|аптек\w*|стоматолог\w*|фитнес|спортзал|клиник\w*|автосалон|барбершоп'
}

$OpeningRegex = 'открыл\w*|начал\w*\s+работ\w*|запуст\w*|состоял\w*\s+открытие'

# Известные HTML-сущности, которые часто встречаются в WordPress RSS-заголовках,
# но не являются валидными XML-сущностями (ломают [xml]-парсинг). amp/lt/gt/quot/apos не трогаем.
$HtmlEntityFix = [ordered]@{
  'laquo'  = [char]0x00AB
  'raquo'  = [char]0x00BB
  'mdash'  = [char]0x2014
  'ndash'  = [char]0x2013
  'hellip' = [char]0x2026
  'nbsp'   = ' '
  'ldquo'  = [char]0x201C
  'rdquo'  = [char]0x201D
  'lsquo'  = [char]0x2018
  'rsquo'  = [char]0x2019
  'deg'    = [char]0x00B0
  'copy'   = [char]0x00A9
  'reg'    = [char]0x00AE
  'trade'  = [char]0x2122
  'bull'   = [char]0x2022
  'middot' = [char]0x00B7
}

function Repair-XmlEntities {
  param([string]$Text)
  $fixed = $Text
  foreach ($name in $HtmlEntityFix.Keys) {
    $fixed = $fixed -replace "&$name;", $HtmlEntityFix[$name]
  }
  return $fixed
}

function Get-RssItems {
  param([string]$Url, [string]$Label)
  try {
    $resp = Invoke-WebRequest -Uri $Url -TimeoutSec 25 -UseBasicParsing -Headers @{ 'User-Agent' = $UserAgent }
    $fixedContent = Repair-XmlEntities -Text $resp.Content
    [xml]$xml = $fixedContent
    return $xml.rss.channel.item
  } catch {
    Write-Warning "Источник недоступен ($Label): $($_.Exception.Message)"
    return @()
  }
}

# Google News — широкие запросы с указанием региона/города прямо в запросе
$GoogleQueries = @(
  'открылся магазин Кокшетау',
  'новый магазин Акмолинская область',
  'открылось кафе Кокшетау',
  'новый ресторан Акмолинская область',
  'открылась аптека Акмолинская область',
  'открылся салон красоты Кокшетау',
  'СТО открылся Акмолинская область',
  'ТРЦ открылся Акмолинская область',
  'открылся магазин Степногорск',
  'открылся магазин Атбасар'
)

# Региональные сайты — сами уже про Акмолинскую область, поэтому запросы проще
$RegionalSites = @(
  @{ Name = 'Кокшетау Сегодня';   FeedBase = 'https://kokshetoday.kz/feed/?s=' }
  @{ Name = 'Akmol Inform';       FeedBase = 'https://akmolinform.kz/feed/?s=' }
  @{ Name = 'Акмолинская правда'; FeedBase = 'https://apgazeta.kz/feed/?s=' }
)
$SiteQueries = @(
  'открылся магазин',
  'открылось кафе',
  'новая аптека',
  'открылся салон красоты',
  'СТО открылся',
  'ТРЦ открылся',
  'открылся ресторан',
  'начал работу магазин'
)

$raw = New-Object System.Collections.Generic.List[object]

Write-Host '--- Google News RSS ---'
foreach ($q in $GoogleQueries) {
  Write-Host "Запрос: $q"
  $enc = [uri]::EscapeDataString($q)
  $items = Get-RssItems -Url "https://news.google.com/rss/search?q=$enc&hl=ru&gl=KZ&ceid=KZ:ru" -Label "Google News: $q"
  $count = 0
  foreach ($it in $items) {
    if ($count -ge $MaxResultsPerQuery) { break }
    $raw.Add([pscustomobject]@{
      SourceOverride = $null   # источник вытащим из хвоста заголовка " - Источник"
      Title          = [string]$it.title
      Link           = [string]$it.link
      PubDate        = [string]$it.pubDate
    })
    $count++
  }
  Start-Sleep -Milliseconds 500
}

Write-Host ''
Write-Host '--- Региональные СМИ ---'
foreach ($site in $RegionalSites) {
  foreach ($q in $SiteQueries) {
    Write-Host "$($site.Name): $q"
    $enc = [uri]::EscapeDataString($q)
    $items = Get-RssItems -Url "$($site.FeedBase)$enc" -Label "$($site.Name): $q"
    $count = 0
    foreach ($it in $items) {
      if ($count -ge $MaxResultsPerQuery) { break }
      $raw.Add([pscustomobject]@{
        SourceOverride = $site.Name
        Title          = [string]$it.title
        Link           = [string]$it.link
        PubDate        = [string]$it.pubDate
      })
      $count++
    }
    Start-Sleep -Milliseconds 500
  }
}

Write-Host ''
Write-Host "Собрано сырых записей: $($raw.Count)"

# ---- Извлечение полей ----

function Get-Category {
  param([string]$Text)
  foreach ($key in $Categories.Keys) {
    if ($Text -imatch $Categories[$key]) { return $key }
  }
  return $null
}

function Get-Settlement {
  param([string]$Text)
  foreach ($s in $Settlements) {
    if ($Text -imatch [regex]::Escape($s)) { return $s }
  }
  return $null
}

function Get-SourceFromTitle {
  param([string]$Title)
  if ($Title -match '\s-\s([^-]+)$') { return $Matches[1].Trim() }
  return 'не определён'
}

function Get-CleanTitle {
  param([string]$Title, [bool]$StripSourceSuffix)
  if ($StripSourceSuffix) {
    return ($Title -replace '\s-\s[^-]+$', '').Trim()
  }
  return $Title.Trim()
}

$structured = New-Object System.Collections.Generic.List[object]

foreach ($r in $raw) {
  $category = Get-Category -Text $r.Title
  if (-not $category) { continue }  # нерелевантно — не про торговлю/услуги/общепит

  $settlement = Get-Settlement -Text $r.Title
  $hasOpeningVerb = $r.Title -imatch $OpeningRegex

  $confidenceScore = 0
  if ($category) { $confidenceScore++ }
  if ($settlement) { $confidenceScore++ }
  if ($hasOpeningVerb) { $confidenceScore++ }
  # Собственные региональные СМИ по умолчанию точнее бьют по региону, чем общий агрегатор
  if ($r.SourceOverride) { $confidenceScore++ }

  $confidence = switch ($confidenceScore) {
    { $_ -ge 3 } { 'высокий' }
    2            { 'средний' }
    default      { 'низкий' }
  }

  $date = 'не определена'
  try {
    $dt = [datetime]::Parse($r.PubDate, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None)
    $date = $dt.ToString('yyyy-MM-dd')
  } catch {}

  $source = if ($r.SourceOverride) { $r.SourceOverride } else { Get-SourceFromTitle -Title $r.Title }
  $cleanTitle = Get-CleanTitle -Title $r.Title -StripSourceSuffix ([bool]($null -eq $r.SourceOverride))

  $normTitle = ($cleanTitle -replace '[^\p{L}\p{N}]', '').ToLower()

  $structured.Add([pscustomobject]@{
    'Название (черновое)'     = $cleanTitle
    'Категория'                = $category
    'Населённый пункт/район'   = $(if ($settlement) { $settlement } else { 'не определён' })
    'Адрес'                    = ''
    'Дата публикации'          = $date
    'Источник'                 = $source
    'Ссылка'                   = $r.Link
    'Достоверность'            = $confidence
    'Статус'                   = ''   # заполняется вручную: да / нет
    '_NormTitle'               = $normTitle
  })
}

Write-Host "После фильтра по категории: $($structured.Count)"

# ---- Дедупликация (по ссылке и по нормализованному заголовку) ----

$seenLinks = New-Object System.Collections.Generic.HashSet[string]
$seenTitles = New-Object System.Collections.Generic.HashSet[string]
$deduped = New-Object System.Collections.Generic.List[object]

foreach ($item in $structured) {
  $key = $item._NormTitle.Substring(0, [Math]::Min(60, $item._NormTitle.Length))
  if ($seenLinks.Contains($item.'Ссылка') -or $seenTitles.Contains($key)) { continue }
  [void]$seenLinks.Add($item.'Ссылка')
  [void]$seenTitles.Add($key)
  $deduped.Add($item)
}

$dupRemoved = $structured.Count - $deduped.Count
Write-Host "Удалено дублей: $dupRemoved"
Write-Host "Итоговых записей: $($deduped.Count)"

# ---- Экспорт ----

$outDir = Join-Path $PSScriptRoot '..\output'
if (-not (Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir | Out-Null }
$stamp = Get-Date -Format 'yyyy-MM'
$outFile = Join-Path $outDir "pending_review_$stamp.csv"

$deduped |
  Select-Object 'Название (черновое)', 'Категория', 'Населённый пункт/район', 'Адрес', 'Дата публикации', 'Источник', 'Ссылка', 'Достоверность', 'Статус' |
  Export-Csv -Path $outFile -NoTypeInformation -Encoding UTF8 -Delimiter ';'

Write-Host ''
Write-Host "Файл на модерацию сохранён: $outFile"
Write-Host 'Откройте его (Excel/Numbers), для каждой строки заполните колонку "Статус": да / нет.'
Write-Host 'Затем запустите approve_report.ps1 — он соберёт финальный отчёт только из подтверждённых строк.'

# ---- Сводная статистика ----
Write-Host ''
Write-Host '=== Сводка по категориям ==='
$deduped | Group-Object 'Категория' | Sort-Object Count -Descending | ForEach-Object { Write-Host "$($_.Name): $($_.Count)" }

Write-Host ''
Write-Host '=== Сводка по населённым пунктам ==='
$deduped | Group-Object 'Населённый пункт/район' | Sort-Object Count -Descending | ForEach-Object { Write-Host "$($_.Name): $($_.Count)" }

Write-Host ''
Write-Host '=== Сводка по источникам ==='
$deduped | Group-Object 'Источник' | Sort-Object Count -Descending | ForEach-Object { Write-Host "$($_.Name): $($_.Count)" }

Write-Host ''
Write-Host '=== Сводка по достоверности ==='
$deduped | Group-Object 'Достоверность' | Sort-Object Count -Descending | ForEach-Object { Write-Host "$($_.Name): $($_.Count)" }
