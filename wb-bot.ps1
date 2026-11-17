[CmdletBinding()]
param(
    [ValidateSet('Setup', 'SetupTelegram', 'TestTelegram', 'Check', 'Watch',
        'ExportTransfer', 'ImportTransfer', 'EnableStartup', 'DisableStartup', 'Menu', 'SetupOzon', 'CheckOzon',
        'StartBackground', 'StopBackground', 'ShowBackgroundLog')]
    [string]$Mode = 'Watch',

    [ValidateRange(1, 1440)]
    [int]$IntervalMinutes = 30,

    [switch]$SendEveryCheck,

    [switch]$Background
)

$tokenDirectory = Join-Path $env:LOCALAPPDATA 'WB Feedback Bot'
$backgroundStopPath = Join-Path $tokenDirectory 'background.stop'
$backgroundLogPath = Join-Path $tokenDirectory 'background.log'
$tokenPath = Join-Path $tokenDirectory 'token.dpapi'
$telegramPath = Join-Path $tokenDirectory 'telegram.json'
$ozonPath = Join-Path $tokenDirectory 'ozon.json'
$configPath = Join-Path $PSScriptRoot 'config.json'
$transferPath = Join-Path $PSScriptRoot 'wb-transfer.json'
$feedbackUri = 'https://feedbacks-api.wildberries.ru/api/v1/feedbacks'
$newsUri = 'https://common-api.wildberries.ru/api/communications/v2/news'

function Invoke-WbRequest {
    param(
        [System.Security.SecureString]$Token,
        [string]$Uri
    )

    $tokenPointer = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($Token)
    try {
        $tokenValue = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($tokenPointer).Trim()
        if ($tokenValue -notmatch '^[A-Za-z0-9_-]+\.[A-Za-z0-9_-]+\.[A-Za-z0-9_-]+$') {
            throw 'Ожидался полный токен WB API. Скопируйте его из WB Partners.'
        }

        try {
            $response = Invoke-WebRequest -Uri $Uri -Method Get -UseBasicParsing `
                -Headers @{ Authorization = $tokenValue; Accept = 'application/json' } `
                -TimeoutSec 20 -MaximumRedirection 0 -ErrorAction Stop
        }
        catch {
            $httpResponse = $_.Exception.Response
            $status = if ($null -eq $httpResponse) { 0 } else { [int]$httpResponse.StatusCode }
            $message = switch ($status) {
                0   { 'Сбой соединения с WB API. Проверьте интернет на ноутбуке.' }
                401 { 'WB API вернул 401. Проверьте срок действия токена и при необходимости обновите его через Setup.' }
                402 { 'WB API вернул 402. Проверьте условия доступа для выбранного токена в WB Partners.' }
                403 { 'WB API вернул 403. Доступ к запрошенным данным запрещён.' }
                429 { 'WB API вернул 429. Превышен лимит запросов.' }
                default { "WB API вернул HTTP $status." }
            }
            $failure = [InvalidOperationException]::new($message)
            $failure.Data['Status'] = $status
            if ($status -eq 429) {
                $retrySeconds = 0
                $retryHeader = $httpResponse.Headers.'X-Ratelimit-Retry'
                if ([int]::TryParse([string]$retryHeader, [ref]$retrySeconds) -and $retrySeconds -gt 0) {
                    $failure.Data['RetrySeconds'] = $retrySeconds
                }
            }
            throw $failure
        }
    }
    finally {
        [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($tokenPointer)
    }

    try {
        $payload = ConvertFrom-Json -InputObject $response.Content -ErrorAction Stop
    }
    catch {
        throw 'WB API вернул ответ, который не удалось прочитать как JSON.'
    }

    $payload
}

function Get-WbFeedbackData {
    param([securestring]$Token, [string]$Uri)

    $payload = Invoke-WbRequest -Token $Token -Uri $Uri
    if ($payload.error -isnot [bool] -or $payload.error -or $null -eq $payload.data) {
        throw 'WB API не вернул данные об отзывах.'
    }
    $payload.data
}

function Get-WbFeedbackCounts {
    param([System.Security.SecureString]$Token)

    $data = Get-WbFeedbackData -Token $Token -Uri "$feedbackUri/count-unanswered"
    $total = $data.countUnanswered
    $today = $data.countUnansweredToday
    if (($total -isnot [int] -and $total -isnot [long]) -or
        ($today -isnot [int] -and $today -isnot [long]) -or
        $total -lt 0 -or $today -lt 0) {
        throw 'В ответе WB API нет корректных счётчиков необработанных отзывов.'
    }

    [pscustomobject]@{ Total = $total; Today = $today }
}

function Get-WbUnansweredIds {
    param([System.Security.SecureString]$Token)

    $ids = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    $pageSize = 1000
    for ($skip = 0; $skip -le 199990; $skip += $pageSize) {
        Start-Sleep -Milliseconds 400
        $uri = "${feedbackUri}?isAnswered=false&take=$pageSize&skip=$skip&order=dateAsc"
        $data = Get-WbFeedbackData -Token $Token -Uri $uri
        $feedbacksProperty = $data.PSObject.Properties['feedbacks']
        if ($null -eq $feedbacksProperty -or ($null -ne $data.feedbacks -and $data.feedbacks -isnot [array])) {
            throw 'В ответе WB API нет корректного списка отзывов.'
        }

        foreach ($feedback in $data.feedbacks) {
            if ($feedback.id -isnot [string] -or [string]::IsNullOrWhiteSpace($feedback.id)) {
                throw 'WB API вернул отзыв без ID.'
            }
            [void]$ids.Add($feedback.id)
        }
        if ($data.feedbacks.Count -lt $pageSize -or ($skip + $pageSize -ge 200000 -and
                $data.countUnanswered -eq $ids.Count)) {
            return ,$ids
        }
    }
    throw 'Список необработанных отзывов превышает доступный лимит пагинации WB API.'
}

function Write-LocalFile {
    param([string]$Path, [string]$Content)

    $directory = [IO.Path]::GetDirectoryName($Path)
    [void][IO.Directory]::CreateDirectory($directory)
    $temporaryPath = Join-Path $directory ([guid]::NewGuid().ToString() + '.tmp')
    try {
        [IO.File]::WriteAllText($temporaryPath, $Content)
        if ([IO.File]::Exists($Path)) {
            [IO.File]::Replace($temporaryPath, $Path, [NullString]::Value)
        }
        else {
            [IO.File]::Move($temporaryPath, $Path)
        }
    }
    finally {
        [IO.File]::Delete($temporaryPath)
    }
}

function Invoke-TelegramRequest {
    param(
        [System.Security.SecureString]$Token,
        [string]$Method,
        [hashtable]$Parameters = @{}
    )

    $tokenPointer = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($Token)
    try {
        $tokenValue = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($tokenPointer).Trim()
        if ($tokenValue -notmatch '^[0-9]+:[A-Za-z0-9_-]+$') {
            throw 'Скопируйте полный токен бота из сообщения BotFather.'
        }

        try {
            $body = ConvertTo-Json -InputObject $Parameters -Compress
            $response = Invoke-WebRequest -Uri "https://api.telegram.org/bot$tokenValue/$Method" `
                -Method Post -UseBasicParsing -ContentType 'application/json; charset=utf-8' `
                -Body ([Text.Encoding]::UTF8.GetBytes($body)) -TimeoutSec 25 -MaximumRedirection 0 `
                -Verbose:$false -Debug:$false -ErrorAction Stop
            $payload = ConvertFrom-Json -InputObject $response.Content -ErrorAction Stop
        }
        catch {
            $httpResponse = $_.Exception.Response
            if ($null -eq $httpResponse) {
                throw 'Не удалось получить ответ от Telegram Bot API. Проверьте подключение к интернету.'
            }
            $status = [int]$httpResponse.StatusCode
            $message = switch ($status) {
                400 { 'Telegram вернул 400. Проверьте группу или тему и повторите SetupTelegram.' }
                401 { 'Telegram вернул 401. Проверьте токен BotFather и повторите SetupTelegram.' }
                403 { 'Telegram вернул 403. Проверьте, что бот добавлен в группу и может отправлять сообщения.' }
                404 { 'Telegram вернул 404. Проверьте токен BotFather и повторите SetupTelegram.' }
                409 { 'Telegram вернул 409. Для настройки используйте отдельного бота без webhook и других запущенных интеграций.' }
                429 { 'Telegram вернул 429. Дождитесь следующей проверки или повторите команду позже.' }
                default { "Telegram вернул HTTP $status." }
            }
            throw $message
        }
    }
    finally {
        [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($tokenPointer)
    }

    if ($payload.ok -ne $true -or $null -eq $payload.result) {
        throw 'Telegram Bot API не подтвердил выполнение запроса.'
    }
    return ,$payload.result
}

function Send-TelegramMessage {
    param($Telegram, [string]$Message, [switch]$DisableLinkPreview)

    $parameters = @{ chat_id = $Telegram.ChatId; text = $Message }
    if ($DisableLinkPreview) {
        $parameters.link_preview_options = @{ is_disabled = $true }
    }
    if ($Telegram.ThreadId -gt 1) {
        $parameters.message_thread_id = $Telegram.ThreadId
    }
    $sent = Invoke-TelegramRequest -Token $Telegram.Token -Method sendMessage -Parameters $parameters
    if ($null -eq $sent.message_id) {
        throw 'Telegram Bot API не подтвердил отправку сообщения.'
    }
}

function Connect-TelegramGroup {
    param([string]$Name = 'WB', [string]$OtherBotId)

    $token = Read-Host "Вставьте токен Telegram для $Name из BotFather" -AsSecureString
    $bot = Invoke-TelegramRequest -Token $token -Method getMe
    if ($bot.username -notmatch '^[A-Za-z0-9_]+$') {
        throw 'Telegram не вернул имя пользователя бота.'
    }

    if ($OtherBotId -and [string]$bot.id -eq $OtherBotId) {
        throw 'Для WB и Ozon нужны разные боты. Введите токен отдельного бота из BotFather.'
    }
    $command = '/' + $Name.ToLowerInvariant() + '_' + [guid]::NewGuid().ToString('N').Substring(0, 12) + '@' + $bot.username
    Write-Host "Добавьте @$($bot.username) в нужную группу и отправьте в ней команду ниже."
    Write-Host $command
    Write-Host 'Ожидаю команду до трёх минут.'

    $deadline = [DateTime]::UtcNow.AddMinutes(3)
    $parameters = @{ timeout = 10; allowed_updates = @('message') }
    while ([DateTime]::UtcNow -lt $deadline) {
        $updates = Invoke-TelegramRequest -Token $token -Method getUpdates -Parameters $parameters
        foreach ($update in $updates) {
            $parameters.offset = [long]$update.update_id + 1
            $message = $update.message
            if ($message.chat.type -in 'group', 'supergroup' -and $message.text -ceq $command) {
                $destination = [pscustomobject]@{
                    ChatId = [long]$message.chat.id
                    ThreadId = if ($message.is_topic_message) { [long]$message.message_thread_id } else { 0 }
                    Title = $message.chat.title
                    Token = $token
                }
                return $destination
            }
        }
    }
    throw 'Команда из группы не получена. Повторите SetupTelegram и отправьте новую команду в нужную группу.'
}

function Read-TelegramSettings {
    try {
        $settings = ConvertFrom-Json -InputObject ([IO.File]::ReadAllText($telegramPath)) -ErrorAction Stop
        ConvertTo-TelegramDestination $settings
    }
    catch {
        throw 'Не удалось прочитать настройки Telegram. Запустите SetupTelegram под текущим пользователем Windows.'
    }
}

function ConvertTo-TelegramDestination {
    param($Settings)

    if ([long]$Settings.ChatId -ge 0 -or [long]$Settings.ThreadId -lt 0) {
        throw 'Некорректный ID группы или темы.'
    }
    [pscustomobject]@{
        ChatId = [long]$Settings.ChatId
        ThreadId = [long]$Settings.ThreadId
        Token = ConvertTo-SecureString -String $Settings.EncryptedToken -ErrorAction Stop
    }
}

function ConvertFrom-TelegramDestination {
    param($Telegram)

    @{
        ChatId = $Telegram.ChatId
        ThreadId = $Telegram.ThreadId
        EncryptedToken = ConvertFrom-SecureString -SecureString $Telegram.Token -ErrorAction Stop
    }
}

function Read-NotifiedIds {
    param([string]$Path)

    $ids = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    if ([IO.File]::Exists($Path)) {
        foreach ($id in [IO.File]::ReadLines($Path)) {
            [void]$ids.Add($id)
        }
    }
    return ,$ids
}

function Write-WatchLog {
    param([string]$Message, [switch]$IsError)

    $line = '[{0:dd.MM.yyyy HH:mm:ss}] {1}' -f (Get-Date), $Message
    if ($Background) {
        [IO.File]::AppendAllText($backgroundLogPath, $line + [Environment]::NewLine)
    }
    elseif ($IsError) {
        [Console]::Error.WriteLine($line)
    }
    else {
        Write-Host $line
    }
}

function Send-WbUpdates {
    param($Telegram, $Token, $Counts, [string]$StatePath)

    if ($Counts.Total -eq 0) {
        return
    }

    $notifiedIds = Read-NotifiedIds -Path $StatePath
    $currentIds = Get-WbUnansweredIds -Token $Token
    $newCount = 0
    foreach ($id in $currentIds) {
        if ($notifiedIds.Add($id)) {
            $newCount++
        }
    }
    if ($newCount -eq 0) {
        return
    }

    $message = "Wildberries — новых отзывов без ответа $newCount.`nВсего ждут ответа $($currentIds.Count).`nWB Partners → Товары и цены → Отзывы и вопросы."
    Send-TelegramMessage -Telegram $Telegram -Message $message
    Write-LocalFile -Path $StatePath -Content ($notifiedIds -join "`n")
    Write-WatchLog 'Уведомление отправлено в Telegram.'
}

function Show-WbFeedbackCounts {
    param($Counts)

    Write-WatchLog ('WB — необработанных отзывов {0}, из них за сегодня {1}.' -f $Counts.Total, $Counts.Today)
}

function Read-WbNewsState {
    param([string]$Path)

    if (-not [IO.File]::Exists($Path)) { return }
    try {
        $state = ConvertFrom-Json -InputObject ([IO.File]::ReadAllText($Path)) -ErrorAction Stop
        $date = [DateTime]::ParseExact($state.FromDate, 'yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture)
        $lastId = [uint64]::Parse([string]$state.LastId, [Globalization.CultureInfo]::InvariantCulture)
        [pscustomobject]@{ FromDate = $date.ToString('yyyy-MM-dd'); LastId = $lastId }
    }
    catch {
        throw 'Не удалось прочитать историю новостей WB. Проверьте файл переноса или локальный файл состояния.'
    }
}

function Save-WbNewsState {
    param([string]$Path, [string]$FromDate, [uint64]$LastId)

    $state = @{ FromDate = $FromDate; LastId = [string]$LastId }
    Write-LocalFile -Path $Path -Content (ConvertTo-Json -InputObject $state -Compress)
}

function Send-WbNews {
    param($Telegram, [securestring]$Token, [string]$StatePath)

    $state = Read-WbNewsState -Path $StatePath
    $lastId = [uint64]0
    $fromDate = [DateTime]::UtcNow.AddDays(-2).ToString('yyyy-MM-dd')
    if ($null -ne $state) {
        $lastId = $state.LastId
        $fromDate = $state.FromDate
    }
    $query = if ($lastId -gt 0) { "fromID=$lastId" } else { "from=$fromDate" }
    $payload = Invoke-WbRequest -Token $Token -Uri "${newsUri}?$query"
    if ($null -eq $payload.PSObject.Properties['data'] -or
        ($null -ne $payload.data -and $payload.data -isnot [array])) {
        throw 'WB API вернул некорректный список новостей.'
    }

    $newsById = @{}
    foreach ($item in $payload.data) {
        $id = [uint64]0
        if (-not [uint64]::TryParse([string]$item.id, [ref]$id) -or $id -eq 0) {
            throw 'WB API вернул новость без корректного ID.'
        }
        if ($id -gt $lastId) {
            $newsById[$id] = (([string]$item.header -replace '\s+', ' ').Trim())
        }
    }
    $newIds = @($newsById.Keys | Sort-Object)
    if ($null -eq $state) {
        if ($newIds.Count -gt 0) { $lastId = $newIds[-1] }
        Save-WbNewsState -Path $StatePath -FromDate $fromDate -LastId $lastId
        Write-WatchLog 'Новости WB подключены. Будут отправляться новые публикации.'
        return
    }
    if ($newIds.Count -eq 0) { return }

    $lines = foreach ($id in ($newIds | Select-Object -First 8)) {
        $title = $newsById[$id]
        if (-not $title) { $title = "Публикация $id" }
        if ($title.Length -gt 280) { $title = $title.Substring(0, 280) + '…' }
        "$title`nhttps://seller.wildberries.ru/news-v2/news-details?id=$id"
        $lastId = $id
    }
    $message = "Новости WB Partners`n`n" + ($lines -join "`n`n")
    Send-TelegramMessage -Telegram $Telegram -Message $message -DisableLinkPreview
    Save-WbNewsState -Path $StatePath -FromDate $fromDate -LastId $lastId
    Write-WatchLog 'Новости WB отправлены в Telegram.'
}

function Open-WatchLock {
    [void][IO.Directory]::CreateDirectory($tokenDirectory)
    $lockPath = Join-Path $tokenDirectory 'watch.lock'
    try {
        [IO.File]::Open($lockPath, [IO.FileMode]::OpenOrCreate, [IO.FileAccess]::Write, [IO.FileShare]::None)
    }
    catch {
        throw 'Бот уже запущен или настройки заняты. Остановите бота пунктом 11 в setup-bot.cmd или через Ctrl+C в его окне.'
    }
}

if ($Mode -in 'ExportTransfer', 'ImportTransfer', 'EnableStartup', 'DisableStartup', 'Menu',
    'StartBackground', 'StopBackground', 'ShowBackgroundLog') {
    try {
        . (Join-Path $PSScriptRoot 'wb-bot-setup.ps1')
        if ($Mode -eq 'Menu') {
            Show-WbSetupMenu
        }
        elseif ($Mode -in 'StartBackground', 'StopBackground', 'ShowBackgroundLog') {
            Invoke-WbBackgroundControl -Action $Mode
        }
        elseif ($Mode -in 'ExportTransfer', 'ImportTransfer') {
            $transferLock = Open-WatchLock
            try {
                if ($Mode -eq 'ExportTransfer') { Export-WbTransfer }
                else { Import-WbTransfer }
            }
            finally {
                $transferLock.Dispose()
            }
        }
        else {
            Set-WbStartup -Enabled ($Mode -eq 'EnableStartup')
        }
    }
    catch {
        [Console]::Error.WriteLine($_.Exception.Message)
        exit 1
    }
    exit 0
}

if ($Mode -in 'SetupOzon', 'CheckOzon') {
    try {
        . (Join-Path $PSScriptRoot 'ozon-bot.ps1')
        if ($Mode -eq 'SetupOzon') {
            $setupLock = Open-WatchLock
            try { Set-OzonBot }
            finally { $setupLock.Dispose() }
        }
        else {
            foreach ($account in (Read-OzonSettings).Accounts) {
                Test-OzonCredentials $account
                Test-OzonAccount $account
            }
        }
    }
    catch {
        [Console]::Error.WriteLine($_.Exception.Message)
        exit 1
    }
    exit 0
}

if ($Mode -eq 'Setup') {
    try {
        $token = Read-Host 'Вставьте токен WB API и нажмите Enter' -AsSecureString
        $counts = Get-WbFeedbackCounts -Token $token
        $encrypted = ConvertFrom-SecureString -SecureString $token -ErrorAction Stop
        Write-LocalFile -Path $tokenPath -Content $encrypted
        Write-Host 'Доступ проверен. Токен сохранён для текущего пользователя Windows.'
        Show-WbFeedbackCounts -Counts $counts
    }
    catch {
        [Console]::Error.WriteLine($_.Exception.Message)
        exit 1
    }
    exit 0
}

if ($Mode -in 'SetupTelegram', 'TestTelegram') {
    try {
        if ($Mode -eq 'TestTelegram') {
            if (-not [IO.File]::Exists($telegramPath)) {
                throw 'Сначала запустите скрипт с параметром -Mode SetupTelegram.'
            }
            $telegram = Read-TelegramSettings
            Send-TelegramMessage -Telegram $telegram -Message 'WB — тестовое уведомление.'
            Write-Host 'Тестовое сообщение отправлено в настроенную группу.'
        }
        else {
            $otherBotId = ''
            if ([IO.File]::Exists($ozonPath)) {
                . (Join-Path $PSScriptRoot 'ozon-bot.ps1')
                $ozonTelegram = (Read-OzonSettings).Telegram
                $otherBotId = ([pscredential]::new('Ozon', $ozonTelegram.Token).GetNetworkCredential().Password.Trim() -split ':')[0]
            }
            $telegram = Connect-TelegramGroup -OtherBotId $otherBotId
            $settings = ConvertFrom-TelegramDestination $telegram
            Send-TelegramMessage -Telegram $telegram -Message 'WB — эта группа подключена к уведомлениям об отзывах.'
            Write-LocalFile -Path $telegramPath -Content (ConvertTo-Json -InputObject $settings -Compress)
            Write-Host "Группа «$($telegram.Title)» подключена. Токен сохранён для текущего пользователя Windows."
        }
    }
    catch {
        [Console]::Error.WriteLine($_.Exception.Message)
        exit 1
    }
    exit 0
}

if ($Mode -eq 'Check') {
    try {
        if (-not [IO.File]::Exists($tokenPath)) { throw 'Сначала настройте токен WB через setup-bot.cmd.' }
        $token = ConvertTo-SecureString -String ([IO.File]::ReadAllText($tokenPath)) -ErrorAction Stop
        Show-WbFeedbackCounts -Counts (Get-WbFeedbackCounts -Token $token)
    }
    catch {
        [Console]::Error.WriteLine($_.Exception.Message)
        exit 1
    }
    exit 0
}

$watchLock = $null
try {
    $watchLock = Open-WatchLock
    $sendNews = $false
    $ozonOrderReminderMinutes = 60
    $wbEnabled = [IO.File]::Exists($tokenPath)
    $ozon = $null
    if ([IO.File]::Exists($ozonPath)) {
        try {
            . (Join-Path $PSScriptRoot 'ozon-bot.ps1')
            $ozon = Read-OzonSettings
        }
        catch { Write-WatchLog $_.Exception.Message -IsError }
    }
    if ([IO.File]::Exists($configPath)) {
        try {
            $config = ConvertFrom-Json -InputObject ([IO.File]::ReadAllText($configPath)) -ErrorAction Stop
            if ($config -isnot [pscustomobject] -or
                ($config.IntervalMinutes -isnot [int] -and $config.IntervalMinutes -isnot [long]) -or
                $config.IntervalMinutes -lt 1 -or $config.IntervalMinutes -gt 1440 -or
                $config.SendEveryCheck -isnot [bool] -or
                ($config.PSObject.Properties['SendNews'] -and $config.SendNews -isnot [bool])) {
                throw 'Invalid configuration'
            }
        }
        catch {
            throw 'Проверьте config.json. IntervalMinutes должен быть целым числом от 1 до 1440, а SendEveryCheck и SendNews — false или true.'
        }
        if (-not $PSBoundParameters.ContainsKey('IntervalMinutes')) { $IntervalMinutes = $config.IntervalMinutes }
        if (-not $PSBoundParameters.ContainsKey('SendEveryCheck')) { $SendEveryCheck = $config.SendEveryCheck }
        $sendNews = $config.SendNews -eq $true
        if ($config.PSObject.Properties['OzonOrderReminderMinutes']) {
            $minutes = $config.OzonOrderReminderMinutes
            if (($minutes -isnot [int] -and $minutes -isnot [long]) -or $minutes -lt 0 -or $minutes -gt 10080) {
                throw 'Проверьте OzonOrderReminderMinutes в config.json. Укажите целое число от 0 до 10080.'
            }
            $ozonOrderReminderMinutes = $minutes
        }
    }

    $telegram = $null
    if ($wbEnabled) {
        try {
            $token = ConvertTo-SecureString -String ([IO.File]::ReadAllText($tokenPath)) -ErrorAction Stop
            if ([IO.File]::Exists($telegramPath)) {
                $telegram = Read-TelegramSettings
                $statePath = Join-Path $tokenDirectory "notified-$($telegram.ChatId)-$($telegram.ThreadId).txt"
                $newsStatePath = Join-Path $tokenDirectory "news-$($telegram.ChatId)-$($telegram.ThreadId).json"
            }
            elseif ($SendEveryCheck) {
                throw 'Для отправки сводок WB в Telegram сначала выполните SetupTelegram.'
            }
            else {
                Write-WatchLog 'Для уведомлений WB выполните SetupTelegram.'
            }
        }
        catch {
            Write-WatchLog "WB. $($_.Exception.Message)" -IsError
            $wbEnabled = $false
        }
    }
    if (-not $wbEnabled -and $null -eq $ozon) { throw 'Сначала настройте WB или Ozon через setup-bot.cmd.' }

    $services = if ($wbEnabled -and $ozon) { 'WB и Ozon' } elseif ($wbEnabled) { 'WB' } else { 'Ozon' }
    Write-WatchLog "Проверка $services каждые $IntervalMinutes мин."
    if (-not $Background) { Write-Host 'Для остановки нажмите Ctrl+C.' }

    $wbRetryAt = [datetime]::MinValue
    while (-not $Background -or -not [IO.File]::Exists($backgroundStopPath)) {
        $intervalSeconds = $IntervalMinutes * 60
        $wbWaitSeconds = $intervalSeconds
        if ($null -ne $ozon) { Watch-Ozon -Ozon $ozon -OrderReminderMinutes $ozonOrderReminderMinutes }
        if ($wbEnabled -and [datetime]::UtcNow -ge $wbRetryAt) {
            if ($sendNews -and $null -ne $telegram) {
                try {
                    Send-WbNews -Telegram $telegram -Token $token -StatePath $newsStatePath
                }
                catch {
                    Write-WatchLog "Новости WB. $($_.Exception.Message)" -IsError
                }
            }
            try {
                $counts = Get-WbFeedbackCounts -Token $token
                Show-WbFeedbackCounts -Counts $counts
                if ($SendEveryCheck) {
                    $message = "Wildberries — сводка по отзывам.`nВсего ждут ответа $($counts.Total).`nЗа сегодня $($counts.Today)."
                    Send-TelegramMessage -Telegram $telegram -Message $message
                    Write-WatchLog 'Сводка отправлена в Telegram.'
                }
                elseif ($null -ne $telegram) {
                    Send-WbUpdates -Telegram $telegram -Token $token -Counts $counts -StatePath $statePath
                }
            }
            catch {
                Write-WatchLog $_.Exception.Message -IsError
                if ($_.Exception.Data['Status'] -in 401, 402, 403) {
                    if ($null -eq $ozon) { exit 1 }
                    $wbEnabled = $false
                }
                $retrySeconds = $_.Exception.Data['RetrySeconds']
                if ($retrySeconds -gt $wbWaitSeconds) {
                    $wbWaitSeconds = $retrySeconds
                    $wbRetryAt = [datetime]::UtcNow.AddSeconds($retrySeconds)
                }
            }
        }
        $waitSeconds = if ($null -eq $ozon) { $wbWaitSeconds } else { $intervalSeconds }
        if ($Background) {
            $resumeAt = [DateTime]::UtcNow.AddSeconds($waitSeconds)
            while ([DateTime]::UtcNow -lt $resumeAt -and -not [IO.File]::Exists($backgroundStopPath)) {
                Start-Sleep -Seconds 1
            }
        }
        else {
            Start-Sleep -Seconds $waitSeconds
        }
    }
    Write-WatchLog 'Фоновый бот остановлен.'
}
catch {
    Write-WatchLog $_.Exception.Message -IsError
    exit 1
}
finally {
    if ($null -ne $watchLock) {
        $watchLock.Dispose()
    }
}
