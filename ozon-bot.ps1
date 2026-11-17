function Invoke-OzonRequest {
    param($Account, [string]$Path, [hashtable]$Body)

    try {
        $apiKey = [pscredential]::new('Ozon', $Account.ApiKey).GetNetworkCredential().Password.Trim()
        $response = Invoke-WebRequest -Uri "https://api-seller.ozon.ru$Path" -Method Post -UseBasicParsing `
            -Headers @{ 'Client-Id' = $Account.ClientId; 'Api-Key' = $apiKey } `
            -ContentType 'application/json; charset=utf-8' -Body (ConvertTo-Json $Body -Depth 5 -Compress) `
            -TimeoutSec 20 -MaximumRedirection 0 -Verbose:$false -Debug:$false -ErrorAction Stop
        # Windows PowerShell can use a legacy encoding when the server omits charset.
        $json = [Text.Encoding]::UTF8.GetString($response.RawContentStream.ToArray()).TrimStart([char]0xFEFF)
        ConvertFrom-Json -InputObject $json -ErrorAction Stop
    }
    catch {
        $response = $_.Exception.Response
        $status = if ($null -eq $response) { 0 } else { [int]$response.StatusCode }
        $serverMessage = ''
        if ($status -in 401, 403 -and $_.ErrorDetails.Message) {
            try {
                $serverMessage = [string](ConvertFrom-Json -InputObject $_.ErrorDetails.Message -ErrorAction Stop).message
            }
            catch {
                # Keep the status-based guidance when the response is not JSON.
            }
        }
        $details = switch ($status) {
            0 { 'сбой соединения или некорректный ответ API' }
            401 { 'HTTP 401, Ozon отклонил авторизацию. Проверьте пару Client ID и ключа в одном кабинете Seller API, срок действия ключа и разрешённые сети' }
            403 { 'HTTP 403, проверьте права API-ключа и доступ к методу по подписке' }
            429 { 'HTTP 429, превышен лимит запросов' }
            default { "HTTP $status" }
        }
        # Translate known reasons without echoing arbitrary server text or credentials.
        $hint = switch -Regex ($serverMessage) {
            '(?:expired|ист[её]к)' { 'Ozon сообщает, что срок действия ключа истёк.'; break }
            '(?:\bip\b|allowlist|whitelist|разреш[её]нн\S* сет)' { 'Проверьте, входит ли внешний IP компьютера в разрешённые сети ключа.'; break }
            '(?:invalid|incorrect|неверн|недействит).*(?:client|key|ключ)' { 'Ozon сообщает, что Client ID или ключ недействителен.'; break }
        }
        $failure = [InvalidOperationException]::new(("Ozon $Path — $details. $hint").TrimEnd())
        $failure.Data['OzonStatus'] = $status
        throw $failure
    }
}

function Get-OzonItems {
    param($Account, [string]$Path, [hashtable]$Body, [string]$List)

    $seenCursors = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    do {
        $page = Invoke-OzonRequest -Account $Account -Path $Path -Body $Body
        if ($null -eq $page.PSObject.Properties[$List] -or
            ($null -ne $page.$List -and $page.$List -isnot [array])) {
            throw "Ozon $Path вернул некорректный список."
        }
        $page.$List
        $hasNext = [string]$page.has_next -eq 'true'
        if ($hasNext) {
            $next = [string]$page.cursor
            if (-not $next -or -not $seenCursors.Add($next)) {
                throw "Ozon $Path вернул некорректный курсор."
            }
            $Body.cursor = $next
        }
    } while ($hasNext)
}

function Get-OzonPostings {
    param($Account)

    $today = [datetime]::UtcNow
    $body = @{
        limit = 100
        sort_dir = 'ASC'
        cursor = ''
        filter = @{
            cutoff_from = $today.AddDays(-330).ToString('yyyy-MM-ddTHH:mm:ssZ')
            cutoff_to = $today.AddDays(30).ToString('yyyy-MM-ddTHH:mm:ssZ')
            statuses = @('awaiting_packaging')
            warehouse_ids = @($Account.Warehouses | ForEach-Object { $_.Id })
        }
    }
    Get-OzonItems $Account '/v4/posting/fbs/unfulfilled/list' $body 'postings' |
        Where-Object { $_.status -eq 'awaiting_packaging' -and $_.delivery_schema -in $null, '', 'fbs' }
}

function Get-OzonUnreadChats {
    param($Account)

    Get-OzonItems $Account '/v3/chat/list' @{ limit = 100; cursor = ''; filter = @{ unread_only = $true } } 'chats' |
        Where-Object { $_.chat.chat_type -in 'Buyer_Seller', 'BUYER_SELLER' -and $_.unread_count -gt 0 }
}

function Get-OzonUnreadMessages {
    param($Account, $Chat)

    $firstUnread = [decimal]0
    if (-not [decimal]::TryParse([string]$Chat.first_unread_message_id, [ref]$firstUnread) -or $firstUnread -le 0) {
        throw 'Ozon не вернул ID первого непрочитанного сообщения.'
    }
    $body = @{ chat_id = $Chat.chat.chat_id; direction = 'Backward'; limit = 100 }
    do {
        $page = Invoke-OzonRequest $Account '/v3/chat/history' $body
        if ($null -eq $page.PSObject.Properties['messages']) {
            throw 'Ozon не вернул историю чата.'
        }
        $oldest = [decimal]::MaxValue
        foreach ($message in $page.messages) {
            $id = [decimal]0
            if (-not [decimal]::TryParse([string]$message.message_id, [ref]$id) -or $id -le 0) {
                throw 'Ozon вернул некорректный ID сообщения.'
            }
            $oldest = [Math]::Min($oldest, $id)
            if ($id -ge $firstUnread -and [string]$message.is_read -eq 'false' -and
                $message.user.type -in 'Customer', 'Сustomer', 'Buyer') {
                $message
            }
        }
        if ($oldest -le $firstUnread -or [string]$page.has_next -ne 'true') { return }
        if ($oldest -eq [decimal]::MaxValue -or ($body.from_message_id -and $oldest -ge $body.from_message_id)) {
            throw 'Ozon вернул повторную страницу истории чата.'
        }
        $body.from_message_id = $oldest
    } while ($true)
}

function Save-OzonEventHistory {
    param([string]$Path, $History)

    $lines = foreach ($entry in $History.GetEnumerator()) {
        if ($entry.Value -gt 0) { "$($entry.Key)`t$($entry.Value)" }
        else { $entry.Key }
    }
    Write-LocalFile -Path $Path -Content ($lines -join "`n")
}

function Send-OzonEvents {
    param($Telegram, [string]$StatePath, [string]$Heading, [object[]]$Events, [int]$ReminderMinutes = 0)

    if ($Events.Count -eq 0) { return }
    $notified = [Collections.Generic.Dictionary[string,long]]::new([StringComparer]::Ordinal)
    foreach ($line in (Read-NotifiedIds -Path $StatePath)) {
        $parts = $line.Split([char]9, 2)
        $id = $parts[0]
        $sentAt = 0L
        if ($parts.Count -eq 2) { [void][long]::TryParse($parts[1], [ref]$sentAt) }
        if (-not $notified.ContainsKey($id) -or $sentAt -gt $notified[$id]) { $notified[$id] = $sentAt }
    }

    $now = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds()
    $repeatFrom = $now - ($ReminderMinutes * 60)
    $initialized = $false
    $pendingIds = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    $pending = @(
        foreach ($event in $Events) {
            if (-not $pendingIds.Add($event.Id)) { continue }
            if (-not $notified.ContainsKey($event.Id)) {
                $event
            }
            elseif ($ReminderMinutes -gt 0) {
                if ($notified[$event.Id] -eq 0) {
                    # Older history recorded IDs without a send time.
                    $notified[$event.Id] = $now
                    $initialized = $true
                }
                elseif ($notified[$event.Id] -le $repeatFrom) {
                    @{ Id = $event.Id; Line = "$($event.Line) — всё ещё ждёт сборки" }
                }
            }
        }
    )
    if ($initialized) { Save-OzonEventHistory -Path $StatePath -History $notified }
    if ($pending.Count -eq 0) { return }
    for ($offset = 0; $offset -lt $pending.Count; $offset += 15) {
        $end = [Math]::Min($offset + 14, $pending.Count - 1)
        $batch = $pending[$offset..$end]
        $message = "$Heading — $($batch.Count)`n" + (($batch | ForEach-Object { $_.Line }) -join "`n")
        Send-TelegramMessage -Telegram $Telegram -Message $message
        # Save only the batches Telegram has confirmed.
        $sentAt = if ($ReminderMinutes -gt 0) { [DateTimeOffset]::UtcNow.ToUnixTimeSeconds() } else { 0L }
        foreach ($event in $batch) { $notified[$event.Id] = $sentAt }
        Save-OzonEventHistory -Path $StatePath -History $notified
    }
    Write-Host "$Heading — отправлено $($pending.Count)."
}

function Watch-Ozon {
    param($Ozon, [int]$OrderReminderMinutes = 60)

    $telegram = $Ozon.Telegram
    $historyPrefix = Join-Path $tokenDirectory "ozon-$($telegram.ChatId)-$($telegram.ThreadId)"
    foreach ($account in $Ozon.Accounts) {
        $label = ($account.Warehouses.Country -join ' / ')
        try {
            $postings = @(Get-OzonPostings $account)
            foreach ($warehouse in $account.Warehouses) {
                $events = @(
                    foreach ($posting in $postings) {
                        if ([string]$posting.delivery_method.warehouse_id -eq $warehouse.Id -and $posting.posting_number) {
                            @{ Id = "$($account.ClientId)/$($warehouse.Id)/$($posting.posting_number)"; Line = "• $($posting.posting_number)" }
                        }
                    }
                )
                Send-OzonEvents -Telegram $telegram -StatePath "$historyPrefix-orders.txt" `
                    -Heading "Ozon FBS · $($warehouse.Country) · $($warehouse.Name) — к сборке" `
                    -Events $events -ReminderMinutes $OrderReminderMinutes
            }
            if (@($postings | Where-Object { [string]$_.delivery_method.warehouse_id -notin $account.Warehouses.Id }).Count -gt 0) {
                Write-Host "Ozon FBS ($label). В ответе есть отправления вне выбранных складов."
            }
        }
        catch { [Console]::Error.WriteLine("Ozon FBS ($label). $($_.Exception.Message)") }

        try {
            $events = @(
                foreach ($chat in (Get-OzonUnreadChats $account)) {
                    if (-not $chat.chat.chat_id) { continue }
                    try {
                        $messages = @(Get-OzonUnreadMessages $account $chat)
                        foreach ($message in $messages) {
                            $line = "• Чат $($chat.chat.chat_id)"
                            if ($message.context.order_number) { $line += " · заказ $($message.context.order_number)" }
                            @{ Id = "$($account.ClientId)/$($chat.chat.chat_id)/$($message.message_id)"; Line = $line }
                        }
                    }
                    catch { [Console]::Error.WriteLine("Ozon Чат ($label). $($_.Exception.Message)") }
                }
            )
            Send-OzonEvents $telegram "$historyPrefix-chats.txt" "Ozon · $label · непрочитанные сообщения покупателей" $events
        }
        catch { [Console]::Error.WriteLine("Ozon Чаты ($label). $($_.Exception.Message)") }
    }
}

function Repair-OzonWarehouseName {
    param([string]$Name)

    if ($Name -notmatch '[ÐÑ][^\x00-\x7F]') { return $Name }

    # Earlier versions could save a UTF-8 name decoded as a Western code page.
    $utf8 = [Text.UTF8Encoding]::new($false, $true)
    foreach ($codePage in 28591, 1252) {
        $encoding = [Text.Encoding]::GetEncoding($codePage, [Text.EncoderFallback]::ExceptionFallback,
            [Text.DecoderFallback]::ExceptionFallback)
        try {
            $corrected = $utf8.GetString($encoding.GetBytes($Name))
            if ($corrected -match '[А-Яа-яЁё]') { return $corrected }
        }
        catch [Text.EncoderFallbackException] {}
        catch [Text.DecoderFallbackException] {}
    }
    $Name
}

function Read-OzonSettings {
    try {
        $settings = ConvertFrom-Json -InputObject ([IO.File]::ReadAllText($ozonPath)) -ErrorAction Stop
        $telegram = ConvertTo-TelegramDestination -Settings $settings.Telegram
        $accounts = @(
            foreach ($saved in $settings.Accounts) {
                if ([string]$saved.ClientId -notmatch '^\d+$' -or @($saved.Warehouses).Count -lt 1) {
                    throw 'Некорректные данные кабинета'
                }
                foreach ($warehouse in $saved.Warehouses) {
                    $warehouse.Name = Repair-OzonWarehouseName $warehouse.Name
                }
                [pscustomobject]@{
                    ClientId = [string]$saved.ClientId
                    ApiKey = ConvertTo-SecureString -String $saved.EncryptedApiKey -ErrorAction Stop
                    Warehouses = $saved.Warehouses
                }
            }
        )
        if ($accounts.Count -eq 0) { throw 'Нет кабинетов' }
        [pscustomobject]@{ Telegram = $telegram; Accounts = $accounts }
    }
    catch {
        throw 'Не удалось прочитать настройки Ozon. Откройте setup-bot.cmd и повторите настройку Ozon или импорт.'
    }
}

function Read-OzonAccount {
    param([string]$Country)

    $clientId = (Read-Host "Client ID кабинета Ozon ($Country)").Trim()
    if ($clientId -notmatch '^\d+$') { throw 'Client ID должен состоять из цифр.' }
    $apiKey = Read-Host "API-ключ Ozon ($Country)" -AsSecureString
    $keyText = [pscredential]::new('Ozon', $apiKey).GetNetworkCredential().Password.Trim()
    if (-not $keyText -or $keyText -match '\s') { throw 'Вставьте только значение ключа Seller API одной строкой.' }
    if ($keyText -match '^\d+:[A-Za-z0-9_-]+$') {
        throw 'Похоже, введён токен Telegram. На этом шаге нужен API-ключ из Ozon Seller.'
    }
    [pscustomobject]@{ ClientId = $clientId; ApiKey = $apiKey; Warehouses = @() }
}

function Select-OzonWarehouse {
    param([object[]]$Warehouses, [string]$Country)

    if ($Warehouses.Count -eq 0) { throw "Ozon не вернул склады для выбора ($Country)." }
    Write-Host "Выберите склад ($Country)."
    for ($i = 0; $i -lt $Warehouses.Count; $i++) {
        $item = $Warehouses[$i]
        Write-Host "$($i + 1). $($item.name) · ID $($item.warehouse_id) · $($item.address_info.address)"
    }
    $index = 0
    if (-not [int]::TryParse((Read-Host 'Номер склада'), [ref]$index) -or $index -lt 1 -or $index -gt $Warehouses.Count) {
        throw 'Выберите номер склада из списка.'
    }
    $selected = $Warehouses[$index - 1]
    [pscustomobject]@{ Id = [string]$selected.warehouse_id; Name = $selected.name; Country = $Country }
}

function Get-OzonWarehouses {
    param($Account)
    Get-OzonItems $Account '/v2/warehouse/list' @{ limit = 100; cursor = '' } 'warehouses' |
        Where-Object { $_.warehouse_id -and $_.name }
}

function Test-OzonCredentials {
    param($Account)

    $response = Invoke-OzonRequest -Account $Account -Path '/v1/roles' -Body @{}
    if ($null -eq $response.PSObject.Properties['roles']) { throw 'Ozon не вернул список ролей API-ключа.' }

    $roles = @($response.roles | ForEach-Object { $_.name } | Where-Object { $_ })
    Write-Host 'Ozon принял Client ID и API-ключ.'
    if ($roles.Count -gt 0) { Write-Host "Роли ключа — $($roles -join ', ')." }
    if ($response.expires_at) { Write-Host "Срок действия ключа по данным Ozon — $($response.expires_at)." }
}

function Get-OzonSetupWarehouses {
    param($Account)

    Test-OzonCredentials $Account
    try {
        Get-OzonWarehouses $Account
    }
    catch {
        if ($_.Exception.Data['OzonStatus'] -in 401, 403) {
            throw 'Ozon принял ключ на /v1/roles, но отклонил запрос списка складов. Проверьте доступ к /v2/warehouse/list в настройках роли ключа.'
        }
        throw
    }
}

function Test-OzonAccount {
    param($Account)

    $label = ($Account.Warehouses.Country -join ' / ')
    $postings = @(Get-OzonPostings $Account)
    Write-Host "Ozon FBS ($label). Доступ подтверждён. Найдено отправлений к сборке $($postings.Count)."
    try {
        $chats = @(Get-OzonUnreadChats $Account)
        Write-Host "Ozon Чаты ($label). Доступ подтверждён. Чатов с непрочитанными сообщениями $($chats.Count)."
        if ($chats.Count -gt 0) { $null = @(Get-OzonUnreadMessages $Account $chats[0]) }
    }
    catch { [Console]::Error.WriteLine("Ozon Чаты ($label). $($_.Exception.Message)") }
}

function Set-OzonBot {
    $kz = Read-OzonAccount 'Казахстан'
    $warehouses = @(Get-OzonSetupWarehouses $kz)
    $first = Select-OzonWarehouse $warehouses 'Казахстан'
    $kz.Warehouses = @($first)
    $accounts = @($kz)
    $shared = (Read-Host 'Склад России в этом же кабинете? 1 — да, 2 — другой кабинет').Trim()
    if ($shared -eq '1') {
        $remaining = @($warehouses | Where-Object { [string]$_.warehouse_id -ne $first.Id })
        $kz.Warehouses += Select-OzonWarehouse $remaining 'Россия'
    }
    elseif ($shared -eq '2') {
        $ru = Read-OzonAccount 'Россия'
        if ($ru.ClientId -eq $kz.ClientId) { throw 'Для одного Client ID выберите вариант 1.' }
        $ru.Warehouses = @(Select-OzonWarehouse @(Get-OzonSetupWarehouses $ru) 'Россия')
        $accounts += $ru
    }
    else { throw 'Выберите 1 или 2.' }
    foreach ($account in $accounts) { Test-OzonAccount $account }

    $otherBotId = ''
    if ([IO.File]::Exists($telegramPath)) {
        $wbTelegram = Read-TelegramSettings
        $otherBotId = ([pscredential]::new('WB', $wbTelegram.Token).GetNetworkCredential().Password.Trim() -split ':')[0]
    }
    $telegram = Connect-TelegramGroup -Name 'Ozon' -OtherBotId $otherBotId
    $configuration = @{
        Telegram = ConvertFrom-TelegramDestination $telegram
        Accounts = @(
            foreach ($account in $accounts) {
                @{ ClientId = $account.ClientId; EncryptedApiKey = ConvertFrom-SecureString $account.ApiKey -ErrorAction Stop
                    Warehouses = $account.Warehouses }
            }
        )
    }
    Send-TelegramMessage $telegram 'Ozon — эта группа подключена к уведомлениям.'
    Write-LocalFile $ozonPath (ConvertTo-Json $configuration -Depth 6 -Compress)
    Write-Host 'Бот Ozon настроен. Запустите start-bot.cmd.'
}
