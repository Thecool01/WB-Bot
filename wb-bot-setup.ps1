function Get-TransferKey {
    param([securestring]$Password, [byte[]]$Salt)

    $pointer = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($Password)
    try {
        $text = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($pointer)
        $derivation = [Security.Cryptography.Rfc2898DeriveBytes]::new(
            $text, $Salt, 600000, [Security.Cryptography.HashAlgorithmName]::SHA256)
        try {
            return ,$derivation.GetBytes(64)
        }
        finally {
            $derivation.Dispose()
        }
    }
    finally {
        [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($pointer)
    }
}

function Protect-Transfer {
    param([string]$Content, [securestring]$Password)

    $salt = [byte[]]::new(16)
    $random = [Security.Cryptography.RandomNumberGenerator]::Create()
    $aes = [Security.Cryptography.Aes]::Create()
    $hmac = [Security.Cryptography.HMACSHA256]::new()
    try {
        $random.GetBytes($salt)
        $key = Get-TransferKey -Password $Password -Salt $salt
        $aes.Key = [byte[]]($key[0..31])
        $hmac.Key = [byte[]]($key[32..63])
        $aes.GenerateIV()
        $bytes = [Text.Encoding]::UTF8.GetBytes($Content)
        $encryptor = $aes.CreateEncryptor()
        try {
            $encrypted = $encryptor.TransformFinalBlock($bytes, 0, $bytes.Length)
        }
        finally {
            $encryptor.Dispose()
            [Array]::Clear($bytes, 0, $bytes.Length)
        }
        $authenticated = [byte[]]($aes.IV + $encrypted)
        @{
            Version = 1
            Salt = [Convert]::ToBase64String($salt)
            Data = [Convert]::ToBase64String($authenticated)
            Tag = [Convert]::ToBase64String($hmac.ComputeHash($authenticated))
        } | ConvertTo-Json -Compress
    }
    finally {
        if ($null -ne $key) { [Array]::Clear($key, 0, $key.Length) }
        $random.Dispose()
        $aes.Dispose()
        $hmac.Dispose()
    }
}

function Unprotect-Transfer {
    param([string]$Content, [securestring]$Password)

    try {
        $bundle = ConvertFrom-Json -InputObject $Content -ErrorAction Stop
        $salt = [Convert]::FromBase64String($bundle.Salt)
        $data = [Convert]::FromBase64String($bundle.Data)
        $tag = [Convert]::FromBase64String($bundle.Tag)
        if ($bundle.Version -ne 1 -or $salt.Length -ne 16 -or
            $data.Length -lt 32 -or $data.Length % 16 -ne 0 -or $tag.Length -ne 32) {
            throw 'Invalid bundle'
        }
        $key = Get-TransferKey -Password $Password -Salt $salt
        $hmac = [Security.Cryptography.HMACSHA256]::new([byte[]]($key[32..63]))
        $expectedTag = $hmac.ComputeHash($data)
        $difference = 0
        for ($i = 0; $i -lt $tag.Length; $i++) {
            $difference = $difference -bor ($tag[$i] -bxor $expectedTag[$i])
        }
        if ($difference -ne 0) { throw 'Invalid authentication tag' }
        $aes = [Security.Cryptography.Aes]::Create()
        $aes.Key = [byte[]]($key[0..31])
        $aes.IV = [byte[]]($data[0..15])
        $decryptor = $aes.CreateDecryptor()
        $bytes = $decryptor.TransformFinalBlock($data, 16, $data.Length - 16)
        ConvertFrom-Json -InputObject ([Text.Encoding]::UTF8.GetString($bytes)) -ErrorAction Stop
    }
    catch {
        throw 'Не удалось открыть файл переноса. Проверьте пароль и целостность wb-transfer.json.'
    }
    finally {
        if ($null -ne $bytes) { [Array]::Clear($bytes, 0, $bytes.Length) }
        if ($null -ne $key) { [Array]::Clear($key, 0, $key.Length) }
        if ($null -ne $decryptor) { $decryptor.Dispose() }
        if ($null -ne $aes) { $aes.Dispose() }
        if ($null -ne $hmac) { $hmac.Dispose() }
    }
}

function Export-WbTransfer {
    if (-not [IO.File]::Exists($tokenPath) -or -not [IO.File]::Exists($telegramPath)) {
        throw 'Для переноса на этом компьютере сначала должны быть настроены WB и Telegram.'
    }
    $wbToken = ConvertTo-SecureString -String ([IO.File]::ReadAllText($tokenPath)) -ErrorAction Stop
    $telegram = Read-TelegramSettings
    $historyPath = Join-Path $tokenDirectory "notified-$($telegram.ChatId)-$($telegram.ThreadId).txt"
    $history = Read-NotifiedIds -Path $historyPath
    $newsPath = Join-Path $tokenDirectory "news-$($telegram.ChatId)-$($telegram.ThreadId).json"
    $news = Read-WbNewsState -Path $newsPath

    $password = Read-Host 'Придумайте пароль переноса (от 12 символов)' -AsSecureString
    if ($password.Length -lt 12) { throw 'Введите пароль длиной от 12 символов.' }
    $confirmation = Read-Host 'Повторите пароль переноса' -AsSecureString
    $passwordText = [pscredential]::new('transfer', $password).GetNetworkCredential().Password
    $confirmationText = [pscredential]::new('transfer', $confirmation).GetNetworkCredential().Password
    if ($passwordText -cne $confirmationText) { throw 'Пароли не совпали. Повторите экспорт.' }

    $bundle = @{
        WbToken = [pscredential]::new('WB', $wbToken).GetNetworkCredential().Password
        TelegramToken = [pscredential]::new('Telegram', $telegram.Token).GetNetworkCredential().Password
        ChatId = $telegram.ChatId
        ThreadId = $telegram.ThreadId
        NotifiedIds = @($history | ForEach-Object { $_ })
        News = $news
    }
    if ([IO.File]::Exists($ozonPath)) {
        . (Join-Path $PSScriptRoot 'ozon-bot.ps1')
        $ozon = Read-OzonSettings
        $ozonTelegram = $ozon.Telegram
        $prefix = Join-Path $tokenDirectory "ozon-$($ozonTelegram.ChatId)-$($ozonTelegram.ThreadId)"
        $history = @{}
        foreach ($kind in 'orders', 'chats') {
            $history[$kind] = @((Read-NotifiedIds "$prefix-$kind.txt") | ForEach-Object { $_ })
        }
        $bundle.Ozon = @{
            TelegramToken = [pscredential]::new('Ozon', $ozonTelegram.Token).GetNetworkCredential().Password
            ChatId = $ozonTelegram.ChatId
            ThreadId = $ozonTelegram.ThreadId
            Accounts = @(
                foreach ($account in $ozon.Accounts) {
                    @{
                        ClientId = $account.ClientId
                        ApiKey = [pscredential]::new('Ozon', $account.ApiKey).GetNetworkCredential().Password
                        Warehouses = $account.Warehouses
                    }
                }
            )
            History = $history
        }
    }
    $content = ConvertTo-Json $bundle -Depth 8 -Compress
    $encrypted = Protect-Transfer -Content $content -Password $password
    Write-LocalFile -Path $transferPath -Content $encrypted
    Write-Host "Файл переноса создан. Скопируйте всю папку на новый компьютер."
    Write-Host $transferPath
}

function Import-WbTransfer {
    if (-not [IO.File]::Exists($transferPath)) {
        throw 'Положите wb-transfer.json в папку рядом с ботом и повторите импорт.'
    }
    $password = Read-Host 'Введите пароль переноса' -AsSecureString
    $data = Unprotect-Transfer -Content ([IO.File]::ReadAllText($transferPath)) -Password $password
    $chatId = [long]$data.ChatId
    $threadId = [long]$data.ThreadId
    if ($data.WbToken -notmatch '^[A-Za-z0-9_-]+\.[A-Za-z0-9_-]+\.[A-Za-z0-9_-]+$' -or
        $data.TelegramToken -notmatch '^[0-9]+:[A-Za-z0-9_-]+$' -or
        $chatId -ge 0 -or $threadId -lt 0 -or $data.NotifiedIds -isnot [array]) {
        throw 'В файле переноса отсутствуют корректные настройки бота.'
    }
    $newsPath = Join-Path $tokenDirectory "news-$chatId-$threadId.json"
    if ($null -ne $data.News) {
        try {
            $importDate = [DateTime]::ParseExact($data.News.FromDate, 'yyyy-MM-dd',
                [Globalization.CultureInfo]::InvariantCulture)
            $importId = [uint64]::Parse([string]$data.News.LastId, [Globalization.CultureInfo]::InvariantCulture)
        }
        catch {
            throw 'В файле переноса некорректная история новостей WB.'
        }
        $localNews = Read-WbNewsState -Path $newsPath
        if ($null -eq $localNews -or $importId -gt $localNews.LastId) {
            Save-WbNewsState -Path $newsPath -FromDate $importDate.ToString('yyyy-MM-dd') -LastId $importId
        }
    }
    $wbToken = ConvertTo-SecureString -String $data.WbToken -AsPlainText -Force -ErrorAction Stop
    $telegramToken = ConvertTo-SecureString -String $data.TelegramToken -AsPlainText -Force -ErrorAction Stop
    $encryptedWb = ConvertFrom-SecureString -SecureString $wbToken -ErrorAction Stop
    $telegram = @{
        ChatId = $chatId
        ThreadId = $threadId
        EncryptedToken = ConvertFrom-SecureString -SecureString $telegramToken -ErrorAction Stop
    } | ConvertTo-Json -Compress

    $historyPath = Join-Path $tokenDirectory "notified-$chatId-$threadId.txt"
    $history = Read-NotifiedIds -Path $historyPath
    foreach ($id in $data.NotifiedIds) { [void]$history.Add($id) }
    Write-LocalFile -Path $historyPath -Content ($history -join "`n")
    Write-LocalFile -Path $telegramPath -Content $telegram
    Write-LocalFile -Path $tokenPath -Content $encryptedWb
    if ($null -ne $data.Ozon) {
        $oz = $data.Ozon
        $ozChat = [long]$oz.ChatId
        $ozThread = [long]$oz.ThreadId
        if ($ozChat -ge 0 -or $ozThread -lt 0 -or $oz.TelegramToken -notmatch '^\d+:[A-Za-z0-9_-]+$' -or
            ($oz.TelegramToken -split ':')[0] -eq ($data.TelegramToken -split ':')[0] -or @($oz.Accounts).Count -eq 0) {
            throw 'Файл переноса содержит некорректные настройки Ozon.'
        }
        $ozToken = ConvertTo-SecureString $oz.TelegramToken -AsPlainText -Force -ErrorAction Stop
        $settings = @{
            Telegram = ConvertFrom-TelegramDestination @{ ChatId = $ozChat; ThreadId = $ozThread; Token = $ozToken }
            Accounts = @(
                foreach ($account in $oz.Accounts) {
                    if ($account.ClientId -notmatch '^\d+$' -or -not $account.ApiKey -or @($account.Warehouses).Count -eq 0) {
                        throw 'Файл переноса содержит некорректный кабинет Ozon.'
                    }
                    $key = ConvertTo-SecureString $account.ApiKey -AsPlainText -Force -ErrorAction Stop
                    @{
                        ClientId = $account.ClientId
                        EncryptedApiKey = ConvertFrom-SecureString $key -ErrorAction Stop
                        Warehouses = $account.Warehouses
                    }
                }
            )
        }
        $prefix = Join-Path $tokenDirectory "ozon-$ozChat-$ozThread"
        foreach ($kind in 'orders', 'chats') {
            $path = "$prefix-$kind.txt"
            $ids = Read-NotifiedIds $path
            foreach ($id in $oz.History.$kind) { [void]$ids.Add($id) }
            Write-LocalFile $path ($ids -join "`n")
        }
        Write-LocalFile $ozonPath (ConvertTo-Json $settings -Depth 6 -Compress)
    }
    Write-Host 'Настройки перенесены. Запустите start-bot.cmd двойным щелчком.'
}

function Invoke-WbBackgroundControl {
    param([string]$Action)

    if ($Action -eq 'ShowBackgroundLog') {
        Write-Host $backgroundLogPath
        if ([IO.File]::Exists($backgroundLogPath)) { Get-Content -LiteralPath $backgroundLogPath -Tail 30 -Encoding UTF8 }
        else { Write-Host 'Журнал пока пуст.' }
        return
    }

    [void][IO.Directory]::CreateDirectory($tokenDirectory)
    $controlPath = Join-Path $tokenDirectory 'background-control.lock'
    $control = [IO.File]::Open($controlPath, [IO.FileMode]::OpenOrCreate, [IO.FileAccess]::Write, [IO.FileShare]::None)
    try {
        $processPath = Join-Path $tokenDirectory 'background-process.json'
        $worker = $null
        if ([IO.File]::Exists($processPath)) {
            $saved = ConvertFrom-Json -InputObject ([IO.File]::ReadAllText($processPath)) -ErrorAction Stop
            $candidate = Get-Process -Id $saved.Id -ErrorAction SilentlyContinue
            if ($candidate -and $candidate.ProcessName -eq 'powershell' -and
                [string]$candidate.StartTime.ToUniversalTime().Ticks -ceq $saved.Started) {
                $worker = $candidate
            }
        }

        if ($Action -eq 'StopBackground') {
            if ($null -eq $worker) {
                Write-Host 'Фоновый бот сейчас не запущен.'
                return
            }
            Write-LocalFile -Path $backgroundStopPath -Content 'stop'
            if ($worker.WaitForExit(5000)) { Write-Host 'Фоновый бот остановлен.' }
            else { Write-Host 'Бот остановится после текущей проверки.' }
            return
        }

        if ($null -ne $worker) {
            Write-Host 'Фоновый бот уже запущен.'
            return
        }
        $probe = Open-WatchLock
        $probe.Dispose()
        [IO.File]::Delete($backgroundStopPath)
        Write-LocalFile -Path $backgroundLogPath -Content ''
        $script = Join-Path $PSScriptRoot 'wb-bot.ps1'
        $arguments = "-NoProfile -NonInteractive -ExecutionPolicy Bypass -File `"$script`" -Mode Watch -Background"
        $worker = Start-Process -FilePath (Join-Path $PSHOME 'powershell.exe') -ArgumentList $arguments `
            -WorkingDirectory $PSScriptRoot -WindowStyle Hidden -PassThru -ErrorAction Stop
        if ($worker.WaitForExit(1200)) { throw 'Фоновый бот завершился при запуске. Детали в журнале, пункт 12.' }
        $state = @{ Id = $worker.Id; Started = [string]$worker.StartTime.ToUniversalTime().Ticks }
        Write-LocalFile -Path $processPath -Content (ConvertTo-Json $state -Compress)
        Write-Host 'Фоновый бот запущен. Окно настройки можно закрыть. Журнал — пункт 12.'
    }
    finally {
        $control.Dispose()
    }
}

function Set-WbStartup {
    param([bool]$Enabled)

    $ErrorActionPreference = 'Stop'
    $startup = [Environment]::GetFolderPath([Environment+SpecialFolder]::Startup, [Environment+SpecialFolderOption]::Create)
    if ([string]::IsNullOrWhiteSpace($startup)) { throw 'Windows не вернула папку автозагрузки текущего пользователя.' }
    $shortcutPath = Join-Path $startup 'WB Feedback Bot.lnk'
    if ($Enabled) {
        $script = Join-Path $PSScriptRoot 'wb-bot.ps1'
        $shell = New-Object -ComObject WScript.Shell -ErrorAction Stop
        $shortcut = $shell.CreateShortcut($shortcutPath)
        $shortcut.TargetPath = Join-Path $PSHOME 'powershell.exe'
        $shortcut.Arguments = "-NoProfile -NonInteractive -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$script`" -Mode StartBackground"
        $shortcut.WorkingDirectory = $PSScriptRoot
        $shortcut.Description = 'WB Feedback Bot'
        $shortcut.Save()
        Write-Host 'Автозапуск включён для текущего пользователя Windows.'
    }
    else {
        [IO.File]::Delete($shortcutPath)
        Write-Host 'Автозапуск отключён для текущего пользователя Windows.'
    }
}

function Show-WbSetupMenu {
    Write-Host '1. Настроить или заменить токен WB'
    Write-Host '2. Подключить Telegram для WB'
    Write-Host '3. Экспортировать настройки для переноса'
    Write-Host '4. Импортировать настройки на этом компьютере'
    Write-Host '5. Включить автозапуск при входе в Windows'
    Write-Host '6. Отключить автозапуск'
    Write-Host '7. Проверить доступ к WB'
    Write-Host '8. Настроить Ozon и его бота Telegram'
    Write-Host '9. Проверить доступ к Ozon'
    Write-Host '10. Запустить бота в фоне'
    Write-Host '11. Остановить фонового бота'
    Write-Host '12. Показать журнал фонового бота'
    Write-Host '0. Выход'
    $selectedMode = switch (Read-Host 'Выберите действие') {
        '1' { 'Setup' }
        '2' { 'SetupTelegram' }
        '3' { 'ExportTransfer' }
        '4' { 'ImportTransfer' }
        '5' { 'EnableStartup' }
        '6' { 'DisableStartup' }
        '7' { 'Check' }
        '8' { 'SetupOzon' }
        '9' { 'CheckOzon' }
        '10' { 'StartBackground' }
        '11' { 'StopBackground' }
        '12' { 'ShowBackgroundLog' }
        '0' { return }
        default { throw 'Выберите номер от 0 до 12.' }
    }
    & (Join-Path $PSScriptRoot 'wb-bot.ps1') -Mode $selectedMode
    exit $LASTEXITCODE
}
