$mailserver = "yourexchangeserver.example.com" # Укажите имя своего Exchange-сервера
$topnum = "10" # Определите количество строк в результирующих таблицах
[bool]$report = $false # Формировать суммарный отчет по всем ящикам в csv-файл. $true - да, $false - нет
function Get-ExchangeStat {
$exchdata = Get-WmiObject -Namespace root\MicrosoftExchangeV2 -class Exchange_Mailbox -ComputerName $mailserver | Select-Object -Property LastLoggedOnUserAccount,MailboxDisplayName,Size,TotalItems,StorageGroupName,StoreName,LastLogonTime,LastLogoffTime | ? -FilterScript { ( $_.MailboxDisplayName -notmatch "SMTP") -and ( $_.MailboxDisplayName -notmatch "SystemMailbox" ) -and ($_.LastLoggedOnUserAccount.Count -gt 0) }
# суммарная статистика
foreach ($Items in $exchdata.totalitems) {$totalItems += $items} 
foreach ($Size in $exchdata.Size) {$totalSize += $Size}
$totalSizeCV = $totalSize /1048576
# статистика по определенным критериям
$exchdatasize = $exchdata | Sort-Object Size -Descending | Select-Object -First $topnum
$exchdataitems = $exchdata | Sort-Object TotalItems -Descending | Select-Object -First $topnum
$exchdatalogon = $exchdata | Sort-Object LastLogonTime | Select-Object -First $topnum
# определяем формат времени
$exchdata | % {
    $fields = @("LastLogonTime","LastLogoffTime")
    foreach ($field in $fields) {
            if ($_.$field) {
            $_.$field = $_.$field -replace "\.\d{6}\+\*{3}",""
            $_.$field = [datetime]::ParseExact($_.$field,"yyyyMMddHHmmss",$null)
            $_.$field = '{0:dd.MM.yyyy HH:mm:ss}' -f $_.$field
            }
    }
}
# выводим глобальную статистику в csv-файл
if ($report) { $exchdata | Export-Csv -Path $PSScriptRoot\mailboxstat.csv -Delimiter ";" -Encoding Default -NoTypeInformation } # Экспорт статистики в CSV
# конвертируем размеры из КБ в ГБ
$exchdata | % {
$_.Size = $_.Size /1048576
$_.Size = "{0:N1}" -f $_.Size
}
# выводим статистику в консоль
Write-Host;
Write-Host "TOP"$topnum": Самых больших почтовых ящиков" -ForegroundColor Gray
$exchdatasize | ft
Write-Host "TOP"$topnum": Больше всего писем" -ForegroundColor Gray
$exchdataitems | ft
Write-Host "TOP"$topnum": Неактивных ящиков" -ForegroundColor Gray
$exchdatalogon | ft
Write-Host "ИТОГО: суммарные данные" -ForegroundColor Gray
Write-Host;
"Количество писем     : " + "{0:N0}" -f $totalItems
"Суммарный размер (GB): " + "{0:N1}" -f $totalSizeCV
if ($report) {
Write-Host;
Write-Host "Сформирован csv-отчет:" $PSScriptRoot\mailboxstat.csv -ForegroundColor Gray
}
}
Get-ExchangeStat
