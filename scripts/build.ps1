# =====================================================================
#  An Tam So / Cyber Shield SME - BH PVI x MISA
#  build.ps1 : doc file Excel bao cao raw trong /inbox (+ /archive),
#              gop - loai trung - xuat data/data.js cho dashboard.
#  Khong can cai Python / Excel. Chay duoc tren Windows PowerShell 5.1+.
# =====================================================================
param(
    [string]$Root = '',
    [switch]$NoArchive,
    [switch]$NoClipboard,
    [switch]$NoImage,
    [switch]$Force
)

$ErrorActionPreference = 'Stop'
if (-not $Root) { $here = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }; $Root = Split-Path -Parent $here }
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
$utf8 = New-Object System.Text.UTF8Encoding($false)
$inv  = [System.Globalization.CultureInfo]::InvariantCulture

function Write-Step($msg, $color = 'Cyan') { Write-Host ">> $msg" -ForegroundColor $color }

# ---------- Config ----------
$cfgPath = Join-Path $Root 'config.json'
if (-not (Test-Path $cfgPath)) { throw "Khong tim thay config.json tai $cfgPath" }
$config = [System.IO.File]::ReadAllText($cfgPath, $utf8) | ConvertFrom-Json

$inbox   = Join-Path $Root 'inbox'
$archive = Join-Path $Root 'archive'
$dataDir = Join-Path $Root 'data'
foreach ($d in @($inbox, $archive, $dataDir)) { if (-not (Test-Path $d)) { New-Item -ItemType Directory -Path $d | Out-Null } }

# ---------- Helpers ----------
function Remove-Diacritics([string]$s) {
    if ($null -eq $s) { return '' }
    $s = $s.Replace([string][char]0x0110, 'D').Replace([string][char]0x0111, 'd')
    $norm = $s.Normalize([System.Text.NormalizationForm]::FormD)
    $sb = New-Object System.Text.StringBuilder
    foreach ($ch in $norm.ToCharArray()) {
        if ([System.Globalization.CharUnicodeInfo]::GetUnicodeCategory($ch) -ne [System.Globalization.UnicodeCategory]::NonSpacingMark) { [void]$sb.Append($ch) }
    }
    return $sb.ToString().Normalize([System.Text.NormalizationForm]::FormC)
}
function Get-Key([string]$s) { return ((Remove-Diacritics $s).ToLower() -replace '[^a-z0-9]', '') }

function Get-ColIndex([string]$ref) {
    $letters = ($ref -replace '[0-9]', '').ToUpper()
    $n = 0
    foreach ($c in $letters.ToCharArray()) { $n = $n * 26 + ([int][char]$c - 64) }
    return $n - 1
}

function Read-ZipText($zip, [string]$name) {
    $e = $zip.GetEntry($name)
    if ($null -eq $e) { return $null }
    $sr = New-Object System.IO.StreamReader($e.Open(), $utf8)
    try { return $sr.ReadToEnd() } finally { $sr.Close() }
}

# Doc sheet dau tien (hoac sheet co ten chua "BH") cua file xlsx -> mang cac dong (mang chuoi)
function Read-XlsxRows([string]$path) {
    $tmp = [System.IO.Path]::GetTempFileName()
    Copy-Item -LiteralPath $path -Destination $tmp -Force   # tranh loi khi file dang mo trong Excel
    $zip = [System.IO.Compression.ZipFile]::OpenRead($tmp)
    try {
        $shared = @()
        $ssXml = Read-ZipText $zip 'xl/sharedStrings.xml'
        if ($ssXml) {
            [xml]$ss = $ssXml
            foreach ($si in $ss.sst.si) {
                if ($si.t -ne $null) {
                    if ($si.t -is [string]) { $shared += $si.t } else { $shared += $si.t.'#text' }
                } else {
                    $txt = ''
                    foreach ($r in $si.r) { if ($r.t -is [string]) { $txt += $r.t } else { $txt += $r.t.'#text' } }
                    $shared += $txt
                }
            }
        }
        # tim sheet
        [xml]$wb = Read-ZipText $zip 'xl/workbook.xml'
        [xml]$rels = Read-ZipText $zip 'xl/_rels/workbook.xml.rels'
        $sheets = @($wb.workbook.sheets.sheet)
        $pick = $sheets | Where-Object { (Get-Key $_.name) -match 'bh' } | Select-Object -First 1
        if ($null -eq $pick) { $pick = $sheets[0] }
        $rid = $pick.GetAttribute('id', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
        $target = ($rels.Relationships.Relationship | Where-Object { $_.Id -eq $rid }).Target
        $target = $target -replace '^/xl/', '' -replace '^/', ''
        $sheetPath = if ($target.StartsWith('xl/')) { $target } else { 'xl/' + $target }
        [xml]$sx = Read-ZipText $zip $sheetPath

        $rows = New-Object System.Collections.ArrayList
        foreach ($row in $sx.worksheet.sheetData.row) {
            $vals = @{}
            $max = -1
            foreach ($c in $row.c) {
                $idx = Get-ColIndex $c.r
                $t = $c.GetAttribute('t')
                $v = $null
                if ($t -eq 's') { $v = $shared[[int]$c.v] }
                elseif ($t -eq 'inlineStr') { if ($c.is.t -is [string]) { $v = $c.is.t } else { $v = $c.is.t.'#text' } }
                else { $v = $c.v }
                if ($v -is [System.Xml.XmlElement]) { $v = $v.InnerText }
                $vals[$idx] = $v
                if ($idx -gt $max) { $max = $idx }
            }
            $arr = @()
            for ($i = 0; $i -le $max; $i++) { $arr += , $vals[$i] }
            [void]$rows.Add($arr)
        }
        return , $rows
    } finally { $zip.Dispose(); Remove-Item $tmp -Force -ErrorAction SilentlyContinue }
}

function To-Number($v) {
    if ($null -eq $v) { return 0 }
    $s = ([string]$v).Trim()
    if ($s -eq '') { return 0 }
    $d = 0.0
    if ([double]::TryParse($s, [System.Globalization.NumberStyles]::Float, $inv, [ref]$d)) { return $d }
    $s2 = $s -replace '[^0-9\-]', ''
    if ($s2 -ne '' -and [double]::TryParse($s2, [ref]$d)) { return $d }
    return 0
}

function To-IsoDate($v) {
    if ($null -eq $v) { return $null }
    $s = ([string]$v).Trim()
    if ($s -eq '') { return $null }
    $dt = [datetime]::MinValue
    $fmts = [string[]]@('dd/MM/yyyy', 'd/M/yyyy', 'dd/MM/yyyy HH:mm:ss', 'd/M/yyyy H:mm:ss', 'yyyy-MM-dd', 'yyyy-MM-ddTHH:mm:ss', 'dd-MM-yyyy')
    if ([datetime]::TryParseExact($s, $fmts, $inv, [System.Globalization.DateTimeStyles]::None, [ref]$dt)) { return $dt.ToString('yyyy-MM-dd') }
    $d = 0.0
    if ([double]::TryParse($s, [System.Globalization.NumberStyles]::Float, $inv, [ref]$d) -and $d -gt 20000 -and $d -lt 80000) {
        return [datetime]::FromOADate($d).ToString('yyyy-MM-dd')
    }
    return $null
}

# Che ten khach hang (repo GitHub de public)
$legalWords = @('cong','ty','tong','tnhh','co','phan','cp','mtv','mot','thanh','vien','chi','nhanh','van','phong','dai','dien','doanh','nghiep','tu','nhan','hop','tac','xa','htx','tap','doan','bao','hiem','pvi','va','jsc','ltd','co.,','company','limited','vietnam','viet','nam')
function Mask-Name([string]$name) {
    if (-not $config.mask_customer_names) { return $name }
    $out = @()
    foreach ($w in ($name -split '\s+')) {
        if ($w -eq '') { continue }
        $k = (Remove-Diacritics $w).ToLower().Trim('.,-_()')
        if ($legalWords -contains $k) { $out += $w }
        else { $out += ($w.Substring(0, 1) + '***') }
    }
    return ($out -join ' ')
}

# ---------- Thu thap file ----------
$files = @()
$files += Get-ChildItem -Path $archive -Filter *.xlsx -File -ErrorAction SilentlyContinue
$files += Get-ChildItem -Path $inbox   -Filter *.xlsx -File -ErrorAction SilentlyContinue
$files = $files | Where-Object { $_.Name -notlike '~$*' }
if (-not $files -or @($files).Count -eq 0) { throw "Chua co file .xlsx nao trong thu muc inbox\ . Hay copy bao cao vao do roi chay lai." }

function Get-FileSortKey($f) {
    $m = [regex]::Match($f.Name, '(\d{4})-(\d{2})-(\d{2})(?:[_\-\s\(]*(\d+))?')
    if ($m.Success) {
        $suffix = if ($m.Groups[4].Success) { [int]$m.Groups[4].Value } else { 0 }
        return ('{0}{1}{2}-{3:D4}-{4}' -f $m.Groups[1].Value, $m.Groups[2].Value, $m.Groups[3].Value, $suffix, $f.LastWriteTime.ToString('yyyyMMddHHmmss'))
    }
    return ('{0}-0000-{1}' -f $f.LastWriteTime.ToString('yyyyMMdd'), $f.LastWriteTime.ToString('yyyyMMddHHmmss'))
}
$files = @($files | Sort-Object { Get-FileSortKey $_ })
$inboxFiles = @(Get-ChildItem -Path $inbox -Filter *.xlsx -File -ErrorAction SilentlyContinue | Where-Object { $_.Name -notlike '~$*' })
$archNames = @(Get-ChildItem -Path $archive -Filter *.xlsx -File -ErrorAction SilentlyContinue | ForEach-Object { $_.Name })
Write-Step ("Tim thay {0} file bao cao" -f $files.Count)

# Ten cot -> khoa chuan (so sanh khong dau)
$colMap = [ordered]@{
    stt = 'stt'; soct = 'so_ct'; ngayct = 'ngay_ct'; sodonbh = 'so_don'; khachhang = 'khach_hang'; masp = 'ma_sp';
    loaitien = 'loai_tien'; sotienbh = 'stbh'; phibh = 'phi'; ngaybatdau = 'ngay_bd'; ngayketthuc = 'ngay_kt';
    cbkhaithac = 'cb'; slfiledinhkem = 'so_file'; trangthai = 'trang_thai'
}

$records = [ordered]@{}
$fileLog = @()
$reportDate = $null
foreach ($f in $files) {
    $rows = Read-XlsxRows $f.FullName
    $hdrIdx = -1
    for ($i = 0; $i -lt [Math]::Min($rows.Count, 30); $i++) {
        $keys = @($rows[$i] | ForEach-Object { Get-Key ([string]$_) })
        if ($keys -contains 'sodonbh' -and $keys -contains 'phibh') { $hdrIdx = $i; break }
    }
    if ($hdrIdx -lt 0) { Write-Step ("Bo qua {0}: khong thay dong tieu de (So don BH / Phi BH)" -f $f.Name) 'Yellow'; continue }
    $hdr = @($rows[$hdrIdx] | ForEach-Object { Get-Key ([string]$_) })
    $pos = @{}
    for ($j = 0; $j -lt $hdr.Count; $j++) { if ($colMap.Contains($hdr[$j])) { $pos[$colMap[$hdr[$j]]] = $j } }

    $n = 0
    for ($i = $hdrIdx + 1; $i -lt $rows.Count; $i++) {
        $r = $rows[$i]
        $get = { param($k) if ($pos.ContainsKey($k) -and $pos[$k] -lt $r.Count) { $r[$pos[$k]] } else { $null } }
        $soDon = [string](& $get 'so_don')
        if ([string]::IsNullOrWhiteSpace($soDon)) { continue }
        $soDon = $soDon.Trim()
        $kh = ([string](& $get 'khach_hang')).Trim()
        $isTest = $false
        foreach ($p in $config.exclude_name_patterns) { if ((Remove-Diacritics $kh).ToUpper().Contains((Remove-Diacritics $p).ToUpper())) { $isTest = $true } }
        $ngayCt = To-IsoDate (& $get 'ngay_ct')
        if ($config.count_from -and $ngayCt -and $ngayCt -lt [string]$config.count_from) { $isTest = $true }
        $isSdbs = ($soDon -match '(?i)S\S*BS' -or $soDon -match '/E\d+')
        $rec = [ordered]@{
            so_ct      = ([string](& $get 'so_ct')).Trim()
            ngay_ct    = $ngayCt
            so_don     = $soDon
            goc        = (($soDon -replace '(?i)\s*S\S*BS\s*$', '') -replace '/E\d+$', '').Trim()
            khach_hang = Mask-Name $kh
            ma_sp      = ([string](& $get 'ma_sp')).Trim()
            stbh       = To-Number (& $get 'stbh')
            phi        = To-Number (& $get 'phi')
            ngay_bd    = To-IsoDate (& $get 'ngay_bd')
            ngay_kt    = To-IsoDate (& $get 'ngay_kt')
            cb         = ([string](& $get 'cb')).Trim()
            so_file    = [int](To-Number (& $get 'so_file'))
            trang_thai = ([string](& $get 'trang_thai')).Trim()
            sdbs       = [bool]$isSdbs
            test       = [bool]$isTest
            src        = $f.Name
        }
        $records[$soDon] = $rec    # file moi hon ghi de file cu
        $n++
    }
    $m = [regex]::Match($f.Name, '(\d{4})-(\d{2})-(\d{2})')
    $fd = if ($m.Success) { '{0}-{1}-{2}' -f $m.Groups[1].Value, $m.Groups[2].Value, $m.Groups[3].Value } else { $f.LastWriteTime.ToString('yyyy-MM-dd') }
    if ($null -eq $reportDate -or $fd -gt $reportDate) { $reportDate = $fd }
    $fileLog += [ordered]@{ file = $f.Name; rows = $n; report_date = $fd }
    Write-Step ("  {0}: {1} dong" -f $f.Name, $n) 'Gray'
}
if ($records.Count -eq 0) { throw "Khong doc duoc dong du lieu nao." }

# ---------- Xuat ----------
$recList = New-Object System.Collections.ArrayList
foreach ($k in $records.Keys) { [void]$recList.Add($records[$k]) }
$recList = @($recList | Sort-Object @{ e = { $_.ngay_ct } ; Descending = $true }, @{ e = { $_.so_don } ; Descending = $true })

function Confirm-Step([string]$msg) {
    if ($Force) { return $true }
    Write-Host ''
    Write-Host ("  [?] " + $msg) -ForegroundColor Yellow
    $a = Read-Host '      Van tiep tuc cap nhat? (Y = co / N = huy)'
    return ($a -match '^\s*[yYcC]')
}
function Stop-Cancel { Write-Step 'Da huy - khong cap nhat gi. File van nam trong inbox.' 'Yellow'; exit 2 }

# ---------- Chong chay nham ----------
$dupe = @($inboxFiles | Where-Object { $archNames -contains $_.Name })
if ($dupe.Count -gt 0) {
    if (-not (Confirm-Step ("File da tung xu ly truoc day: " + (($dupe | ForEach-Object { $_.Name }) -join ', ')))) { Stop-Cancel }
}
$latestArch = $null
foreach ($n in $archNames) { $mm = [regex]::Match($n, '(\d{4})-(\d{2})-(\d{2})'); if ($mm.Success) { $d = $mm.Value; if ($null -eq $latestArch -or $d -gt $latestArch) { $latestArch = $d } } }
foreach ($f in $inboxFiles) {
    $mm = [regex]::Match($f.Name, '(\d{4})-(\d{2})-(\d{2})')
    if ($mm.Success -and $latestArch -and $mm.Value -lt $latestArch) {
        if (-not (Confirm-Step ("File {0} co ngay CU HON bao cao gan nhat ({1})." -f $f.Name, $latestArch))) { Stop-Cancel }
    }
}
$prevJsonPath = Join-Path $dataDir 'data.json'
if ($inboxFiles.Count -gt 0 -and (Test-Path $prevJsonPath)) {
    try {
        $prev = [System.IO.File]::ReadAllText($prevJsonPath, $utf8) | ConvertFrom-Json
        $prevSig = ((@($prev.records) | ForEach-Object { '{0}|{1}|{2}' -f $_.so_don, ([double]$_.phi), $_.trang_thai } | Sort-Object) -join ';')
        $sig2 = (($recList | ForEach-Object { '{0}|{1}|{2}' -f $_.so_don, ([double]$_.phi), $_.trang_thai } | Sort-Object) -join ';')
        if ($prevSig -eq $sig2) {
            if (-not (Confirm-Step 'Bao cao moi KHONG co thay doi nao so voi lan cap nhat truoc (khong co don moi).')) { Stop-Cancel }
        }
    } catch { }
}

# ---------- Tong hop so lieu ----------
$vi = [System.Globalization.CultureInfo]::GetCultureInfo('vi-VN')
function F([double]$n) { return $n.ToString('N0', $vi) }
function S([double]$n) {
    if ([Math]::Abs($n) -ge 1e9) { return ($n / 1e9).ToString('0.##', $vi) + ' ty' }
    if ([Math]::Abs($n) -ge 1e6) { return ($n / 1e6).ToString('0.#', $vi) + ' tr' }
    return F $n
}
$rep = [datetime]::ParseExact($reportDate, 'yyyy-MM-dd', $inv)
$t1 = $rep.AddDays(-1)
function K([datetime]$d) { return $d.ToString('yyyy-MM-dd') }
$keyDay = K $t1
$monStart = New-Object datetime($t1.Year, $t1.Month, 1)
$prevStart = $monStart.AddMonths(-1)
$prevEnd = $monStart.AddDays(-1)
$sameDay = [Math]::Min($t1.Day, [datetime]::DaysInMonth($prevStart.Year, $prevStart.Month))
$ranges = [ordered]@{
    day        = @((K $t1), (K $t1))
    prev_day   = @((K $t1.AddDays(-1)), (K $t1.AddDays(-1)))
    week       = @((K $t1.AddDays(-6)), (K $t1))
    prev_week  = @((K $t1.AddDays(-13)), (K $t1.AddDays(-7)))
    month      = @((K $monStart), (K $t1))
    prev_same  = @((K $prevStart), (K $prevStart.AddDays($sameDay - 1)))
    prev       = @((K $prevStart), (K $prevEnd))
    total      = @('0000-00-00', '9999-12-31')
}
$okStatus = @($config.revenue_statuses | ForEach-Object { Get-Key $_ })
function Test-Ok($r) { return ($okStatus.Count -eq 0 -or $okStatus -contains (Get-Key $r.trang_thai)) }
$live = @($recList | Where-Object { -not $_.test })
$newRecs = @($live | Where-Object { -not $_.sdbs -and (Test-Ok $_) })
$sdRecs  = @($live | Where-Object { $_.sdbs })
$addSd = [bool]$config.count_endorsement_premium

function Get-Agg($list, $rg, [bool]$withSd) {
    $n = 0; $phi = 0.0
    foreach ($r in $list) { $d = [string]$r.ngay_ct; if ($d -ge $rg[0] -and $d -le $rg[1]) { $n++; $phi += $r.phi } }
    if ($withSd) { foreach ($r in $sdRecs) { $d = [string]$r.ngay_ct; if ((Test-Ok $r) -and $d -ge $rg[0] -and $d -le $rg[1]) { $phi += $r.phi } } }
    return [ordered]@{ n = $n; phi = $phi; from = $rg[0]; to = $rg[1] }
}

# che do tin nhan: ngay 1 -> thang, thu Hai -> tuan, con lai -> ngay
$mode = 'daily'
if ($rep.Day -eq 1) { $mode = 'monthly' } elseif ($rep.DayOfWeek -eq [DayOfWeek]::Monday) { $mode = 'weekly' }

$summary = [ordered]@{
    mode = $mode; day_key = $keyDay; month_key = $t1.ToString('yyyy-MM'); prev_key = $prevStart.ToString('yyyy-MM')
    count_from = [string]$config.count_from
    first_date = (@($newRecs | ForEach-Object { $_.ngay_ct } | Where-Object { $_ } | Sort-Object) | Select-Object -First 1)
    sdbs_count = $sdRecs.Count
    test_count = @($recList | Where-Object { $_.test }).Count
    si_total = [double](($newRecs | ForEach-Object { $_.stbh } | Measure-Object -Sum).Sum)
}
foreach ($k in $ranges.Keys) { $summary[$k] = Get-Agg $newRecs $ranges[$k] $addSd }

# du bao cuoi thang (can >= 3 ngay du lieu trong thang)
$startForecast = $monStart
if ($config.count_from) { $cf = [datetime]::ParseExact([string]$config.count_from, 'yyyy-MM-dd', $inv); if ($cf -gt $startForecast) { $startForecast = $cf } }
$elapsed = ($t1 - $startForecast).Days + 1
$periodDays = ($monStart.AddMonths(1).AddDays(-1) - $startForecast).Days + 1
$summary['forecast'] = $null
if ($elapsed -ge 3 -and $summary.month.n -gt 0 -and $periodDays -gt 0) {
    $summary['forecast'] = [ordered]@{ phi = [Math]::Round($summary.month.phi / $elapsed * $periodDays); n = [Math]::Round($summary.month.n / $elapsed * $periodDays); days = $elapsed }
}

$pkgOut = @()
$matched = @{}
foreach ($pk in $config.packages) {
    $list = @($newRecs | Where-Object { [double]$_.stbh -eq [double]$pk.stbh })
    foreach ($r in $list) { $matched[$r.so_don] = $true }
    $o = [ordered]@{ name = $pk.name; stbh = $pk.stbh; premium = $pk.premium }
    foreach ($k in $ranges.Keys) { $o[$k] = Get-Agg $list $ranges[$k] $false }
    $pkgOut += $o
}
$others = @($newRecs | Where-Object { -not $matched.ContainsKey($_.so_don) })
if ($others.Count -gt 0) {
    $o = [ordered]@{ name = 'Khac'; stbh = 0; premium = 0 }
    foreach ($k in $ranges.Keys) { $o[$k] = Get-Agg $others $ranges[$k] $false }
    $pkgOut += $o
}
$summary['packages'] = [object[]]$pkgOut

# ---------- Kiem tra du lieu bat thuong (chi don that) ----------
$alerts = @()
foreach ($r in $live) {
    if (-not (Test-Ok $r)) { $alerts += [ordered]@{ code = 'status'; so_don = $r.so_don; detail = $r.trang_thai } }
    if (-not $r.sdbs) {
        $pk = @($config.packages | Where-Object { [double]$_.stbh -eq [double]$r.stbh }) | Select-Object -First 1
        if ($null -eq $pk) { $alerts += [ordered]@{ code = 'package'; so_don = $r.so_don; detail = (F $r.stbh) } }
        elseif ([double]$r.phi -ne [double]$pk.premium) { $alerts += [ordered]@{ code = 'premium'; so_don = $r.so_don; detail = ('{0} / {1}' -f (F $r.phi), (F $pk.premium)) } }
        if ($r.so_file -le 0) { $alerts += [ordered]@{ code = 'nofile'; so_don = $r.so_don; detail = '' } }
        if ($r.ngay_bd -and $r.ngay_ct -and $r.ngay_bd -lt $r.ngay_ct) { $alerts += [ordered]@{ code = 'backdate'; so_don = $r.so_don; detail = $r.ngay_bd } }
    }
}
$summary['alerts'] = [object[]]$alerts

# ---------- Tin nhan Zalo ----------
$tplName = @{ daily = 'message_template.txt'; weekly = 'message_template_weekly.txt'; monthly = 'message_template_monthly.txt' }[$mode]
$tplPath = Join-Path $Root $tplName
if (-not (Test-Path $tplPath)) { $tplPath = Join-Path $Root 'message_template.txt' }
$message = ''
if (Test-Path $tplPath) {
    $tpl = [System.IO.File]::ReadAllText($tplPath, $utf8).Replace("`r`n", "`n")
    $fc = $summary.forecast
    $map = [ordered]@{
        '{NGAY}' = $t1.ToString('dd/MM/yyyy'); '{NGAY_NGAN}' = $t1.ToString('dd/MM'); '{NGAY_TRUOC}' = $t1.AddDays(-1).ToString('dd/MM')
        '{TUAN_TU}' = $t1.AddDays(-6).ToString('dd/MM'); '{TUAN_DEN}' = $t1.ToString('dd/MM')
        '{THANG}' = $t1.ToString('MM/yyyy'); '{THANG_TRUOC}' = $prevStart.ToString('MM/yyyy')
        '{N_NGAY}' = F $summary.day.n; '{PHI_NGAY}' = F $summary.day.phi
        '{N_NGAY_TRUOC}' = F $summary.prev_day.n; '{PHI_NGAY_TRUOC}' = F $summary.prev_day.phi
        '{N_TUAN}' = F $summary.week.n; '{PHI_TUAN}' = F $summary.week.phi
        '{N_TUAN_TRUOC}' = F $summary.prev_week.n; '{PHI_TUAN_TRUOC}' = F $summary.prev_week.phi
        '{N_THANG}' = F $summary.month.n; '{PHI_THANG}' = F $summary.month.phi
        '{N_CUNG_KY}' = F $summary.prev_same.n; '{PHI_CUNG_KY}' = F $summary.prev_same.phi
        '{N_THANG_TRUOC}' = F $summary.prev.n; '{PHI_THANG_TRUOC}' = F $summary.prev.phi
        '{N_LUYKE}' = F $summary.total.n; '{PHI_LUYKE}' = F $summary.total.phi
        '{N_SDBS}' = F $summary.sdbs_count; '{LINK}' = [string]$config.dashboard_url
        '{DU_BAO}' = $(if ($fc) { ($fc.phi / 1e6).ToString('0.#', $vi) + ' tri' + [char]0x1EC7 + 'u' } else { '' })
        '{DU_BAO_N}' = $(if ($fc) { F $fc.n } else { '' })
    }
    $outLines = @()
    foreach ($line in ($tpl -split "`n")) {
        if ($line -match '\{GOI_') {
            foreach ($pk in $pkgOut) {
                $l = $line.Replace('{GOI_TEN}', $pk.name)
                foreach ($pair in @(@('NGAY','day'), @('TUAN','week'), @('THANG_TRUOC','prev'), @('THANG','month'), @('LUYKE','total'))) {
                    $l = $l.Replace('{GOI_N_' + $pair[0] + '}', (F $pk[$pair[1]].n)).Replace('{GOI_PHI_' + $pair[0] + '}', (F $pk[$pair[1]].phi))
                }
                $outLines += $l
            }
            continue
        }
        $l = $line; $skip = $false
        foreach ($k in $map.Keys) { if ($l.Contains($k)) { if ($map[$k] -eq '') { $skip = $true }; $l = $l.Replace($k, $map[$k]) } }
        if (-not $skip) { $outLines += $l }
    }
    $message = ($outLines -join "`r`n").TrimEnd()
    [System.IO.File]::WriteAllText((Join-Path $dataDir 'message.txt'), $message, $utf8)
}

$payload = [ordered]@{
    generated_at = (Get-Date).ToString('yyyy-MM-ddTHH:mm:ss')
    report_date  = $reportDate
    config       = $config
    summary      = $summary
    message      = $message
    files        = [object[]]$fileLog
    records      = [object[]]$recList
}
$json = ConvertTo-Json -InputObject $payload -Depth 8 -Compress
[System.IO.File]::WriteAllText((Join-Path $dataDir 'data.json'), $json, $utf8)
[System.IO.File]::WriteAllText((Join-Path $dataDir 'data.js'), "window.DASH_DATA = $json;`n", $utf8)

# chuyen file inbox -> archive
if (-not $NoArchive) {
    foreach ($f in $inboxFiles) {
        $dest = Join-Path $archive $f.Name
        if (Test-Path $dest) { Remove-Item $dest -Force }
        Move-Item -LiteralPath $f.FullName -Destination $dest
    }
}

Write-Step ("Xong. Ngay {0}: {1} GCN / {2} d | Thang: {3} GCN | Luy ke: {4} GCN / {5} d | {6} ban ghi test da loai" -f $t1.ToString('dd/MM'), $summary.day.n, (F $summary.day.phi), $summary.month.n, $summary.total.n, (F $summary.total.phi), $summary.test_count) 'Green'
if ($mode -ne 'daily') { Write-Step ("Hom nay dung mau tin nhan " + $(if ($mode -eq 'weekly') { 'TONG KET TUAN' } else { 'TONG KET THANG' })) 'Cyan' }
if ($alerts.Count -gt 0) {
    Write-Step ("CANH BAO: {0} diem du lieu can kiem tra (xem muc 'Kiem tra du lieu' tren dashboard):" -f $alerts.Count) 'Yellow'
    $labels = @{ status = 'Chua duyet'; package = 'STBH khong khop goi nao'; premium = 'Phi khac bieu phi (thuc te / bieu phi)'; nofile = 'Khong co file dinh kem'; backdate = 'Hieu luc truoc ngay cap' }
    foreach ($a in ($alerts | Select-Object -First 10)) { Write-Host ("     - {0}: {1} {2}" -f $a.so_don, $labels[$a.code], $a.detail) -ForegroundColor Yellow }
}

# ---------- Anh bao cao (Edge / Chrome headless) ----------
$global:ReportImage = $null
if (-not $NoImage) {
    function JP($a, $b) { if ($a) { return (Join-Path $a $b) } return $null }
    $cands = @(
        (JP ${env:ProgramFiles(x86)} 'Microsoft\Edge\Application\msedge.exe'),
        (JP $env:ProgramFiles 'Microsoft\Edge\Application\msedge.exe'),
        (JP $env:LOCALAPPDATA 'Microsoft\Edge\Application\msedge.exe'),
        (JP $env:ProgramFiles 'Google\Chrome\Application\chrome.exe'),
        (JP ${env:ProgramFiles(x86)} 'Google\Chrome\Application\chrome.exe'),
        (JP $env:LOCALAPPDATA 'Google\Chrome\Application\chrome.exe')
    )
    if ($env:REPORT_BROWSER) { $cands = @($env:REPORT_BROWSER) + $cands }
    $browser = $cands | Where-Object { $_ -and (Test-Path $_) } | Select-Object -First 1
    $reportHtml = Join-Path $Root 'report.html'
    if ($browser -and (Test-Path $reportHtml)) {
        $outDir = Join-Path $Root 'output'
        if (-not (Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir | Out-Null }
        $png = Join-Path $outDir ('bao-cao-{0}.png' -f $keyDay)
        if (Test-Path $png) { Remove-Item $png -Force }
        $prof = Join-Path ([System.IO.Path]::GetTempPath()) 'antamso-headless'
        $argLine = ('--headless=new --disable-gpu --hide-scrollbars --no-first-run --no-default-browser-check --force-device-scale-factor=2 --window-size=540,1240 --virtual-time-budget=6000 "--user-data-dir={0}" "--screenshot={1}" "{2}"' -f $prof, $png, $reportHtml)
        try {
            try { $proc = Start-Process -FilePath $browser -ArgumentList $argLine -PassThru -WindowStyle Hidden -ErrorAction Stop }
            catch { $proc = Start-Process -FilePath $browser -ArgumentList $argLine -PassThru }
            if (-not $proc.WaitForExit(45000)) { try { $proc.Kill() } catch { } }
        } catch { }
        if (Test-Path $png) {
            # cat dung khung 540x1080 (x2) - phan thua do cua so headless
            if ($env:OS -eq 'Windows_NT') {
                try {
                    Add-Type -AssemblyName System.Drawing
                    $bmpT = 'System.Drawing.Bitmap' -as [type]; $fmtT = 'System.Drawing.Imaging.ImageFormat' -as [type]
                    $src = $bmpT::FromFile($png)
                    $h = [Math]::Min($src.Height, 2160); $w = [Math]::Min($src.Width, 1080)
                    $rect = New-Object -TypeName 'System.Drawing.Rectangle' -ArgumentList 0, 0, $w, $h
                    $dst = $src.Clone($rect, $src.PixelFormat)
                    $src.Dispose()
                    $dst.Save($png, $fmtT::Png); $dst.Dispose()
                } catch { }
            }
        }
        if (Test-Path $png) { $global:ReportImage = $png; Write-Step ("Da tao anh bao cao: output\" + (Split-Path $png -Leaf)) 'Green' }
        else { Write-Step 'Khong tao duoc anh bao cao (bo qua).' 'Yellow' }
    } else { Write-Step 'Khong tim thay Edge/Chrome de tao anh bao cao (bo qua).' 'Yellow' }
}

# ---------- Clipboard ----------
if ($message -and -not $NoClipboard) {
    $copied = $false
    try { Set-Clipboard -Value $message -ErrorAction Stop; $copied = $true } catch {
        try { Add-Type -AssemblyName System.Windows.Forms; [System.Windows.Forms.Clipboard]::SetText($message); $copied = $true } catch { }
    }
    if ($copied) { Write-Step 'Da chep tin nhan bao cao vao clipboard - vao nhom Zalo bam Ctrl+V' 'Green' }
    else { Write-Step 'Khong chep duoc vao clipboard - mo file data\message.txt de copy' 'Yellow' }
    if ($global:ReportImage -and -not $NoImage) { try { Start-Process explorer.exe ('/select,"{0}"' -f $global:ReportImage) } catch { } }
}
