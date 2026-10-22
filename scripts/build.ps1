# =====================================================================
#  An Tam So / Cyber Shield SME - BH PVI x MISA
#  build.ps1 : doc file Excel bao cao raw trong /inbox (+ /archive),
#              gop - loai trung - xuat data/data.js cho dashboard.
#  Khong can cai Python / Excel. Chay duoc tren Windows PowerShell 5.1+.
# =====================================================================
param(
    [string]$Root = '',
    [switch]$NoArchive,
    [switch]$NoClipboard
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
        $isSdbs = ($soDon -match '(?i)S\S*BS' -or $soDon -match '/E\d+')
        $rec = [ordered]@{
            so_ct      = ([string](& $get 'so_ct')).Trim()
            ngay_ct    = To-IsoDate (& $get 'ngay_ct')
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

# ---------- Tong hop so lieu ----------
$vi = [System.Globalization.CultureInfo]::GetCultureInfo('vi-VN')
function F([double]$n) { return $n.ToString('N0', $vi) }
$t1 = ([datetime]::ParseExact($reportDate, 'yyyy-MM-dd', $inv)).AddDays(-1)
$keyDay = $t1.ToString('yyyy-MM-dd')
$keyMon = $t1.ToString('yyyy-MM')
$keyPrev = $t1.AddMonths(-1).ToString('yyyy-MM')
$okStatus = @($config.revenue_statuses | ForEach-Object { Get-Key $_ })
function Test-Ok($r) { return ($okStatus.Count -eq 0 -or $okStatus -contains (Get-Key $r.trang_thai)) }

$newRecs = @($recList | Where-Object { -not $_.test -and -not $_.sdbs -and (Test-Ok $_) })
$sdRecs  = @($recList | Where-Object { -not $_.test -and $_.sdbs })
$addSd = [bool]$config.count_endorsement_premium

function In-Period($r, [string]$period) {
    $d = [string]$r.ngay_ct
    switch ($period) {
        'day'   { return ($d -eq $keyDay) }
        'month' { return ($d.Length -ge 7 -and $d.Substring(0, 7) -eq $keyMon -and $d -le $keyDay) }
        'prev'  { return ($d.Length -ge 7 -and $d.Substring(0, 7) -eq $keyPrev) }
        default { return $true }
    }
}
function Get-Agg($list, [string]$period) {
    $n = 0; $phi = 0.0
    foreach ($r in $list) { if (In-Period $r $period) { $n++; $phi += $r.phi } }
    if ($addSd) { foreach ($r in $sdRecs) { if ((Test-Ok $r) -and (In-Period $r $period)) { $phi += $r.phi } } }
    return [ordered]@{ n = $n; phi = $phi }
}
$periods = @('day', 'month', 'prev', 'total')
$summary = [ordered]@{
    day_key = $keyDay; month_key = $keyMon; prev_key = $keyPrev
    first_date = (@($newRecs | ForEach-Object { $_.ngay_ct } | Where-Object { $_ } | Sort-Object) | Select-Object -First 1)
    sdbs_count = $sdRecs.Count
    test_count = @($recList | Where-Object { $_.test }).Count
    si_total = [double](($newRecs | ForEach-Object { $_.stbh } | Measure-Object -Sum).Sum)
}
foreach ($p in $periods) { $summary[$p] = Get-Agg $newRecs $p }
$pkgOut = @()
$matched = @{}
foreach ($pk in $config.packages) {
    $list = @($newRecs | Where-Object { [double]$_.stbh -eq [double]$pk.stbh })
    foreach ($r in $list) { $matched[$r.so_don] = $true }
    $o = [ordered]@{ name = $pk.name; stbh = $pk.stbh; premium = $pk.premium }
    $saveSd = $addSd; $addSd = $false
    foreach ($p in $periods) { $o[$p] = Get-Agg $list $p }
    $addSd = $saveSd
    $pkgOut += $o
}
$others = @($newRecs | Where-Object { -not $matched.ContainsKey($_.so_don) })
if ($others.Count -gt 0) {
    $o = [ordered]@{ name = 'Khac'; stbh = 0; premium = 0 }
    $saveSd = $addSd; $addSd = $false
    foreach ($p in $periods) { $o[$p] = Get-Agg $others $p }
    $addSd = $saveSd
    $pkgOut += $o
}
$summary['packages'] = [object[]]$pkgOut

# ---------- Tin nhan Zalo (tu mau message_template.txt) ----------
$message = ''
$tplPath = Join-Path $Root 'message_template.txt'
if (Test-Path $tplPath) {
    $tpl = [System.IO.File]::ReadAllText($tplPath, $utf8).Replace("`r`n", "`n")
    $map = @{
        '{NGAY}' = $t1.ToString('dd/MM/yyyy'); '{NGAY_NGAN}' = $t1.ToString('dd/MM')
        '{THANG}' = $t1.ToString('MM/yyyy'); '{THANG_TRUOC}' = $t1.AddMonths(-1).ToString('MM/yyyy')
        '{N_NGAY}' = F $summary.day.n; '{PHI_NGAY}' = F $summary.day.phi
        '{N_THANG}' = F $summary.month.n; '{PHI_THANG}' = F $summary.month.phi
        '{N_THANG_TRUOC}' = F $summary.prev.n; '{PHI_THANG_TRUOC}' = F $summary.prev.phi
        '{N_LUYKE}' = F $summary.total.n; '{PHI_LUYKE}' = F $summary.total.phi
        '{N_SDBS}' = F $summary.sdbs_count; '{LINK}' = [string]$config.dashboard_url
    }
    $outLines = @()
    foreach ($line in ($tpl -split "`n")) {
        if ($line -match '\{GOI_') {
            foreach ($pk in $pkgOut) {
                $l = $line.Replace('{GOI_TEN}', $pk.name)
                $l = $l.Replace('{GOI_N_NGAY}', (F $pk.day.n)).Replace('{GOI_PHI_NGAY}', (F $pk.day.phi))
                $l = $l.Replace('{GOI_N_THANG}', (F $pk.month.n)).Replace('{GOI_PHI_THANG}', (F $pk.month.phi))
                $l = $l.Replace('{GOI_N_THANG_TRUOC}', (F $pk.prev.n)).Replace('{GOI_PHI_THANG_TRUOC}', (F $pk.prev.phi))
                $l = $l.Replace('{GOI_N_LUYKE}', (F $pk.total.n)).Replace('{GOI_PHI_LUYKE}', (F $pk.total.phi))
                $outLines += $l
            }
        } else {
            $l = $line
            foreach ($k in $map.Keys) { $l = $l.Replace($k, $map[$k]) }
            $outLines += $l
        }
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

# chep tin nhan vao clipboard
if ($message -and -not $NoClipboard) {
    $copied = $false
    try { Set-Clipboard -Value $message -ErrorAction Stop; $copied = $true } catch {
        try { Add-Type -AssemblyName System.Windows.Forms; [System.Windows.Forms.Clipboard]::SetText($message); $copied = $true } catch { }
    }
    if ($copied) { Write-Step 'Da chep tin nhan bao cao vao clipboard - mo Zalo va bam Ctrl+V' 'Green' }
    else { Write-Step ('Khong chep duoc vao clipboard - mo file data\message.txt de copy') 'Yellow' }
}

# chuyen file inbox -> archive
if (-not $NoArchive) {
    foreach ($f in (Get-ChildItem -Path $inbox -Filter *.xlsx -File -ErrorAction SilentlyContinue)) {
        if ($f.Name -like '~$*') { continue }
        $dest = Join-Path $archive $f.Name
        if (Test-Path $dest) { Remove-Item $dest -Force }
        Move-Item -LiteralPath $f.FullName -Destination $dest
    }
}

Write-Step ("Xong. Ngay bao cao {0} | Ngay {1}: {2} GCN / {3} d | Luy ke: {4} GCN / {5} d" -f $reportDate, $t1.ToString('dd/MM'), $summary.day.n, (F $summary.day.phi), $summary.total.n, (F $summary.total.phi)) 'Green'
