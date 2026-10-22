# =====================================================================
#  An Tam So / Cyber Shield SME - BH PVI x MISA
#  build.ps1 : doc file Excel bao cao raw trong /inbox (+ /archive),
#              gop - loai trung - xuat data/data.js cho dashboard.
#  Khong can cai Python / Excel. Chay duoc tren Windows PowerShell 5.1+.
# =====================================================================
param(
    [string]$Root = '',
    [switch]$NoArchive
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

$payload = [ordered]@{
    generated_at = (Get-Date).ToString('yyyy-MM-ddTHH:mm:ss')
    report_date  = $reportDate
    config       = $config
    files        = [object[]]$fileLog
    records      = [object[]]$recList
}
$json = ConvertTo-Json -InputObject $payload -Depth 8 -Compress
[System.IO.File]::WriteAllText((Join-Path $dataDir 'data.json'), $json, $utf8)
[System.IO.File]::WriteAllText((Join-Path $dataDir 'data.js'), "window.DASH_DATA = $json;`n", $utf8)

# chuyen file inbox -> archive
if (-not $NoArchive) {
    foreach ($f in (Get-ChildItem -Path $inbox -Filter *.xlsx -File -ErrorAction SilentlyContinue)) {
        if ($f.Name -like '~$*') { continue }
        $dest = Join-Path $archive $f.Name
        if (Test-Path $dest) { Remove-Item $dest -Force }
        Move-Item -LiteralPath $f.FullName -Destination $dest
    }
}

$valid = @($recList | Where-Object { -not $_.test -and -not $_.sdbs })
$sum = 0; foreach ($v in $valid) { $sum += $v.phi }
Write-Step ("Xong. Ngay bao cao {0} | {1} ban ghi | {2} don cap moi (khong tinh test) | Doanh phi {3:N0} d" -f $reportDate, $recList.Count, $valid.Count, $sum) 'Green'
