param(
  [string]$Root
)

# QC Excel 处理工具
# 功能概览：读取 qc_export/*.xls，按日期对齐，输出到 templet.xls。
# 缺失值支持两种模式：留空 / 用最接近均值的原始值填充。
# 历史下拉项、模板和排除关键词通过 qc_config.json 持久化。
if ([string]::IsNullOrWhiteSpace($Root)) {
  if ($PSScriptRoot) {
    $Root = $PSScriptRoot
  } else {
    try {
      $exePath = [System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
      $Root = [System.IO.Path]::GetDirectoryName($exePath)
    } catch {
      $Root = [System.AppDomain]::CurrentDomain.BaseDirectory
    }
  }
}

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type @"
using System;
using System.Runtime.InteropServices;
public static class Win32Icon {
  [DllImport("user32.dll", CharSet = CharSet.Auto)]
  public static extern bool DestroyIcon(IntPtr handle);
}
"@

$ConfigPath = Join-Path $Root "qc_config.json"
$ExcludePath = Join-Path $Root "exclude.txt"
$ExportDir = Join-Path $Root "qc_export"
$TemplatePath = Join-Path $Root "templet.xls"
$OutputDir = Join-Path $Root "output"
$IconPath = Join-Path $Root "icon.png"

function Set-WindowIconFromPng($form, [string]$pngPath) {
  if ($null -eq $form) { return }
  if ([string]::IsNullOrWhiteSpace($pngPath)) { return }
  if (-not (Test-Path -LiteralPath $pngPath)) { return }

  $bmp = $null
  $icon = $null
  $hIcon = [IntPtr]::Zero
  try {
    $bmp = New-Object System.Drawing.Bitmap($pngPath)
    $hIcon = $bmp.GetHicon()
    $icon = [System.Drawing.Icon]::FromHandle($hIcon)
    # Clone before destroying native handle; Form.Icon should keep the managed clone.
    $form.Icon = [System.Drawing.Icon]$icon.Clone()
  } catch {
    # Non-fatal: app should still work even if icon loading fails.
  } finally {
    if ($null -ne $icon) { $icon.Dispose() }
    if ($null -ne $bmp) { $bmp.Dispose() }
    if ($hIcon -ne [IntPtr]::Zero) { [Win32Icon]::DestroyIcon($hIcon) | Out-Null }
  }
}

# 读取配置对象。
# - instruments: 仪器历史
# - batches: 批号历史
# - batch_templates: 模板列表
# - batch_template_selected / batch_template_default: 模板状态
# - exclude_keywords: 排除关键词（与 exclude.txt 等价）
function Read-Config {
  if (Test-Path -LiteralPath $ConfigPath) {
    try {
      return (Get-Content -LiteralPath $ConfigPath -Raw | ConvertFrom-Json)
    } catch {
      return [pscustomobject]@{
        instruments = @()
        batches = @()
        batch_templates = @("常规化学{year}{value}")
        batch_template_selected = "常规化学{year}{value}"
        batch_template_default = "常规化学{year}{value}"
        exclude_keywords = @()
      }
    }
  }
  return [pscustomobject]@{
    instruments = @()
    batches = @()
    batch_templates = @("常规化学{year}{value}")
    batch_template_selected = "常规化学{year}{value}"
    batch_template_default = "常规化学{year}{value}"
    exclude_keywords = @()
  }
}

function Try-GetProp($obj, [string]$name) {
  if ($null -eq $obj) { return $null }
  $prop = $obj.PSObject.Properties[$name]
  if ($null -eq $prop) { return $null }
  return $prop.Value
}

# 保存配置：自动去重并确保默认模板始终存在。
function Save-Config($cfg) {
  try {
    $inst = @()
    $bats = @()
    $excludes = @()
    $templates = @("常规化学{year}{value}")
    $selected = "常规化学{year}{value}"
    $default = "常规化学{year}{value}"
    if ($null -ne $cfg) {
      $instRaw = Try-GetProp $cfg "instruments"
      if ($null -ne $instRaw) {
        foreach ($i in $instRaw) {
          if (-not [string]::IsNullOrWhiteSpace([string]$i)) { $inst += [string]$i }
        }
      }
      $batchRaw = Try-GetProp $cfg "batches"
      if ($null -ne $batchRaw) {
        foreach ($b in $batchRaw) {
          if (-not [string]::IsNullOrWhiteSpace([string]$b)) { $bats += [string]$b }
        }
      }
      $excludeRaw = Try-GetProp $cfg "exclude_keywords"
      if ($null -ne $excludeRaw) {
        foreach ($e in $excludeRaw) {
          if (-not [string]::IsNullOrWhiteSpace([string]$e)) { $excludes += [string]$e }
        }
      }
      $tplRaw = Try-GetProp $cfg "batch_templates"
      if ($null -ne $tplRaw) {
        foreach ($t in $tplRaw) {
          if (-not [string]::IsNullOrWhiteSpace([string]$t)) { $templates += [string]$t }
        }
      }
      $selectedRaw = Try-GetProp $cfg "batch_template_selected"
      $defaultRaw = Try-GetProp $cfg "batch_template_default"
      $legacyTpl = Try-GetProp $cfg "batch_template"
      if ($null -ne $legacyTpl) { $selected = [string]$legacyTpl }
      if ($null -ne $selectedRaw) { $selected = [string]$selectedRaw }
      if ($null -ne $defaultRaw) {
        $default = [string]$defaultRaw
      } elseif ($null -ne $selectedRaw) {
        $default = [string]$selectedRaw
      } elseif ($null -ne $legacyTpl) {
        $default = [string]$legacyTpl
      }
    }
    if ([string]::IsNullOrWhiteSpace($selected)) {
      $selected = "常规化学{year}{value}"
    }
    if ([string]::IsNullOrWhiteSpace($default)) {
      $default = "常规化学{year}{value}"
    }
    if (-not ($templates -contains "常规化学{year}{value}")) {
      $templates = @("常规化学{year}{value}") + $templates
    }
    if (-not ($templates -contains $selected)) {
      $templates += $selected
    }
    if (-not ($templates -contains $default)) {
      $templates += $default
    }
    $inst = Unique-List $inst
    $bats = Unique-List $bats
    $excludes = Unique-List $excludes
    $templates = Unique-List $templates
    $out = [pscustomobject]@{
      instruments = $inst
      batches = $bats
      batch_templates = $templates
      batch_template_selected = $selected
      batch_template_default = $default
      exclude_keywords = $excludes
    }
    $json = $out | ConvertTo-Json -Depth 4
    $json | Set-Content -LiteralPath $ConfigPath -Encoding UTF8
  } catch {
    Show-Error ("保存配置失败: {0}" -f $_.Exception.Message)
  }
}

function Normalize-List($values) {
  $items = @()
  if ($null -eq $values) { return $items }

  if ($values -is [string]) {
    if (-not [string]::IsNullOrWhiteSpace($values)) {
      return @([string]$values)
    }
    return @()
  }

  if ($values -is [System.Collections.IEnumerable]) {
    foreach ($v in $values) {
      if (-not [string]::IsNullOrWhiteSpace([string]$v)) {
        $items += [string]$v
      }
    }
    return $items
  }

  $s = [string]$values
  if (-not [string]::IsNullOrWhiteSpace($s)) {
    return @([string]$s)
  }
  return @()
}

# 列表去重并保持原始顺序。
function Unique-List($items) {
  $seen = @{}
  $out = @()
  foreach ($v in $items) {
    $s = [string]$v
    if ([string]::IsNullOrWhiteSpace($s)) { continue }
    if (-not $seen.ContainsKey($s)) {
      $seen[$s] = $true
      $out += $s
    }
  }
  return $out
}

# 将配置列表填充到下拉控件。
function Add-ComboItems($combo, $list) {
  if ($null -eq $combo) { return }
  if ($null -eq $list) { return }
  foreach ($item in $list) {
    $text = [string]$item
    if (-not [string]::IsNullOrWhiteSpace($text)) {
      $combo.Items.Add($text) | Out-Null
    }
  }
}

function Add-ComboItem($combo, [string]$value) {
  if ($null -eq $combo) { return $false }
  if ([string]::IsNullOrWhiteSpace($value)) { return $false }
  if (-not $combo.Items.Contains($value)) {
    $combo.Items.Add($value) | Out-Null
    return $true
  }
  return $false
}

# 从下拉控件读取当前可选项（用于持久化）。
function Get-ComboItems($combo) {
  $items = @()
  if ($null -eq $combo) { return $items }
  foreach ($item in $combo.Items) {
    $text = [string]$item
    if (-not [string]::IsNullOrWhiteSpace($text)) {
      $items += $text
    }
  }
  return $items
}

# 将当前 UI 状态写回配置文件。
function Save-ConfigFromUI($comboInstr, $comboNormal, $comboTemplate, [string]$defaultTemplate, $excludeList) {
  $inst = Get-ComboItems $comboInstr
  $bats = Get-ComboItems $comboNormal
  $templates = Get-ComboItems $comboTemplate
  $selected = Safe-Trim $comboTemplate.Text
  if ([string]::IsNullOrWhiteSpace($defaultTemplate)) {
    $defaultTemplate = "常规化学{year}{value}"
  }
  Save-Config ([pscustomobject]@{
    instruments = $inst
    batches = $bats
    batch_templates = $templates
    batch_template_selected = $selected
    batch_template_default = $defaultTemplate
    exclude_keywords = $excludeList
  })
}

# 批号操作：同时写入“正常/异常”两个下拉。
function Add-BatchItem($comboNormal, $comboAbnormal, [string]$value) {
  if (-not $comboNormal.Items.Contains($value)) { $comboNormal.Items.Add($value) | Out-Null }
  if (-not $comboAbnormal.Items.Contains($value)) { $comboAbnormal.Items.Add($value) | Out-Null }
}

# 批号操作：同时删除“正常/异常”两个下拉。
function Remove-BatchItem($comboNormal, $comboAbnormal, [string]$value) {
  if ($comboNormal.Items.Contains($value)) { $comboNormal.Items.Remove($value) | Out-Null }
  if ($comboAbnormal.Items.Contains($value)) { $comboAbnormal.Items.Remove($value) | Out-Null }
}

# 清洗非法文件名字符（用于输出文件名）。
function Sanitize-FileName([string]$name) {
  if ($null -eq $name) { return "" }
  $invalid = [System.IO.Path]::GetInvalidFileNameChars()
  $sb = New-Object System.Text.StringBuilder
  foreach ($ch in $name.ToCharArray()) {
    if ($invalid -contains $ch) {
      [void]$sb.Append("_")
    } else {
      [void]$sb.Append($ch)
    }
  }
  return $sb.ToString()
}

# 从文件名提取日期范围令牌（例如 20260101-20260131）。
function Get-DateRangeToken([System.Collections.IEnumerable]$files) {
  foreach ($f in $files) {
    if ($f.Name -match "(\d{8})-(\d{8})") {
      $startRaw = $Matches[1]
      $endRaw = $Matches[2]
      if ($endRaw.Length -ge 4) {
        $endToken = $endRaw.Substring(4)
        return "$startRaw-$endToken"
      }
      return "$startRaw-$endRaw"
    }
  }
  return ""
}

function Add-Unique([System.Collections.Generic.List[string]]$list, [string]$value) {
  if ($null -eq $list) { return $false }
  if ([string]::IsNullOrWhiteSpace($value)) { return $false }
  if (-not ($list -contains $value)) {
    $list.Add($value) | Out-Null
    return $true
  }
  return $false
}

function Remove-Value([System.Collections.Generic.List[string]]$list, [string]$value) {
  if ($null -eq $list) { return $false }
  if ([string]::IsNullOrWhiteSpace($value)) { return $false }
  if ($list -contains $value) {
    $list.Remove($value)
    return $true
  }
  return $false
}

function Get-ExcludeKeywords {
  if (-not (Test-Path -LiteralPath $ExcludePath)) { return @() }
  $raw = Get-Content -LiteralPath $ExcludePath -Raw
  return $raw.Split(",") | ForEach-Object { $_.Trim() } | Where-Object { $_ }
}

function Get-ExcludeKeywordsFromConfig($cfg) {
  $raw = Try-GetProp $cfg "exclude_keywords"
  if ($null -eq $raw) { return @() }
  return @(Normalize-List $raw)
}

function Get-ProjectNameFromFile([string]$fileName) {
  if ($fileName -match "([A-Za-z0-9-]+)\.xls$") {
    return $Matches[1]
  }
  return [System.IO.Path]::GetFileNameWithoutExtension($fileName)
}

function Format-DateRange([string]$fileName) {
  if ($fileName -match "(\d{8})-(\d{8})") {
    $startRaw = $Matches[1]
    $endRaw = $Matches[2]
    $start = [datetime]::ParseExact($startRaw, "yyyyMMdd", $null)
    $end = [datetime]::ParseExact($endRaw, "yyyyMMdd", $null)
    return ("{0:yyyy-MM-dd}~{1:yyyy-MM-dd}" -f $start, $end)
  }
  return ""
}

function Is-DateLike($value) {
  if ($null -eq $value) { return $false }
  if ($value -is [double]) {
    return ($value -ge 20000 -and $value -le 60000)
  }
  if ($value -is [datetime]) { return $true }
  if ($value -is [string]) {
    if ($value -match "^\d{4}[-/]\d{1,2}[-/]\d{1,2}$") { return $true }
    if ($value -match "^\d{1,2}[-/]\d{1,2}[-/]\d{2,4}$") { return $true }
    if ($value -match "^\d{8}$") { return $true }
  }
  return $false
}

# 解析日期单元格（支持 Excel OADate 与文本日期）。
function Parse-DateValue($value2, [string]$text) {
  try {
    if ($value2 -is [double]) {
      if ($value2 -ge 20000 -and $value2 -le 60000) {
        return [datetime]::FromOADate($value2)
      }
    } elseif ($value2 -is [datetime]) {
      return $value2
    }
    if (-not [string]::IsNullOrWhiteSpace($text)) {
      $dt = [datetime]::MinValue
      if ([datetime]::TryParse($text, [ref]$dt)) { return $dt }
    }
  } catch {
  }
  return $null
}

# 从文件名解析日期范围。
function Get-DateRangeFromFileName([string]$fileName) {
  if ($fileName -match "(\d{8})-(\d{8})") {
    $start = [datetime]::ParseExact($Matches[1], "yyyyMMdd", $null)
    $end = [datetime]::ParseExact($Matches[2], "yyyyMMdd", $null)
    return [pscustomobject]@{ Start = $start; End = $end }
  }
  return $null
}

# 安全创建日期（非法年月日返回 null）。
function Try-CreateDate($year, $month, $day) {
  try {
    return [datetime]::new($year, $month, $day)
  } catch {
    return $null
  }
}

# 规范化日期键：统一输出 yyyy-MM-dd，兼容仅日数字段。
function Normalize-DateKey([string]$key, [string]$text, $range) {
  $raw = ""
  if (-not [string]::IsNullOrWhiteSpace($key)) {
    $raw = $key.Trim()
  } elseif (-not [string]::IsNullOrWhiteSpace($text)) {
    $raw = $text.Trim()
  }
  if ([string]::IsNullOrWhiteSpace($raw)) {
    return [pscustomobject]@{ Key = ""; Text = "" }
  }
  if ($null -eq $range) {
    return [pscustomobject]@{ Key = $raw; Text = $raw }
  }

  $dt = $null
  if ($raw -match "^\d{8}$") {
    try { $dt = [datetime]::ParseExact($raw, "yyyyMMdd", $null) } catch {}
  } elseif ($raw -match "^\d{4}[-/]\d{1,2}[-/]\d{1,2}$") {
    try { $dt = [datetime]::Parse($raw) } catch {}
  }
  if ($null -ne $dt) {
    return [pscustomobject]@{ Key = $dt.ToString("yyyy-MM-dd"); Text = $dt.ToString("yyyy/M/d") }
  }

  $dayMatch = [regex]::Match($raw, "\d{1,2}")
  if ($dayMatch.Success) {
    $day = [int]$dayMatch.Value
    if ($day -ge 1 -and $day -le 31) {
      $cand1 = Try-CreateDate $range.Start.Year $range.Start.Month $day
      $cand2 = Try-CreateDate $range.End.Year $range.End.Month $day
      if ($cand1 -ne $null -and $cand1 -ge $range.Start -and $cand1 -le $range.End) {
        return [pscustomobject]@{ Key = $cand1.ToString("yyyy-MM-dd"); Text = $cand1.ToString("yyyy/M/d") }
      }
      if ($cand2 -ne $null -and $cand2 -ge $range.Start -and $cand2 -le $range.End) {
        return [pscustomobject]@{ Key = $cand2.ToString("yyyy-MM-dd"); Text = $cand2.ToString("yyyy/M/d") }
      }
      if ($cand1 -ne $null) {
        return [pscustomobject]@{ Key = $cand1.ToString("yyyy-MM-dd"); Text = $cand1.ToString("yyyy/M/d") }
      }
    }
  }
  return [pscustomobject]@{ Key = $raw; Text = $raw }
}

# 构建期望日期序列（优先使用文件名范围）。
function Build-ExpectedDates($range, $fallbackTexts, $fallbackKeys) {
  $keys = @()
  $texts = @()
  if ($null -ne $range) {
    $d = $range.Start
    while ($d -le $range.End) {
      $keys += $d.ToString("yyyy-MM-dd")
      $texts += $d.ToString("yyyy/M/d")
      $d = $d.AddDays(1)
    }
    return [pscustomobject]@{ Keys = $keys; Texts = $texts }
  }
  $keys = @(Normalize-List $fallbackKeys)
  $texts = @(Normalize-List $fallbackTexts)
  if ($texts.Count -lt $keys.Count) {
    $texts = $keys
  }
  return [pscustomobject]@{ Keys = $keys; Texts = $texts }
}

# 提取数值（非数值返回 null）。
function Get-NumericValue($value) {
  if ($null -eq $value) { return $null }
  if ($value -is [double] -or $value -is [float] -or $value -is [int] -or $value -is [decimal]) {
    return [double]$value
  }
  $s = [string]$value
  if ([string]::IsNullOrWhiteSpace($s)) { return $null }
  $num = 0.0
  if ([double]::TryParse($s, [ref]$num)) { return $num }
  return $null
}

# 计算均值（仅统计可解析数值）。
function Get-AverageValue($map, $expectedKeys) {
  $vals = @()
  foreach ($k in $expectedKeys) {
    if ($map.ContainsKey($k)) {
      $n = Get-NumericValue $map[$k]
      if ($null -ne $n) { $vals += $n }
    }
  }
  if ($vals.Count -eq 0) { return $null }
  $avg = ($vals | Measure-Object -Average).Average
  return $avg
}

# 找到最接近均值的原始值（用于缺失填充）。
function Get-NearestValueToAverage($map, $expectedKeys) {
  $vals = @()
  foreach ($k in $expectedKeys) {
    if ($map.ContainsKey($k)) {
      $raw = $map[$k]
      $n = Get-NumericValue $raw
      if ($null -ne $n) {
        $vals += [pscustomobject]@{ Num = $n; Raw = $raw }
      }
    }
  }
  if ($vals.Count -eq 0) { return $null }
  $avg = ($vals.Num | Measure-Object -Average).Average
  $best = $null
  $bestDiff = [double]::PositiveInfinity
  foreach ($v in $vals) {
    $diff = [math]::Abs($v.Num - $avg)
    if ($diff -lt $bestDiff) {
      $bestDiff = $diff
      $best = $v.Raw
    }
  }
  return $best
}

function Get-StartRow($ws) {
  $row1 = $ws.Cells.Item(1,1).Value2
  $row2 = $ws.Cells.Item(2,1).Value2
  if (Is-DateLike $row1) { return 1 }
  if (Is-DateLike $row2) { return 2 }
  return 1
}

# 读取单个 Excel 文件的数据区（日期、正常、异常）。
function Read-DataFromFile($excel, [string]$path) {
  $wb = $excel.Workbooks.Open($path, 0, $true)
  $ws = $null
  $used = $null
  try {
    $ws = $wb.Worksheets.Item(1)
    $used = $ws.UsedRange
    $firstRow = $used.Row
    $lastRow = $used.Row + $used.Rows.Count - 1
    $startRow = Get-StartRow $ws
    if ($startRow -lt $firstRow) { $startRow = $firstRow }

    $lastDataRow = $lastRow
    while ($lastDataRow -ge $startRow) {
      $text = [string]$ws.Cells.Item($lastDataRow,1).Text
      if (-not [string]::IsNullOrWhiteSpace($text)) { break }
      $lastDataRow--
    }
    if ($lastDataRow -lt $startRow) {
      return [pscustomobject]@{ Normal = @(); Abnormal = @(); DataLen = 0 }
    }

    $datesText = New-Object System.Collections.Generic.List[string]
    $datesKey = New-Object System.Collections.Generic.List[string]
    $normal = New-Object System.Collections.Generic.List[object]
    $abnormal = New-Object System.Collections.Generic.List[object]
    for ($r = $startRow; $r -le $lastDataRow; $r++) {
      $dateValue = $ws.Cells.Item($r,1).Value2
      $dateText = [string]$ws.Cells.Item($r,1).Text
      if ($null -eq $dateText) { $dateText = "" }
      $dt = Parse-DateValue $dateValue $dateText
      if ($null -ne $dt) {
        $datesKey.Add($dt.ToString("yyyy-MM-dd")) | Out-Null
        $datesText.Add($dt.ToString("yyyy/M/d")) | Out-Null
      } else {
        $datesKey.Add($dateText.Trim()) | Out-Null
        $datesText.Add($dateText.Trim()) | Out-Null
      }
      $normal.Add($ws.Cells.Item($r,2).Value2) | Out-Null
      $abnormal.Add($ws.Cells.Item($r,3).Value2) | Out-Null
    }
    return [pscustomobject]@{
      DatesKey = $datesKey
      DatesText = $datesText
      Normal = $normal
      Abnormal = $abnormal
      DataLen = ($lastDataRow - $startRow + 1)
    }
  } finally {
    if ($used -ne $null) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($used) | Out-Null }
    if ($ws -ne $null) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ws) | Out-Null }
    $wb.Close($false)
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null
  }
}

# 构建输出行：按日期对齐，先正常后异常。
# missingMode: blank | avg
function Build-OutputRows($files, [string]$instrument, [string]$batchNormal, [string]$batchAbnormal, [string]$missingMode) {
  $rows = New-Object System.Collections.Generic.List[object]
  $warnings = New-Object System.Collections.Generic.List[string]
  $excel = $null
  try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.ScreenUpdating = $false

    foreach ($file in $files) {
      $project = Get-ProjectNameFromFile $file.Name
      $data = Read-DataFromFile $excel $file.FullName
      if ($data.DataLen -le 0) { continue }
      $range = Get-DateRangeFromFileName $file.Name
      $expected = Build-ExpectedDates $range $data.DatesText $data.DatesKey
      $expectedKeys = @($expected.Keys)
      $expectedTexts = @($expected.Texts)

      $normalMap = @{}
      $abnormalMap = @{}
      $seenKeys = @{}
      for ($i = 0; $i -lt $data.DataLen; $i++) {
        $norm = Normalize-DateKey ([string]$data.DatesKey[$i]) ([string]$data.DatesText[$i]) $range
        $key = [string]$norm.Key
        if ([string]::IsNullOrWhiteSpace($key)) { continue }
        $normalMap[$key] = $data.Normal[$i]
        $abnormalMap[$key] = $data.Abnormal[$i]
        $seenKeys[$key] = $true
      }

      if ($null -ne $range) {
        $missingKeys = @()
        foreach ($k in $expectedKeys) {
          if (-not $normalMap.ContainsKey($k) -and -not $abnormalMap.ContainsKey($k)) {
            $missingKeys += $k
          }
        }
        $extraKeys = @()
        foreach ($k2 in $seenKeys.Keys) {
          if (-not ($expectedKeys -contains $k2)) { $extraKeys += $k2 }
        }
        if (($expectedKeys.Count -ne $data.DataLen) -or ($missingKeys.Count -gt 0) -or ($extraKeys.Count -gt 0)) {
          $missingText = ($missingKeys | ForEach-Object {
            try { ([datetime]::ParseExact($_, "yyyy-MM-dd", $null)).ToString("yyyy/M/d") } catch { $_ }
          }) -join ", "
          $extraText = ($extraKeys | ForEach-Object { $_ }) -join ", "
          $warnings.Add(("文件: {0} 期望{1}条, 实际{2}条; 缺失: {3}; 超出范围: {4}" -f $file.Name, $expectedKeys.Count, $data.DataLen, $missingText, $extraText)) | Out-Null
        }
      }

      $fillNormal = $null
      $fillAbnormal = $null
      if ($missingMode -eq "avg") {
        $fillNormal = Get-NearestValueToAverage $normalMap $expectedKeys
        $fillAbnormal = Get-NearestValueToAverage $abnormalMap $expectedKeys
      }

      for ($i = 0; $i -lt $expectedKeys.Count; $i++) {
        $key = $expectedKeys[$i]
        $dateText = $expectedTexts[$i]
        $nval = $null
        if ($normalMap.ContainsKey($key)) {
          $nval = $normalMap[$key]
        } elseif ($missingMode -eq "avg" -and $null -ne $fillNormal) {
          $nval = $fillNormal
        } else {
          $nval = ""
        }
        $rows.Add([pscustomobject]@{
          A = $instrument
          B = $project
          C = $batchNormal
          D = $nval
          E = $dateText
        }) | Out-Null
      }

      for ($i = 0; $i -lt $expectedKeys.Count; $i++) {
        $key = $expectedKeys[$i]
        $dateText = $expectedTexts[$i]
        $aval = $null
        if ($abnormalMap.ContainsKey($key)) {
          $aval = $abnormalMap[$key]
        } elseif ($missingMode -eq "avg" -and $null -ne $fillAbnormal) {
          $aval = $fillAbnormal
        } else {
          $aval = ""
        }
        $rows.Add([pscustomobject]@{
          A = $instrument
          B = $project
          C = $batchAbnormal
          D = $aval
          E = $dateText
        }) | Out-Null
      }
    }
  } finally {
    if ($excel -ne $null) {
      $excel.Quit()
      [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
  }
  return [pscustomobject]@{ Rows = $rows; Warnings = $warnings }
}

# 将数据写入模板并保存输出文件。
function Write-Template([System.Collections.IList]$rows, [string]$outputName) {
  if (-not (Test-Path -LiteralPath $TemplatePath)) {
    throw "找不到 templet.xls"
  }
  if (-not (Test-Path -LiteralPath $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir | Out-Null
  }

  $excel = $null
  $wb = $null
  $ws = $null
  $used = $null
  try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.ScreenUpdating = $false

    $wb = $excel.Workbooks.Open($TemplatePath, 0, $false)
    $ws = $wb.Worksheets.Item(1)

    $used = $ws.UsedRange
    $lastRow = $used.Row + $used.Rows.Count - 1
    if ($lastRow -ge 2) {
      $ws.Range("A2", "E$lastRow").ClearContents() | Out-Null
    }

    $rowCount = $rows.Count
    if ($rowCount -gt 0) {
      $data = New-Object 'object[,]' $rowCount, 5
      for ($i = 0; $i -lt $rowCount; $i++) {
        $data[$i,0] = $rows[$i].A
        $data[$i,1] = $rows[$i].B
        $data[$i,2] = $rows[$i].C
        $data[$i,3] = $rows[$i].D
        $data[$i,4] = $rows[$i].E
      }
      $startRow = 2
      $endRow = $startRow + $rowCount - 1
      $ws.Range("A$startRow", "E$endRow").Value2 = $data
    }

    if ([string]::IsNullOrWhiteSpace($outputName)) {
      $stamp = Get-Date -Format "yyyyMMdd_HHmmss"
      $outputName = "templet_filled_{0}.xls" -f $stamp
    }
    if (-not $outputName.ToLower().EndsWith(".xls")) {
      $outputName = "$outputName.xls"
    }
    $outPath = Join-Path $OutputDir $outputName
    $wb.SaveAs($outPath, 56)
    return $outPath
  } finally {
    if ($wb -ne $null) { $wb.Close($false) }
    if ($used -ne $null) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($used) | Out-Null }
    if ($ws -ne $null) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ws) | Out-Null }
    if ($wb -ne $null) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null }
    if ($excel -ne $null) {
      $excel.Quit()
      [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
  }
}

# 批号模板处理：支持 {year}/{value} 占位符。
function Build-BatchValue([string]$raw, [bool]$useTemplate, [string]$templateText) {
  if ($null -eq $raw) { return "" }
  $raw = $raw.Trim()
  if ([string]::IsNullOrWhiteSpace($raw)) { return "" }
  if ($useTemplate) {
    $tpl = $templateText
    if ([string]::IsNullOrWhiteSpace($tpl)) {
      $tpl = "常规化学{year}{value}"
    }
    $year = (Get-Date).Year.ToString()
    $tpl = $tpl.Replace("{year}", $year)
    if ($tpl -match "\{value\}") {
      $tpl = $tpl.Replace("{value}", $raw)
      return $tpl
    }
    return ($tpl + $raw)
  }
  return $raw
}

# Safe trim helper: always return a non-null string.
function Safe-Trim([string]$value) {
  if ($null -eq $value) { return "" }
  return $value.Trim()
}

# UI message wrappers (centralized so later style/localization updates are easy).
function Show-Error([string]$message) {
  [System.Windows.Forms.MessageBox]::Show($message, "错误", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
}

function Show-Info([string]$message) {
  [System.Windows.Forms.MessageBox]::Show($message, "完成", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
}

function Show-Warn([string]$message) {
  [System.Windows.Forms.MessageBox]::Show($message, "提示", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
}

function Write-ErrorLog([string]$title, [string]$detail) {
  try {
    $logPath = Join-Path $Root "qc_tool_error.log"
    $time = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $block = @(
      "[{0}] {1}" -f $time, $title
      $detail
      ""
    ) -join "`r`n"
    Add-Content -LiteralPath $logPath -Value $block -Encoding UTF8
  } catch {
    # Keep logging best-effort only, never crash on log write failures.
  }
}

function Show-Fatal([string]$message) {
  try {
    [System.Windows.Forms.MessageBox]::Show($message, "错误", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
  } catch {
    # If GUI is unavailable, do nothing; log is still written separately.
  }
}

function Register-GlobalExceptionHandlers {
  [System.Windows.Forms.Application]::SetUnhandledExceptionMode([System.Windows.Forms.UnhandledExceptionMode]::CatchException)

  $script:HandlingGlobalException = $false
  $onThreadException = {
    param($sender, $eventArgs)
    if ($script:HandlingGlobalException) { return }
    $script:HandlingGlobalException = $true
    try {
      $ex = $null
      if ($null -ne $eventArgs) {
        $prop = $eventArgs.PSObject.Properties["Exception"]
        if ($null -ne $prop) { $ex = $prop.Value }
      }
      $detail = if ($null -ne $ex) { $ex.ToString() } else { ("未知线程异常: {0}" -f ($eventArgs | Out-String)) }
      Write-ErrorLog "未处理异常(ThreadException)" $detail
      Show-Fatal "程序发生未处理异常，详情已写入 qc_tool_error.log。"
    } finally {
      $script:HandlingGlobalException = $false
    }
  }.GetNewClosure()
  [System.Windows.Forms.Application]::add_ThreadException($onThreadException)

  $onUnhandledException = {
    param($sender, $eventArgs)
    if ($script:HandlingGlobalException) { return }
    $script:HandlingGlobalException = $true
    try {
      $obj = $null
      if ($null -ne $eventArgs) {
        $prop = $eventArgs.PSObject.Properties["ExceptionObject"]
        if ($null -ne $prop) { $obj = $prop.Value }
      }
      $detail = if ($null -ne $obj) { $obj.ToString() } else { ("未知域级异常: {0}" -f ($eventArgs | Out-String)) }
      Write-ErrorLog "未处理异常(UnhandledException)" $detail
      Show-Fatal "程序发生未处理异常，详情已写入 qc_tool_error.log。"
    } finally {
      $script:HandlingGlobalException = $false
    }
  }.GetNewClosure()
  [System.AppDomain]::CurrentDomain.add_UnhandledException($onUnhandledException)
}

# Parse CSS-like hex colors (e.g. #1A73E8) into WinForms Color objects.
function New-Color([string]$hex) {
  return [System.Drawing.ColorTranslator]::FromHtml($hex)
}

function Resolve-Color($value, [System.Drawing.Color]$fallback) {
  if ($value -is [System.Drawing.Color]) { return $value }
  if ($null -eq $value) { return $fallback }

  $raw = [string]$value
  if ([string]::IsNullOrWhiteSpace($raw)) { return $fallback }
  $raw = $raw.Trim()

  if ($raw -match '^\[System\.Drawing\.Color\]::(?<name>[A-Za-z]+)$') {
    $byLegacyName = [System.Drawing.Color]::FromName($Matches['name'])
    if ($byLegacyName.A -ne 0 -or $Matches['name'] -eq 'Transparent') { return $byLegacyName }
  }

  try {
    $byHtml = [System.Drawing.ColorTranslator]::FromHtml($raw)
    if ($byHtml.A -ne 0 -or $raw -eq 'Transparent') { return $byHtml }
  } catch {}

  $byName = [System.Drawing.Color]::FromName($raw)
  if ($byName.A -ne 0 -or $raw -eq 'Transparent') { return $byName }
  return $fallback
}

# Shift RGB channels by delta to build hover/pressed variants from one base color.
function Shift-Color([System.Drawing.Color]$base, [int]$delta) {
  $r = [Math]::Min(255, [Math]::Max(0, $base.R + $delta))
  $g = [Math]::Min(255, [Math]::Max(0, $base.G + $delta))
  $b = [Math]::Min(255, [Math]::Max(0, $base.B + $delta))
  return [System.Drawing.Color]::FromArgb($r, $g, $b)
}

# Create a rounded-rectangle drawing path used for card/button clipping.
function New-RoundedPath([System.Drawing.Rectangle]$rect, [int]$radius) {
  $path = New-Object System.Drawing.Drawing2D.GraphicsPath
  $diameter = [Math]::Max(2, $radius * 2)
  $arc = New-Object System.Drawing.Rectangle($rect.X, $rect.Y, $diameter, $diameter)
  $path.AddArc($arc, 180, 90)
  $arc.X = $rect.Right - $diameter
  $path.AddArc($arc, 270, 90)
  $arc.Y = $rect.Bottom - $diameter
  $path.AddArc($arc, 0, 90)
  $arc.X = $rect.Left
  $path.AddArc($arc, 90, 90)
  $path.CloseFigure()
  return $path
}

# Apply rounded corners by setting the control Region and keep it synced on resize.
function Set-RoundedRegion($control, [int]$radius) {
  if ($null -eq $control) { return }
  $applyRegion = {
    if ($this.Width -le 1 -or $this.Height -le 1) { return }
    $rect = New-Object System.Drawing.Rectangle(0, 0, $this.Width, $this.Height)
    $path = New-RoundedPath $rect $radius
    if ($this.Region) { $this.Region.Dispose() }
    $this.Region = New-Object System.Drawing.Region($path)
    $path.Dispose()
  }.GetNewClosure()
  $control.Add_SizeChanged($applyRegion)
  if ($control.Width -gt 1 -and $control.Height -gt 1) {
    $rect = New-Object System.Drawing.Rectangle(0, 0, $control.Width, $control.Height)
    $path = New-RoundedPath $rect $radius
    if ($control.Region) { $control.Region.Dispose() }
    $control.Region = New-Object System.Drawing.Region($path)
    $path.Dispose()
  }
}

# Reduce repaint flicker for custom-drawn controls.
function Enable-DoubleBuffer($control) {
  if ($null -eq $control) { return }
  $flags = [System.Reflection.BindingFlags]::Instance -bor [System.Reflection.BindingFlags]::NonPublic
  $prop = $control.GetType().GetProperty('DoubleBuffered', $flags)
  if ($null -ne $prop) {
    $prop.SetValue($control, $true, $null)
  }
}

# Give text inputs a subtle focus accent to improve pointer/keyboard orientation.
function Add-FocusAccent($control, [System.Drawing.Color]$focusColor) {
  if ($null -eq $control) { return }
  $baseColor = [System.Drawing.Color]::White
  $control.BackColor = $baseColor
  $onEnter = { $this.BackColor = $focusColor }.GetNewClosure()
  $onLeave = { $this.BackColor = $baseColor }.GetNewClosure()
  $control.Add_Enter($onEnter)
  $control.Add_Leave($onLeave)
}

# Apple-like micro interactions:
# - hover: brighten + lift 1px
# - press: darken + sink 1px
# - release: return to hover state
function Add-InteractiveButtonStyle($button, $baseColor, $textColor) {
  if ($null -eq $button) { return }
  $baseColor = Resolve-Color $baseColor (New-Color "#1A73E8")
  $textColor = Resolve-Color $textColor ([System.Drawing.Color]::White)
  $hoverColor = Shift-Color $baseColor 16
  $pressedColor = Shift-Color $baseColor -20
  $borderColor = Shift-Color $baseColor -35
  # Cache scalar coordinates to avoid runtime arithmetic on array-typed values in event closures.
  $baseX = [int]$button.Left
  $baseY = [int]$button.Top

  $button.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
  $button.FlatAppearance.BorderSize = 1
  $button.FlatAppearance.BorderColor = $borderColor
  $button.BackColor = $baseColor
  $button.ForeColor = $textColor
  $button.Cursor = [System.Windows.Forms.Cursors]::Hand

  $onMouseEnter = {
    $this.BackColor = $hoverColor
    $this.Left = $baseX
    $this.Top = $baseY - 1
  }.GetNewClosure()
  $onMouseLeave = {
    $this.BackColor = $baseColor
    $this.Left = $baseX
    $this.Top = $baseY
  }.GetNewClosure()
  $onMouseDown = {
    $this.BackColor = $pressedColor
    $this.Left = $baseX
    $this.Top = $baseY + 1
  }.GetNewClosure()
  $onMouseUp = {
    $this.BackColor = $hoverColor
    $this.Left = $baseX
    $this.Top = $baseY - 1
  }.GetNewClosure()
  $button.Add_MouseEnter($onMouseEnter)
  $button.Add_MouseLeave($onMouseLeave)
  $button.Add_MouseDown($onMouseDown)
  $button.Add_MouseUp($onMouseUp)
}

# Smooth window fade-in on first show to avoid abrupt visual pop.
function Start-FadeIn($form, [double]$step = 0.08, [int]$interval = 14) {
  if ($null -eq $form) { return }
  $form.Opacity = 0
  $timer = New-Object System.Windows.Forms.Timer
  $timer.Interval = $interval
  $onTick = {
    if ($form.IsDisposed) {
      $timer.Stop()
      $timer.Dispose()
      return
    }
    $next = $form.Opacity + $step
    if ($next -ge 1) {
      $form.Opacity = 1
      $timer.Stop()
      $timer.Dispose()
    } else {
      $form.Opacity = $next
    }
  }.GetNewClosure()
  $onShown = { $timer.Start() }.GetNewClosure()
  $timer.Add_Tick($onTick)
  $form.Add_Shown($onShown)
}

# ------------------------------- UI Composition -------------------------------
# 1) Load persisted history/config.
# 2) Build modernized WinForms UI (card layout + vivid Material palette).
# 3) Bind event handlers for save/delete/run actions.
# 4) Keep original data-processing logic unchanged.
try {
  Register-GlobalExceptionHandlers
  $cfg = Read-Config
  if ($null -eq $cfg) {
    $cfg = [pscustomobject]@{ instruments = @(); batches = @(); batch_template = "常规化学{year}{value}" }
  }
  $instList = @(Normalize-List $cfg.instruments)
  $batchList = @(Normalize-List $cfg.batches)
  $excludeList = @(Get-ExcludeKeywordsFromConfig $cfg)
  if ($excludeList.Count -eq 0) {
    $excludeList = @(Get-ExcludeKeywords)
  }
  $rawTemplates = $null
  if ($cfg.PSObject -and $cfg.PSObject.Properties.Match("batch_templates").Count -gt 0) {
    $rawTemplates = $cfg.batch_templates
  } elseif ($cfg.PSObject -and $cfg.PSObject.Properties.Match("batch_template").Count -gt 0) {
    $rawTemplates = @([string]$cfg.batch_template)
  }
  $templateList = @(Normalize-List $rawTemplates)
  $templateSelected = ""
  $templateDefault = ""
  if ($cfg.PSObject -and $cfg.PSObject.Properties.Match("batch_template_default").Count -gt 0) {
    $templateDefault = [string]$cfg.batch_template_default
  } elseif ($cfg.PSObject -and $cfg.PSObject.Properties.Match("batch_template_selected").Count -gt 0) {
    $templateDefault = [string]$cfg.batch_template_selected
  } elseif ($cfg.PSObject -and $cfg.PSObject.Properties.Match("batch_template").Count -gt 0) {
    $templateDefault = [string]$cfg.batch_template
  }
  if ($cfg.PSObject -and $cfg.PSObject.Properties.Match("batch_template_selected").Count -gt 0) {
    $templateSelected = [string]$cfg.batch_template_selected
  } elseif (-not [string]::IsNullOrWhiteSpace($templateDefault)) {
    $templateSelected = $templateDefault
  } elseif ($cfg.PSObject -and $cfg.PSObject.Properties.Match("batch_template").Count -gt 0) {
    $templateSelected = [string]$cfg.batch_template
  }
  if ([string]::IsNullOrWhiteSpace($templateSelected)) {
    $templateSelected = "常规化学{year}{value}"
  }
  if ([string]::IsNullOrWhiteSpace($templateDefault)) {
    $templateDefault = "常规化学{year}{value}"
  }
  if ($templateList.Count -eq 0) {
    $templateList = @("常规化学{year}{value}")
  }
  if (-not ($templateList -contains $templateDefault)) {
    $templateList += $templateDefault
  }
  if (-not ($templateList -contains $templateSelected)) {
    $templateList += $templateSelected
  }
  $script:TemplateDefault = $templateDefault

  # Main window shell.
  $form = New-Object System.Windows.Forms.Form
  $form.Text = "QC Excel 处理  |  by YRZ"
  $form.Size = New-Object System.Drawing.Size(780, 430)
  $form.StartPosition = "CenterScreen"
  $form.FormBorderStyle = "FixedDialog"
  $form.MaximizeBox = $false
  $form.MinimizeBox = $false
  $form.BackColor = New-Color "#E8F0FE"
  Set-WindowIconFromPng $form $IconPath
  Enable-DoubleBuffer $form

  # Typography system: keep one coherent scale for readability.
  $uiFont = New-Object System.Drawing.Font("Microsoft YaHei UI", 9.5, [System.Drawing.FontStyle]::Regular)
  $labelFont = New-Object System.Drawing.Font("Microsoft YaHei UI", 9.5, [System.Drawing.FontStyle]::Bold)
  $buttonFont = New-Object System.Drawing.Font("Microsoft YaHei UI", 9.0, [System.Drawing.FontStyle]::Bold)
  $authorFont = New-Object System.Drawing.Font("Segoe UI", 9.0, [System.Drawing.FontStyle]::Italic)
  $form.Font = $uiFont

  # Palette: Google-like vivid accents with restrained neutral text tones.
  $accentBlue = New-Color "#1A73E8"
  $accentGreen = New-Color "#34A853"
  $accentYellow = New-Color "#F9AB00"
  $accentCoral = New-Color "#FF7043"
  $textPrimary = New-Color "#1F2937"
  $textSecondary = New-Color "#4B5563"

  # Keep static background color to avoid unstable custom Paint callbacks in PS2EXE runtime.

  # Lightweight author signature in the title area.
  $labelAuthor = New-Object System.Windows.Forms.Label
  $labelAuthor.Text = "by YRZ"
  $labelAuthor.Location = New-Object System.Drawing.Point(680, 10)
  $labelAuthor.AutoSize = $true
  $labelAuthor.Font = $authorFont
  $labelAuthor.ForeColor = New-Color "#6B7280"
  $labelAuthor.BackColor = [System.Drawing.Color]::Transparent

  # Glass-like content card to separate interactive controls from background art.
  $cardPanel = New-Object System.Windows.Forms.Panel
  $cardPanel.Location = New-Object System.Drawing.Point(18, 34)
  $cardPanel.Size = New-Object System.Drawing.Size(732, 344)
  $cardPanel.BackColor = [System.Drawing.Color]::FromArgb(244, 255, 255, 255)
  $cardPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
  Enable-DoubleBuffer $cardPanel
  # Keep card rendering static for reliability.

  $labelInstr = New-Object System.Windows.Forms.Label
  $labelInstr.Text = "仪器 (A列)"
  $labelInstr.Location = New-Object System.Drawing.Point(28, 32)
  $labelInstr.AutoSize = $true

  $comboInstr = New-Object System.Windows.Forms.ComboBox
  $comboInstr.Location = New-Object System.Drawing.Point(182, 28)
  $comboInstr.Size = New-Object System.Drawing.Size(270, 28)
  $comboInstr.DropDownStyle = "DropDown"
  Add-ComboItems $comboInstr $instList

  $btnInstrSave = New-Object System.Windows.Forms.Button
  $btnInstrSave.Text = "保存"
  $btnInstrSave.Location = New-Object System.Drawing.Point(552, 26)
  $btnInstrSave.Size = New-Object System.Drawing.Size(72, 30)

  $btnInstrDelete = New-Object System.Windows.Forms.Button
  $btnInstrDelete.Text = "删除"
  $btnInstrDelete.Location = New-Object System.Drawing.Point(634, 26)
  $btnInstrDelete.Size = New-Object System.Drawing.Size(72, 30)

  $labelNormal = New-Object System.Windows.Forms.Label
  $labelNormal.Text = "正常水平批号 (B列)"
  $labelNormal.Location = New-Object System.Drawing.Point(28, 78)
  $labelNormal.AutoSize = $true

  $comboNormal = New-Object System.Windows.Forms.ComboBox
  $comboNormal.Location = New-Object System.Drawing.Point(182, 74)
  $comboNormal.Size = New-Object System.Drawing.Size(270, 28)
  $comboNormal.DropDownStyle = "DropDown"
  Add-ComboItems $comboNormal $batchList

  $checkNormalTpl = New-Object System.Windows.Forms.CheckBox
  $checkNormalTpl.Text = "使用模板"
  $checkNormalTpl.Location = New-Object System.Drawing.Point(470, 79)
  $checkNormalTpl.AutoSize = $true

  $btnNormalSave = New-Object System.Windows.Forms.Button
  $btnNormalSave.Text = "保存"
  $btnNormalSave.Location = New-Object System.Drawing.Point(552, 74)
  $btnNormalSave.Size = New-Object System.Drawing.Size(72, 30)

  $btnNormalDelete = New-Object System.Windows.Forms.Button
  $btnNormalDelete.Text = "删除"
  $btnNormalDelete.Location = New-Object System.Drawing.Point(634, 74)
  $btnNormalDelete.Size = New-Object System.Drawing.Size(72, 30)

  $labelAbnormal = New-Object System.Windows.Forms.Label
  $labelAbnormal.Text = "异常水平批号 (C列)"
  $labelAbnormal.Location = New-Object System.Drawing.Point(28, 128)
  $labelAbnormal.AutoSize = $true

  $comboAbnormal = New-Object System.Windows.Forms.ComboBox
  $comboAbnormal.Location = New-Object System.Drawing.Point(182, 124)
  $comboAbnormal.Size = New-Object System.Drawing.Size(270, 28)
  $comboAbnormal.DropDownStyle = "DropDown"
  Add-ComboItems $comboAbnormal $batchList

  $checkAbnormalTpl = New-Object System.Windows.Forms.CheckBox
  $checkAbnormalTpl.Text = "使用模板"
  $checkAbnormalTpl.Location = New-Object System.Drawing.Point(470, 129)
  $checkAbnormalTpl.AutoSize = $true

  $btnAbnormalSave = New-Object System.Windows.Forms.Button
  $btnAbnormalSave.Text = "保存"
  $btnAbnormalSave.Location = New-Object System.Drawing.Point(552, 124)
  $btnAbnormalSave.Size = New-Object System.Drawing.Size(72, 30)

  $btnAbnormalDelete = New-Object System.Windows.Forms.Button
  $btnAbnormalDelete.Text = "删除"
  $btnAbnormalDelete.Location = New-Object System.Drawing.Point(634, 124)
  $btnAbnormalDelete.Size = New-Object System.Drawing.Size(72, 30)

  $labelTemplate = New-Object System.Windows.Forms.Label
  $labelTemplate.Text = "批号模板"
  $labelTemplate.Location = New-Object System.Drawing.Point(28, 178)
  $labelTemplate.AutoSize = $true

  $comboTemplate = New-Object System.Windows.Forms.ComboBox
  $comboTemplate.Location = New-Object System.Drawing.Point(182, 174)
  $comboTemplate.Size = New-Object System.Drawing.Size(270, 28)
  $comboTemplate.DropDownStyle = "DropDown"
  Add-ComboItems $comboTemplate $templateList
  $comboTemplate.Text = $templateSelected

  $btnTemplateSave = New-Object System.Windows.Forms.Button
  $btnTemplateSave.Text = "保存模板"
  $btnTemplateSave.Location = New-Object System.Drawing.Point(470, 174)
  $btnTemplateSave.Size = New-Object System.Drawing.Size(72, 30)

  $btnTemplateDefault = New-Object System.Windows.Forms.Button
  $btnTemplateDefault.Text = "设为默认"
  $btnTemplateDefault.Location = New-Object System.Drawing.Point(552, 174)
  $btnTemplateDefault.Size = New-Object System.Drawing.Size(72, 30)

  $btnTemplateDelete = New-Object System.Windows.Forms.Button
  $btnTemplateDelete.Text = "删除模板"
  $btnTemplateDelete.Location = New-Object System.Drawing.Point(634, 174)
  $btnTemplateDelete.Size = New-Object System.Drawing.Size(72, 30)

  $labelMissing = New-Object System.Windows.Forms.Label
  $labelMissing.Text = "缺失处理"
  $labelMissing.Location = New-Object System.Drawing.Point(28, 226)
  $labelMissing.AutoSize = $true

  $radioMissingEmpty = New-Object System.Windows.Forms.RadioButton
  $radioMissingEmpty.Text = "留空"
  $radioMissingEmpty.Location = New-Object System.Drawing.Point(182, 224)
  $radioMissingEmpty.AutoSize = $true
  $radioMissingEmpty.Checked = $true

  $radioMissingAvg = New-Object System.Windows.Forms.RadioButton
  $radioMissingAvg.Text = "均值填充"
  $radioMissingAvg.Location = New-Object System.Drawing.Point(254, 224)
  $radioMissingAvg.AutoSize = $true

  $btnRun = New-Object System.Windows.Forms.Button
  $btnRun.Text = "开始处理"
  $btnRun.Location = New-Object System.Drawing.Point(286, 280)
  $btnRun.Size = New-Object System.Drawing.Size(160, 38)

  # Apply consistent typography/color to all form labels.
  foreach ($label in @($labelInstr, $labelNormal, $labelAbnormal, $labelTemplate, $labelMissing)) {
    $label.ForeColor = $textPrimary
    $label.BackColor = [System.Drawing.Color]::Transparent
    $label.Font = $labelFont
  }

  # Input controls: flat style + focus accent for clearer keyboard flow.
  foreach ($combo in @($comboInstr, $comboNormal, $comboAbnormal, $comboTemplate)) {
    $combo.Font = $uiFont
    $combo.ForeColor = $textPrimary
    $combo.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    Add-FocusAccent $combo (New-Color "#ECF4FF")
  }

  # Toggle controls inherit neutral secondary text with hand cursor.
  foreach ($toggle in @($checkNormalTpl, $checkAbnormalTpl, $radioMissingEmpty, $radioMissingAvg)) {
    $toggle.ForeColor = $textSecondary
    $toggle.BackColor = [System.Drawing.Color]::Transparent
    $toggle.Font = $uiFont
    $toggle.Cursor = [System.Windows.Forms.Cursors]::Hand
  }

  # Use one button font style across all actions.
  foreach ($btn in @($btnInstrSave, $btnInstrDelete, $btnNormalSave, $btnNormalDelete, $btnAbnormalSave, $btnAbnormalDelete, $btnTemplateSave, $btnTemplateDefault, $btnTemplateDelete, $btnRun)) {
    $btn.Font = $buttonFont
  }

  # Color mapping by semantic action:
  # - green: save/confirm
  # - coral: destructive/delete
  # - blue: primary flow and template save
  # - yellow: set default (attention, non-destructive)
  Add-InteractiveButtonStyle $btnInstrSave $accentGreen (New-Color "#FFFFFF")
  Add-InteractiveButtonStyle $btnInstrDelete $accentCoral (New-Color "#FFFFFF")
  Add-InteractiveButtonStyle $btnNormalSave $accentGreen (New-Color "#FFFFFF")
  Add-InteractiveButtonStyle $btnNormalDelete $accentCoral (New-Color "#FFFFFF")
  Add-InteractiveButtonStyle $btnAbnormalSave $accentGreen (New-Color "#FFFFFF")
  Add-InteractiveButtonStyle $btnAbnormalDelete $accentCoral (New-Color "#FFFFFF")
  Add-InteractiveButtonStyle $btnTemplateSave $accentBlue (New-Color "#FFFFFF")
  Add-InteractiveButtonStyle $btnTemplateDefault $accentYellow (New-Color "#3F2A00")
  Add-InteractiveButtonStyle $btnTemplateDelete $accentCoral (New-Color "#FFFFFF")
  Add-InteractiveButtonStyle $btnRun $accentBlue (New-Color "#FFFFFF")

  # The old bottom hint labels are intentionally removed to keep UI cleaner.
  $cardPanel.Controls.AddRange(@(
    $labelInstr, $comboInstr, $btnInstrSave, $btnInstrDelete,
    $labelNormal, $comboNormal, $checkNormalTpl, $btnNormalSave, $btnNormalDelete,
    $labelAbnormal, $comboAbnormal, $checkAbnormalTpl, $btnAbnormalSave, $btnAbnormalDelete,
    $labelTemplate, $comboTemplate, $btnTemplateSave, $btnTemplateDefault, $btnTemplateDelete,
    $labelMissing, $radioMissingEmpty, $radioMissingAvg,
    $btnRun
  ))

  # Add layered controls then start fade-in transition.
  $form.Controls.AddRange(@($cardPanel, $labelAuthor))
  Start-FadeIn $form

  # ----------------------------- Event Handlers ------------------------------
  # Each handler validates input, updates dropdown history, persists config,
  # and gives immediate feedback through message boxes.
  $btnInstrSave.Add_Click({
    $value = Safe-Trim $comboInstr.Text
    if ([string]::IsNullOrWhiteSpace($value)) {
      Show-Error "请输入仪器名称"
      return
    }
    if (-not $comboInstr.Items.Contains($value)) {
      $comboInstr.Items.Add($value) | Out-Null
      Save-ConfigFromUI $comboInstr $comboNormal $comboTemplate $script:TemplateDefault $excludeList
      Show-Info "已保存"
    } else {
      Show-Info "已存在"
    }
  })

  $btnInstrDelete.Add_Click({
    $value = Safe-Trim $comboInstr.Text
    if ([string]::IsNullOrWhiteSpace($value)) {
      Show-Error "请输入要删除的仪器名称"
      return
    }
    if ($comboInstr.Items.Contains($value)) {
      $comboInstr.Items.Remove($value) | Out-Null
      Save-ConfigFromUI $comboInstr $comboNormal $comboTemplate $script:TemplateDefault $excludeList
      Show-Info "已删除"
    } else {
      Show-Info "未找到"
    }
  })

  $btnNormalSave.Add_Click({
    $value = Build-BatchValue $comboNormal.Text $checkNormalTpl.Checked $comboTemplate.Text
    if ([string]::IsNullOrWhiteSpace($value)) {
      Show-Error "请输入正常批号"
      return
    }
    if (-not $comboNormal.Items.Contains($value)) {
      Add-BatchItem $comboNormal $comboAbnormal $value
      Save-ConfigFromUI $comboInstr $comboNormal $comboTemplate $script:TemplateDefault $excludeList
      Show-Info "已保存"
    } else {
      Show-Info "已存在"
    }
  })

  $btnNormalDelete.Add_Click({
    $value = Safe-Trim $comboNormal.Text
    if ([string]::IsNullOrWhiteSpace($value)) {
      Show-Error "请输入要删除的批号"
      return
    }
    if ($comboNormal.Items.Contains($value)) {
      Remove-BatchItem $comboNormal $comboAbnormal $value
      Save-ConfigFromUI $comboInstr $comboNormal $comboTemplate $script:TemplateDefault $excludeList
      Show-Info "已删除"
    } else {
      Show-Info "未找到"
    }
  })

  $btnAbnormalSave.Add_Click({
    $value = Build-BatchValue $comboAbnormal.Text $checkAbnormalTpl.Checked $comboTemplate.Text
    if ([string]::IsNullOrWhiteSpace($value)) {
      Show-Error "请输入异常批号"
      return
    }
    if (-not $comboAbnormal.Items.Contains($value)) {
      Add-BatchItem $comboNormal $comboAbnormal $value
      Save-ConfigFromUI $comboInstr $comboNormal $comboTemplate $script:TemplateDefault $excludeList
      Show-Info "已保存"
    } else {
      Show-Info "已存在"
    }
  })

  $btnAbnormalDelete.Add_Click({
    $value = Safe-Trim $comboAbnormal.Text
    if ([string]::IsNullOrWhiteSpace($value)) {
      Show-Error "请输入要删除的批号"
      return
    }
    if ($comboAbnormal.Items.Contains($value)) {
      Remove-BatchItem $comboNormal $comboAbnormal $value
      Save-ConfigFromUI $comboInstr $comboNormal $comboTemplate $script:TemplateDefault $excludeList
      Show-Info "已删除"
    } else {
      Show-Info "未找到"
    }
  })

  $btnTemplateSave.Add_Click({
    $tpl = Safe-Trim $comboTemplate.Text
    if ([string]::IsNullOrWhiteSpace($tpl)) {
      Show-Error "请输入模板内容"
      return
    }
    if (-not $comboTemplate.Items.Contains($tpl)) {
      $comboTemplate.Items.Add($tpl) | Out-Null
    }
    Save-ConfigFromUI $comboInstr $comboNormal $comboTemplate $script:TemplateDefault $excludeList
    Show-Info "模板已保存"
  })

  $btnTemplateDefault.Add_Click({
    $tpl = Safe-Trim $comboTemplate.Text
    if ([string]::IsNullOrWhiteSpace($tpl)) {
      Show-Error "请输入模板内容"
      return
    }
    if (-not $comboTemplate.Items.Contains($tpl)) {
      $comboTemplate.Items.Add($tpl) | Out-Null
    }
    $script:TemplateDefault = $tpl
    Save-ConfigFromUI $comboInstr $comboNormal $comboTemplate $script:TemplateDefault $excludeList
    Show-Info "已设为默认模板"
  })

  $btnTemplateDelete.Add_Click({
    $tpl = Safe-Trim $comboTemplate.Text
    if ([string]::IsNullOrWhiteSpace($tpl)) {
      Show-Error "请输入要删除的模板"
      return
    }
    if ($comboTemplate.Items.Contains($tpl)) {
      $comboTemplate.Items.Remove($tpl) | Out-Null
      if ($tpl -eq $script:TemplateDefault) {
        $script:TemplateDefault = "常规化学{year}{value}"
      }
      if (-not $comboTemplate.Items.Contains($comboTemplate.Text)) {
        $comboTemplate.Text = $script:TemplateDefault
      }
      if (-not $comboTemplate.Items.Contains($comboTemplate.Text) -and -not [string]::IsNullOrWhiteSpace($comboTemplate.Text)) {
        $comboTemplate.Items.Add($comboTemplate.Text) | Out-Null
      }
      Save-ConfigFromUI $comboInstr $comboNormal $comboTemplate $script:TemplateDefault $excludeList
      Show-Info "已删除"
    } else {
      Show-Info "未找到"
    }
  })

  # Main processing action:
  # Validate inputs -> collect files -> align dates -> export template output.
  $btnRun.Add_Click({
    try {
      if (-not (Test-Path -LiteralPath $ExportDir)) {
        Show-Error "找不到 qc_export 文件夹"
        return
      }
      if (-not (Test-Path -LiteralPath $TemplatePath)) {
        Show-Error "找不到 templet.xls"
        return
      }

      $instrument = Safe-Trim $comboInstr.Text
      if ([string]::IsNullOrWhiteSpace($instrument)) {
        Show-Error "请填写仪器 (A列)"
        return
      }

      $batchNormal = Build-BatchValue $comboNormal.Text $checkNormalTpl.Checked $comboTemplate.Text
      $batchAbnormal = Build-BatchValue $comboAbnormal.Text $checkAbnormalTpl.Checked $comboTemplate.Text
      if ([string]::IsNullOrWhiteSpace($batchNormal) -or [string]::IsNullOrWhiteSpace($batchAbnormal)) {
        Show-Error "请填写正常/异常批号 (B/C列)"
        return
      }

      $updated = $false
      if (Add-ComboItem $comboInstr $instrument) { $updated = $true }
      if (-not $comboNormal.Items.Contains($batchNormal) -or -not $comboAbnormal.Items.Contains($batchNormal)) {
        Add-BatchItem $comboNormal $comboAbnormal $batchNormal
        $updated = $true
      }
      if (-not $comboNormal.Items.Contains($batchAbnormal) -or -not $comboAbnormal.Items.Contains($batchAbnormal)) {
        Add-BatchItem $comboNormal $comboAbnormal $batchAbnormal
        $updated = $true
      }
      if ($checkNormalTpl.Checked -or $checkAbnormalTpl.Checked) {
        Add-ComboItem $comboTemplate $comboTemplate.Text | Out-Null
      }
      Save-ConfigFromUI $comboInstr $comboNormal $comboTemplate $script:TemplateDefault $excludeList

    $exclude = $excludeList
    if ($exclude.Count -eq 0) {
      $exclude = Get-ExcludeKeywords
    }
      $files = Get-ChildItem -LiteralPath $ExportDir -File -Filter "*.xls" | Where-Object { $_.Name -notlike "~$*" }
      if ($exclude.Count -gt 0) {
        $files = $files | Where-Object {
          $name = $_.Name
          -not ($exclude | Where-Object { $name -match [regex]::Escape($_) })
        }
      }

      if ($files.Count -eq 0) {
        Show-Error "没有可处理的 Excel 文件"
        return
      }

      $missingMode = "blank"
      if ($radioMissingAvg.Checked) { $missingMode = "avg" }

      $result = Build-OutputRows $files $instrument $batchNormal $batchAbnormal $missingMode
      $rows = $result.Rows
      $warnings = $result.Warnings
      if ($rows.Count -eq 0) {
        Show-Error "未读取到任何数据"
        return
      }

      $rangeToken = Get-DateRangeToken $files
      if ([string]::IsNullOrWhiteSpace($rangeToken)) {
        $rangeToken = (Get-Date -Format "yyyyMMdd")
      }
      $safeInstrument = Sanitize-FileName $instrument
      $safeNormal = Sanitize-FileName $batchNormal
      $safeAbnormal = Sanitize-FileName $batchAbnormal
      $outName = "{0}_{1}({2}_{3}).xls" -f $safeInstrument, $rangeToken, $safeNormal, $safeAbnormal
      $outPath = Write-Template $rows $outName
      if ($warnings.Count -gt 0) {
        $warnText = ($warnings | Select-Object -First 20) -join "`r`n"
        if ($warnings.Count -gt 20) {
          $warnText += "`r`n... 共 {0} 条" -f $warnings.Count
        }
        Show-Warn ("已生成，但发现缺失日期:`r`n{0}" -f $warnText)
      }
      Show-Info ("已生成: {0}" -f $outPath)
    } catch {
      Show-Error $_.Exception.Message
    }
  })

  [void]$form.ShowDialog()
} catch {
  $detail = $_ | Out-String
  $info = $_.InvocationInfo | Format-List * | Out-String
  $msg = "未处理异常:`r`n$detail`r`n$info"
  Write-ErrorLog "未处理异常(TopLevelCatch)" $msg
  Show-Fatal "程序异常，详情已写入 qc_tool_error.log。"
}

