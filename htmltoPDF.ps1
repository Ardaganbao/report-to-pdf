
param (
  [Parameter(Mandatory = $true, HelpMessage = 'authcookies_key')]
  [string]$authcookies_key,
  [Parameter(Mandatory = $true, HelpMessage = 'authcookies_value')]
  [string]$authcookies_value,
  [Parameter(Mandatory = $true, HelpMessage = 'root URL')]
  [string]$root_URL,
  [Parameter(Mandatory = $true, HelpMessage = 'post to print')]
  [string]$article,
  [Parameter(Mandatory = $true, HelpMessage = 'outputPDF')]
  [string]$outputPDF,
  [Parameter(Mandatory = $true, HelpMessage = 'Page counter postscript : C:\temp_pdf\pagecount.ps')]
  [string]$pageCounterPS,
  [Parameter(Mandatory = $true, HelpMessage = 'Log folder : C:\temp_pdf\ ')]
  [string]$Log_folder 

  
)

Add-Type -AssemblyName Microsoft.PowerShell.Commands.Utility

$wkhtmltopdfPath = (Get-Command wkhtmltopdf.exe -ErrorAction SilentlyContinue).Source
if (-not $wkhtmltopdfPath) {
  $wkhtmltopdfPath = Join-Path $env:ProgramFiles 'wkhtmltopdf\bin\wkhtmltopdf.exe'
}
if (-not (Test-Path $wkhtmltopdfPath)) {
  throw 'wkhtmltopdf.exe was not found. Run install.ps1 as administrator first.'
}

$ghostscriptPath = (Get-Command gswin64c.exe -ErrorAction SilentlyContinue).Source
if (-not $ghostscriptPath) {
  $ghostscriptPath = Get-ChildItem -Path (Join-Path $env:ProgramFiles 'gs\*\bin\gswin64c.exe') -ErrorAction SilentlyContinue |
  Sort-Object LastWriteTime -Descending |
  Select-Object -First 1 -ExpandProperty FullName
}
if (-not $ghostscriptPath) {
  throw 'gswin64c.exe was not found. Run install.ps1 as administrator first.'
}

# wkhtmltopdf cannot run Chart.js 4 (its WebKit is too old), so headless Chrome or Edge opens the page
# first and every ``` graph block drawn by view.js is replaced by a picture of its chart.
function Add-GraphImages {
  param ([string]$Html, [string]$WorkFolder)

  # Chart.js graphs (ardareport/view.js) and Mermaid diagrams both render in the browser only; the
  # browser below draws them so wkhtmltopdf, whose engine cannot, gets pictures.
  $hasGraphs = $Html -match 'id=["'']ardareport-graphs-js["'']'
  $hasMermaid = $Html -match 'class="(?:[^"]*\s)?language-mermaid'
  if (-not $hasGraphs -and -not $hasMermaid) {
    return $Html
  }
  $browserPath = @(
    (Join-Path $env:ProgramFiles 'Google\Chrome\Application\chrome.exe')
    (Join-Path ${env:ProgramFiles(x86)} 'Google\Chrome\Application\chrome.exe')
    (Join-Path $env:LOCALAPPDATA 'Google\Chrome\Application\chrome.exe')
    (Join-Path ${env:ProgramFiles(x86)} 'Microsoft\Edge\Application\msedge.exe')
    (Join-Path $env:ProgramFiles 'Microsoft\Edge\Application\msedge.exe')
  ) | Where-Object { Test-Path $_ } | Select-Object -First 1
  if (-not $browserPath) {
    Write-Warning 'Chrome or Edge was not found: the graphs stay as text in the PDF.'
    return $Html
  }
  $bodyEnd = $Html.LastIndexOf('</body>', [StringComparison]::OrdinalIgnoreCase)
  if ($bodyEnd -lt 0) {
    return $Html
  }

  # Once view.js has drawn the graphs, copy each block with its canvases turned into PNG images.
  $collector = @'
<script>
window.addEventListener('load', function () {
  var started = Date.now();
  // Mermaid draws asynchronously and marks each diagram data-processed when done; wait for them.
  function pending() {
    return document.querySelectorAll('.mermaid:not([data-processed])').length;
  }
  // Chart.js graphs: a copy of each ardareport block with its canvases turned into PNG images.
  function drawGraphs() {
    var blocks = document.querySelectorAll('[data-ardareport-graph]'), graphs = [];
    for (var i = 0; i < blocks.length; i++) {
      if (blocks[i].getAttribute('data-ardareport-graph') === 'Ardagram') {
        continue;
      }
      var container = blocks[i].parentNode.tagName === 'PRE' ? blocks[i].parentNode : blocks[i];
      var copy = container.cloneNode(true), canvases = container.getElementsByTagName('canvas'), copies = copy.getElementsByTagName('canvas');
      for (var j = canvases.length - 1; j >= 0; j--) {
        var image = document.createElement('img');
        image.src = canvases[j].toDataURL('image/png');
        image.alt = canvases[j].getAttribute('aria-label') || '';
        image.style.cssText = canvases[j].style.cssText;
        copies[j].parentNode.replaceChild(image, copies[j]);
      }
      graphs.push(copy.outerHTML);
    }
    var output = document.createElement('pre');
    output.id = 'ardareport-graph-images';
    output.textContent = graphs.join('<!--ardareport-graph-->');
    document.body.appendChild(output);
  }
  // Mermaid diagrams: each rendered SVG drawn to a PNG so wkhtmltopdf keeps the labels (it does not
  // draw foreignObject). One picture per .mermaid block, in page order; the SVG itself as a fallback.
  function drawMermaid() {
    var blocks = document.querySelectorAll('.mermaid'), results = new Array(blocks.length), remaining = blocks.length;
    function done() {
      var holder = document.createElement('pre');
      holder.id = 'ardareport-mermaid-images';
      holder.textContent = [].join.call(results, '<!--ardareport-graph-->');
      document.body.appendChild(holder);
    }
    if (remaining === 0) {
      done();
      return;
    }
    for (var i = 0; i < blocks.length; i++) {
      (function (idx) {
        var svg = blocks[idx].querySelector('svg');
        function settle(html) {
          results[idx] = html;
          if (--remaining === 0) {
            done();
          }
        }
        if (!svg) {
          settle('<pre style="text-align:center;background:none;border:0;">' + blocks[idx].outerHTML + '</pre>');
          return;
        }
        var box = svg.getBoundingClientRect();
        var w = Math.max(1, Math.round(box.width || svg.clientWidth || 800));
        var h = Math.max(1, Math.round(box.height || svg.clientHeight || 600));
        var clone = svg.cloneNode(true);
        clone.setAttribute('width', w);
        clone.setAttribute('height', h);
        var xml = new XMLSerializer().serializeToString(clone);
        function fallback() {
          settle('<pre style="text-align:center;background:none;border:0;">' + svg.outerHTML + '</pre>');
        }
        var img = new Image();
        img.onload = function () {
          try {
            var canvas = document.createElement('canvas');
            canvas.width = w * 2;
            canvas.height = h * 2;
            var context = canvas.getContext('2d');
            context.setTransform(2, 0, 0, 2, 0, 0);
            context.drawImage(img, 0, 0, w, h);
            settle('<pre style="text-align:center;background:none;border:0;"><img style="max-width:100%;height:auto;" src="' + canvas.toDataURL('image/png') + '"></pre>');
          } catch (e) {
            fallback();
          }
        };
        img.onerror = fallback;
        img.src = 'data:image/svg+xml;charset=utf-8,' + encodeURIComponent(xml);
      })(i);
    }
  }
  var timer = setInterval(function () {
    if (pending() === 0 || Date.now() - started > 8000) {
      clearInterval(timer);
      drawGraphs();
      drawMermaid();
    }
  }, 200);
});
</script>
'@
  $sourceFile = Join-Path $WorkFolder 'graphs.html'
  $dumpFile = Join-Path $WorkFolder 'graphs-dump.html'
  Set-Content -LiteralPath $sourceFile -Value $Html.Insert($bodyEnd, $collector) -Encoding UTF8

  # A short profile path: the browser's cache folders are deep and break the 260 characters limit.
  $profileFolder = Join-Path $env:TEMP ('report-to-pdf-' + [Guid]::NewGuid().ToString('N').Substring(0, 8))
  $browserArguments = '--headless', '--disable-gpu', '--no-first-run', '--no-default-browser-check', "--user-data-dir=`"$profileFolder`"", '--virtual-time-budget=15000', '--dump-dom', "`"$sourceFile`""
  $browser = Start-Process -NoNewWindow -PassThru -FilePath $browserPath -ArgumentList $browserArguments -RedirectStandardOutput $dumpFile -RedirectStandardError (Join-Path $WorkFolder 'graphs.log')
  # Edge waits forever when a policy forces the sign-in.
  if (-not $browser.WaitForExit(30000)) {
    & taskkill.exe /T /F /PID $browser.Id | Out-Null
    # The killed processes free the profile files a moment later.
    Start-Sleep -Seconds 2
  }
  Remove-Item -LiteralPath $profileFolder -Recurse -Force -ErrorAction SilentlyContinue

  $dump = [string](Get-Content -LiteralPath $dumpFile -Raw -Encoding UTF8 -ErrorAction SilentlyContinue)

  # The Chart.js graphs: replace each ardareport code block with the picture of its chart.
  if ($hasGraphs) {
    $output = [regex]::Match($dump, '<pre id="ardareport-graph-images">([\s\S]*?)</pre>')
    $graphText = if ($output.Success) { [System.Net.WebUtility]::HtmlDecode($output.Groups[1].Value) } else { '' }
    if ($graphText) {
      $graphs = $graphText.Split([string[]]@('<!--ardareport-graph-->'), [StringSplitOptions]::None)
      # The same blocks in the fetched HTML, in page order: a code with the graph class inside a pre, or the pre itself.
      $types = @([regex]::Matches($graphText, 'data-ardareport-graph="([^"]+)"') | ForEach-Object { [regex]::Escape($_.Groups[1].Value) } | Select-Object -Unique)
      if ($types.Count -gt 0) {
        $classPattern = 'class="(?:[^"]*\s)?language-(?:' + ($types -join '|') + ')(?:\s[^"]*)?"'
        $blockPattern = '<pre\b[^>]*>\s*<code\b[^>]*' + $classPattern + '[^>]*>[\s\S]*?</code>\s*</pre>|<pre\b[^>]*' + $classPattern + '[^>]*>[\s\S]*?</pre>'
        $blockCount = [regex]::Matches($Html, $blockPattern).Count
        if ($blockCount -eq $graphs.Count) {
          Write-Host "Graphs drawn as images: $($graphs.Count)"
          $graphQueue = [System.Collections.Queue]::new($graphs)
          $Html = [regex]::Replace($Html, $blockPattern, [System.Text.RegularExpressions.MatchEvaluator] { param($match) $graphQueue.Dequeue() })
        }
        else {
          Write-Warning "The page has $blockCount graph blocks but $($graphs.Count) graphs were drawn: they stay as text in the PDF."
        }
      }
    }
    elseif (-not $hasMermaid) {
      Write-Warning "$(Split-Path -Leaf $browserPath) did not draw the graphs (see graphs.log): they stay as text in the PDF."
    }
  }

  # The Mermaid diagrams: replace each language-mermaid code block with the picture of its diagram.
  if ($hasMermaid) {
    $output = [regex]::Match($dump, '<pre id="ardareport-mermaid-images">([\s\S]*?)</pre>')
    $mermaidText = if ($output.Success) { [System.Net.WebUtility]::HtmlDecode($output.Groups[1].Value) } else { '' }
    if ($mermaidText) {
      $diagrams = $mermaidText.Split([string[]]@('<!--ardareport-graph-->'), [StringSplitOptions]::None)
      $mermaidPattern = '<pre\b[^>]*>\s*<code\b[^>]*class="(?:[^"]*\s)?language-mermaid(?:\s[^"]*)?"[^>]*>[\s\S]*?</code>\s*</pre>'
      $mermaidCount = [regex]::Matches($Html, $mermaidPattern).Count
      if ($mermaidCount -eq $diagrams.Count) {
        Write-Host "Mermaid diagrams drawn as images: $($diagrams.Count)"
        $mermaidQueue = [System.Collections.Queue]::new($diagrams)
        $Html = [regex]::Replace($Html, $mermaidPattern, [System.Text.RegularExpressions.MatchEvaluator] { param($match) $mermaidQueue.Dequeue() })
      }
      else {
        Write-Warning "The page has $mermaidCount Mermaid blocks but $($diagrams.Count) were drawn: they stay as text in the PDF."
      }
    }
    else {
      Write-Warning "$(Split-Path -Leaf $browserPath) did not draw the Mermaid diagrams (see graphs.log): they stay as text in the PDF."
    }
  }

  return $Html
}
 
# A ``` mose-actions fenced block becomes a Red/Orange/Green action-tracker ("Suivi des actions").
# red/orange/green (or rouge/vert, R/O/G, ok/done...) in the Status column tints each row; a Synthese
# line with the counts is added below. Built as inline-styled HTML so it survives WordPress and prints
# in wkhtmltopdf, whose engine draws the table and backgrounds fine (only modern JavaScript fails it).
function Get-MoseActionStatus {
  param([string]$Value)
  $v = ("$Value").Trim().ToLowerInvariant()
  if (-not $v) { return $null }
  $red = @('r', 'red', 'rouge', 'rood', 'rot', 'critical', 'crit', 'high', 'urgent', 'ko', 'nok', '3')
  $orange = @('o', 'a', 'orange', 'amber', 'oranje', 'medium', 'med', 'warning', 'warn', 'attention', 'ongoing', 'wip', 'in progress', 'en cours', '2')
  $green = @('g', 'v', 'green', 'vert', 'groen', 'ok', 'done', 'resolved', 'resolu', 'closed', 'clos', 'fait', 'low', '1')
  if ($red -contains $v) { return 'red' }
  if ($orange -contains $v) { return 'orange' }
  if ($green -contains $v) { return 'green' }
  switch ($v[0]) {
    'r' { return 'red' }
    'o' { return 'orange' }
    'a' { return 'orange' }
    'g' { return 'green' }
    'v' { return 'green' }
    default { return $null }
  }
}

function ConvertTo-MoseActionsTable {
  param([string]$Text)

  $enc = { param($s) [System.Net.WebUtility]::HtmlEncode("$s") }
  $palette = @{
    red     = @{ Row = '#fdecea'; Bar = '#d93025'; Pill = '#d93025' }
    orange  = @{ Row = '#fff4e5'; Bar = '#e8710a'; Pill = '#e8710a' }
    green   = @{ Row = '#e7f4ea'; Bar = '#188038'; Pill = '#188038' }
    neutral = @{ Row = '#ffffff'; Bar = '#cfd4da'; Pill = '#5f6368' }
  }

  $title = $null
  $parsed = [System.Collections.Generic.List[string[]]]::new()
  foreach ($line in ($Text -split "\r?\n")) {
    $trimmed = $line.Trim()
    if (-not $trimmed) { continue }
    if ($trimmed.StartsWith('#')) { $title = $trimmed.TrimStart('#').Trim(); continue }
    if ($trimmed -match '^\|?\s*[-:| ]+\s*\|?$' -and $trimmed -match '-') { continue } # a markdown divider row
    $parts = $trimmed -split '\|'
    $startIdx = 0
    $endIdx = $parts.Count - 1
    if ($trimmed.StartsWith('|')) { $startIdx = 1 }
    if ($trimmed.EndsWith('|')) { $endIdx = $parts.Count - 2 }
    $cells = [System.Collections.Generic.List[string]]::new()
    for ($i = $startIdx; $i -le $endIdx; $i++) { $cells.Add($parts[$i].Trim()) }
    if ($cells.Count -gt 0) { $parsed.Add($cells.ToArray()) }
  }
  if ($parsed.Count -eq 0) { return '' }

  # The header row: the first line whose cells name the columns. Otherwise the workbook's column order.
  $headerWords = 'issue|solution|owner|due|date|status|statut|comment|action|probl|responsable|.cheance|priorit'
  $header = $null
  $statusIndex = -1; $ownerIndex = -1; $dueIndex = -1
  $first = $parsed[0]
  if (@($first | Where-Object { $_ -match "(?i)$headerWords" }).Count -ge 2) {
    $header = $first
    $parsed.RemoveAt(0)
    for ($i = 0; $i -lt $header.Count; $i++) {
      if ($statusIndex -lt 0 -and $header[$i] -match '(?i)status|statut|rag') { $statusIndex = $i }
      if ($ownerIndex -lt 0 -and $header[$i] -match '(?i)owner|responsable|assign') { $ownerIndex = $i }
      if ($dueIndex -lt 0 -and $header[$i] -match '(?i)due|date|.ch.ance|deadline') { $dueIndex = $i }
    }
  }
  else {
    $header = @('Issue', 'Solution', 'Owner', 'Due date', 'Status', 'Comments')
    $statusIndex = 4; $ownerIndex = 2; $dueIndex = 3
  }
  $columns = $header.Count

  # When no column was named "status", the one whose values all read as red/orange/green.
  if ($statusIndex -lt 0) {
    for ($c = 0; $c -lt $columns; $c++) {
      $hits = 0; $vals = 0
      foreach ($row in $parsed) {
        if ($c -lt $row.Count -and $row[$c]) { $vals++; if (Get-MoseActionStatus $row[$c]) { $hits++ } }
      }
      if ($vals -gt 0 -and $hits -eq $vals) { $statusIndex = $c; break }
    }
  }

  $counts = @{ red = 0; orange = 0; green = 0; neutral = 0 }
  $noOwner = 0; $withDue = 0

  $sb = [System.Text.StringBuilder]::new()
  [void]$sb.Append('<div class="mose-actions" style="margin:14px 0;">')
  if ($title) {
    [void]$sb.Append('<p style="font-family:''Fira Sans'',Arial,sans-serif;font-weight:bold;font-size:13pt;margin:0 0 6px;">' + (& $enc $title) + '</p>')
  }
  [void]$sb.Append('<table style="border-collapse:collapse;width:100%;font-family:''Fira Sans'',Arial,sans-serif;font-size:10pt;">')
  [void]$sb.Append('<thead><tr>')
  for ($i = 0; $i -lt $columns; $i++) {
    $align = if ($i -eq $statusIndex) { 'center' } else { 'left' }
    [void]$sb.Append('<th style="background:#1f3864;color:#fff;text-align:' + $align + ';padding:6px 8px;border:1px solid #cfd4da;font-weight:bold;">' + (& $enc $header[$i]) + '</th>')
  }
  [void]$sb.Append('</tr></thead><tbody>')

  foreach ($row in $parsed) {
    $status = if ($statusIndex -ge 0 -and $statusIndex -lt $row.Count) { Get-MoseActionStatus $row[$statusIndex] } else { $null }
    $key = if ($status) { $status } else { 'neutral' }
    $counts[$key]++
    if ($ownerIndex -ge 0) { if ($ownerIndex -ge $row.Count -or -not $row[$ownerIndex]) { $noOwner++ } }
    if ($dueIndex -ge 0) { if ($dueIndex -lt $row.Count -and $row[$dueIndex]) { $withDue++ } }
    $tint = $palette[$key]
    [void]$sb.Append('<tr style="background:' + $tint.Row + ';">')
    for ($i = 0; $i -lt $columns; $i++) {
      $value = if ($i -lt $row.Count) { $row[$i] } else { '' }
      if ($i -eq 0) {
        $cellStyle = 'padding:6px 8px;border:1px solid #cfd4da;vertical-align:top;border-left:4px solid ' + $tint.Bar + ';'
      }
      else {
        $cellStyle = 'padding:6px 8px;border:1px solid #cfd4da;vertical-align:top;'
      }
      if ($i -eq $statusIndex) {
        if ($status) {
          $label = if ($value) { & $enc $value } else { $key.Substring(0, 1).ToUpper() + $key.Substring(1) }
          $pill = '<span style="display:inline-block;padding:2px 12px;border-radius:11px;background:' + $tint.Pill + ';color:#fff;font-weight:bold;white-space:nowrap;">' + $label + '</span>'
          [void]$sb.Append('<td style="' + $cellStyle + 'text-align:center;">' + $pill + '</td>')
        }
        else {
          [void]$sb.Append('<td style="' + $cellStyle + 'text-align:center;">' + (& $enc $value) + '</td>')
        }
      }
      else {
        [void]$sb.Append('<td style="' + $cellStyle + '">' + (& $enc $value) + '</td>')
      }
    }
    [void]$sb.Append('</tr>')
  }
  [void]$sb.Append('</tbody></table>')

  # Synthese: the counts, the way the workbook's Synthese sheet summed them.
  $total = $parsed.Count
  $synth = '<p style="font-family:''Fira Sans'',Arial,sans-serif;font-size:9.5pt;margin:6px 0 0;"><strong>Synth&#232;se &#8212;</strong> Total : ' + $total
  $synth += ' &#183; <span style="color:#d93025;font-weight:bold;">&#9679; Red : ' + $counts.red + '</span>'
  $synth += ' &#183; <span style="color:#e8710a;font-weight:bold;">&#9679; Orange : ' + $counts.orange + '</span>'
  $synth += ' &#183; <span style="color:#188038;font-weight:bold;">&#9679; Green : ' + $counts.green + '</span>'
  if ($ownerIndex -ge 0) { $synth += ' &#183; Sans owner : ' + $noOwner }
  if ($dueIndex -ge 0) { $synth += ' &#183; Avec &#233;ch&#233;ance : ' + $withDue }
  $synth += '</p>'
  [void]$sb.Append($synth)
  [void]$sb.Append('</div>')
  return $sb.ToString()
}

function Expand-MoseActions {
  param([string]$Html)
  $pattern = '<pre\b[^>]*>\s*<code\b[^>]*class="(?:[^"]*\s)?language-mose-actions(?:\s[^"]*)?"[^>]*>([\s\S]*?)</code>\s*</pre>'
  if (-not [regex]::IsMatch($Html, $pattern)) { return $Html }
  $count = 0
  $Html = [regex]::Replace($Html, $pattern, [System.Text.RegularExpressions.MatchEvaluator] {
      param($match)
      $raw = [System.Net.WebUtility]::HtmlDecode($match.Groups[1].Value)
      $table = ConvertTo-MoseActionsTable -Text $raw
      if ($table) { $script:__moseActionsHit = $true; $table } else { $match.Value }
    })
  $count = ([regex]::Matches($Html, 'class="mose-actions"')).Count
  if ($count -gt 0) { Write-Host "Action-tracker tables drawn: $count" }
  return $Html
}

$session = [Microsoft.PowerShell.Commands.WebRequestSession]::new()
$cookie = [System.Net.Cookie]::new($authcookies_key, $authcookies_value)
$session.Cookies.Add($root_URL, $cookie)

$null = New-Item -ItemType Directory -Path $Log_folder -Force
$TempFolder = Join-Path $Log_folder ([Guid]::NewGuid().ToString('N'))
$null = New-Item -ItemType Directory -Path $TempFolder -Force
Write-Host "Temporary files: $TempFolder"
$printStylesheet = Join-Path $TempFolder 'print-layout.css'
$printCss = @'

'@
 
$coverPrintCss = @'
.wp-block-group.is-nowrap {
  position: fixed !important;
  left: 10mm !important;
  right: 10mm !important;
  bottom: 18mm !important;
  display: table !important;
  table-layout: fixed !important;
  border-collapse: collapse !important;
  margin: 0 !important;
}
@page {
    size: A4 portrait;
    margin: 0;
}

html,
body {
    width: 210mm;
    margin: 0 !important;
    padding: 0 !important;
}

header  {

   
    width: 900px !important;
    height: 1300px !important; 

    margin: 0 !important;
    padding: 0 !important;

    display: grid;
    grid-template-columns: auto;
    grid-template-rows: 0px 530px repeat(4, min-content) auto 10px 10px;
    gap: 0px 0px;
    grid-template-areas:         "emptyHeader"        "logo"        " Title01"        " title02"        " client"        " other"        " emptypage"        " empty"        " footer";
      
    background-image: url("http://gi.grassone.be/wp-content/uploads/sites/5/2026/10/image002-e1791383454105.jpg") !important;
    background-repeat: no-repeat !important;
    background-position: top left !important;
    background-size: 900px !important;

}

.wp-block-create-block-multi-block-title {
    font-family: "Fira Sans", Arial, sans-serif;
    font-size: 26pt;
    font-weight: bold;
    text-align: center;
    padding: 2px;
    background: none;
    color: #000 !important;
}

h1.wp-block-create-block-multi-block-title {
    background: none !important;
    color: #000 !important;
    padding: 2px;
    font-family: "Fira Sans", Arial, sans-serif;
    font-size: 22pt;
    text-align: center; 
}

.wp-block-create-block-multi-block-client {
    background: none !important;
    color: #000 !important;
    padding: 2px;
    font-family: "Fira Sans", Arial, sans-serif;
    font-size: 20pt;
    text-align: center; 
}

#wp--skip-link--target > header > div.wp-block-group {
    min-height: 555px !important; 
}


.Footer{
position: fixed !important;
bottom: 15px !important;
left: 0 !important;
}
 .Footer .wp-block-group.is-nowrap {
    text-align: center !important;
    width: 100% !important;
}

.Footer .wp-block-group.is-nowrap > p {
    display: inline-block !important;
    margin: 0 20px !important;
    white-space: nowrap !important;
    vertical-align: middle !important;
}

.Footer .wp-block-group.is-nowrap > figure {
    display: block !important;
    text-align: center !important;
    margin: 0 auto 15px auto !important;
}
'@
$url = "$root_URL/?p=$article" 

$htmlContent = Invoke-RestMethod   -Uri $url  -WebSession $session
$htmlContent = Add-GraphImages -Html $htmlContent -WorkFolder $TempFolder
$htmlContent = Expand-MoseActions -Html $htmlContent
$checkEmoji = [char]0x2705
$htmlContent = $htmlContent.Replace($checkEmoji.ToString(), '<img align="absmiddle" alt=":white_check_mark:" class="emoji" src="https://gi.grassone.be/wp-content/plugins/wp-githuber-md/assets/vendor/emojify/images/white_check_mark.png" title=":white_check_mark:">')

 
$checkEmoji = '🔴'
$htmlContent = $htmlContent.Replace($checkEmoji.ToString(), '<img align="absmiddle" alt=":red_circle:" class="emoji" src="https://gi.grassone.be/wp-content/plugins/wp-githuber-md/assets/vendor/emojify/images/red_circle.png" title=":red_circle:">')
$checkEmoji = '🟠'
$htmlContent = $htmlContent.Replace($checkEmoji.ToString(), '<img align="absmiddle" alt=":orange_circle:" class="emoji" src="http://gi.grassone.be/wp-content/uploads/sites/5/2026/10/large_orange_circle.png" title=":orange_circle:">')

$checkEmoji = '🟢'
$htmlContent = $htmlContent.Replace($checkEmoji.ToString(), '<img align="absmiddle" alt=":green_circle:" class="emoji" src="http://gi.grassone.be/wp-content/uploads/sites/5/2026/10/large_green_circle.png" title=":green_circle:">')

$checkEmoji = '❌'
$htmlContent = $htmlContent.Replace($checkEmoji.ToString(), '<img align="absmiddle" alt=":x:" class="emoji" src="https://gi.grassone.be/wp-content/plugins/wp-githuber-md/assets/vendor/emojify/images/x.png" title=":x:">')

$checkEmoji = '⚠️'
$htmlContent = $htmlContent.Replace($checkEmoji.ToString(), '<img align="absmiddle" alt=":warning:" class="emoji" src="https://markdown-editor.github.io/emoji/warning.png" title=":warning:">')

$layoutClassMatches = [regex]::Matches($htmlContent, 'class=["'']([^"'']+)["'']', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
$layoutClassNames = @(
  foreach ($classMatch in $layoutClassMatches) {
    $classValue = $classMatch.Groups[1].Value
    if ($classValue -match '(?i)footer|template-part|is-nowrap|wp-block-image|wp-block-table') {
      $classValue
    }
  }
)
if ($layoutClassNames.Count -eq 0) {
  $layoutClassNames = @('(no matching layout classes in fetched HTML)')
}
Set-Content -LiteralPath (Join-Path $TempFolder 'layout-classes.log') -Value $layoutClassNames -Encoding UTF8
# Define the tokens
$startToken = '<pre class="wp-block-code CodeHidden"><code><!--STARTREPORT--></code></pre>'
$endToken = '<pre class="wp-block-code CodeHidden"><code><!--ENDREPORT--></code></pre>'

$LandScapeToken = "<p><!--LANDSCAPE--></p>"
$PortaitToken = "<p><!--PORTRAIT--></p>"

$tokens = @($startToken, $endToken, $LandScapeToken, $PortaitToken)

# Prepare a regex pattern for tokens
$tokenPattern = ($tokens -join "|").Replace("<!--", "<!--") # Ensures correct token format in regex

# Split the content by tokens (preserving tokens in results)
$splitContent = [regex]::Split($htmlContent, "($tokenPattern)")

 
$counter = 0
# Output the split content for inspection
Write-Host "Split Content (Separated by Tokens):" -ForegroundColor Green
$orientation = "portrait"
$splitContent | ForEach-Object {
 
  if ( ($_ -match $startToken)) {
   
     
    Rename-Item -Path "$TempFolder\section$($counter).html" -NewName "start.html"
    return 
  }
   
  if ( ($_ -match $endToken)) {
    return 

  }
     
  $counter++
  $_ | Out-File "$TempFolder\section$($counter).html" -Encoding utf8 
  
    
   
}
 

Rename-Item -Path "$TempFolder\section$($counter).html" -NewName "end.html"
$counter = 0 
foreach ($_ in $splitContent) {
  if ( ($_ -match $LandScapeToken)) {
    $orientation = 'Landscape'
    continue 
  }
   
  if ( ($_ -match $PortaitToken)) {
    $orientation = 'portrait' 
    continue 
  }
    
    
  if ( ($_ -match $startToken)) { 
    remove-Item  -Path "$TempFolder\section$($counter)_M.html"  -Force | Out-Null
    #remove-Item  -Path "C:\temp_pdf\section$($counter).html"  -Force | Out-Null
    remove-Item  -Path "$TempFolder\section$($counter).pdf"  -Force | Out-Null
    $counter = 0
    $orientation = 'portrait'
    continue 
  }
      
  if ( ($_ -match $endToken)) {
    $orientation = 'portrait' 
    break
  }
        
  $counter++ 

  Get-Content -Path "$TempFolder\start.html" | Out-File "$TempFolder\section$($counter)_M.html" -Encoding utf8   
  $_ | Out-File "$TempFolder\section$($counter)_M.html" -Append  -Encoding utf8  
  Get-Content -Path "$TempFolder\end.html" | Out-File "$TempFolder\section$($counter)_M.html" -Append -Encoding utf8 
  #remove-Item  -Path "C:\temp_pdf\section$($counter).html" -Force | Out-Null
  Write-Host "preprare : section$($counter)_M.html  $orientation"
  $output = "$TempFolder\section$($counter).pdf" 
  $sectionPrintCss = $printCss
  if ($counter -eq 1) {
    $sectionPrintCss += [Environment]::NewLine + $coverPrintCss
  }
  Set-Content -LiteralPath $printStylesheet -Value $sectionPrintCss -Encoding Ascii
  $layoutProbe = "(function(){var x=document.getElementsByTagName('*'),sizes={};for(var i=0;i<x.length;i++){if(x[i].children.length===0&&x[i].textContent.trim()){var key=x[i].tagName.toLowerCase()+' '+getComputedStyle(x[i]).fontSize;sizes[key]=(sizes[key]||0)+1;}}console.log('PDF_FONT_SIZES='+JSON.stringify(sizes));})();"
  if ($counter -eq 1) {
    $layoutProbe += "(function(){var x=document.getElementsByTagName('*'),matches=[],tail=[];for(var i=0;i<x.length;i++){var c=x[i].className;if(typeof c==='string'&&/footer|template-part|is-nowrap|wp-block-image|wp-block-table/i.test(c)){matches.push(x[i].tagName.toLowerCase()+'.'+c.replace(/\s+/g,'.'));}}for(var j=Math.max(0,x.length-80);j<x.length;j++){var c=x[j].className;if(typeof c!=='string')c='';tail.push(x[j].tagName.toLowerCase()+(c?'.'+c.replace(/\s+/g,'.'):''));}console.log('PDF_FOOTER_NODES='+JSON.stringify(matches));console.log('PDF_DOM_TAIL='+JSON.stringify(tail));var f=document.getElementsByClassName('Footer')[0];if(!f){console.log('PDF_FOOTER=missing');return;}var s=getComputedStyle(f),r=f.getBoundingClientRect(),g=f.getElementsByClassName('wp-block-group')[0],t=f.getElementsByTagName('td');console.log('PDF_FOOTER='+s.position+' bottom='+s.bottom+' rect='+Math.round(r.top)+':'+Math.round(r.bottom)+' row='+(g?getComputedStyle(g).display:'missing')+' border='+(t.length?getComputedStyle(t[0]).borderTopWidth:'missing')+' extraCell='+(t.length>1?getComputedStyle(t[1]).display:'missing'));})();"
  }
  $layoutProbeArgument = '"' + $layoutProbe + '"'
  $arggumen = '--page-size', 'A4', '--orientation', $orientation, '--background', '--minimum-font-size', '17', '--user-style-sheet', "`"$printStylesheet`"", '--outline', '--log-level', 'info', '--javascript-delay', '1000', '--debug-javascript', '--run-script', $layoutProbeArgument, '--encoding', 'UTF-8', '--print-media-type', '--keep-relative-links', '--image-dpi', '300', "`"$TempFolder\section$($counter)_M.html`"" , "`"$output`""
 
  Start-Process -NoNewWindow -Wait -FilePath $wkhtmltopdfPath   -ArgumentList $arggumen       -RedirectStandardError  (Join-Path $TempFolder "section$($counter)_M.log")
 
  ##Remove-Item -Path $TempFile

          
   
   
}
#$FILE = "C:\\temp_pdf\\section3.pdf"
#$arggumen = "-dBATCH", "-dQUIET", "-dNODISPLAY", "--permit-file-read=$FILE", "-c", "($FILE) (r) file runpdfbegin pdfpagecount = quit"
#
$outputPDFFinal = Join-Path $TempFolder 'merged.pdf'
$p = 0
$arggumenFile = "-f", "`"$pageCounterPS`""
Get-ChildItem -Path $TempFolder  -Filter   "*.pdf" | ForEach-Object {
  $arggumenFile += "-f"
  $arggumenFile += "`"$($_.FullName)`""
  $arggumen = "-q", "-dNODISPLAY", "-dNOSAFER"  , "--permit-file-read=`"$TempFolder`"", "-c ""($($_.name)) (r) file runpdfbegin pdfpagecount = quit"""   
   
  $proc = Start-Process -NoNewWindow -Wait -WorkingDirectory $TempFolder -FilePath $ghostscriptPath -PassThru  -ArgumentList $arggumen   -RedirectStandardOutput (Join-Path $TempFolder 'stdout.txt')
  $text = Get-Content "$($TempFolder)\stdout.txt"
  $p = $p + ([int]$text)
}

$arggumen = "-dBATCH", "-dNOPAUSE", "-sDEVICE=pdfwrite", "-dTotalPage=($($p))", "-dPDFSETTINGS=/prepress", "-sOutputFile=`"$outputPDFFinal`""
$arggumen += $arggumenFile
write-host $arggumen
Start-Process -NoNewWindow -Wait -FilePath $ghostscriptPath   -ArgumentList $arggumen    -RedirectStandardError  (Join-Path $TempFolder 'merging.log')
# Keep no earlier report from being overwritten: when a PDF of this name already exists (another
# report written the same day), keep the name and add a number, the way Windows does for a copy,
# "<name> (2).pdf", "<name> (3).pdf"... The launcher opens the name written here, in lastpdf.txt.
$finalPDF = $outputPDF
if (Test-Path -LiteralPath $finalPDF) {
  $pdfFolder = Split-Path -Path $outputPDF -Parent
  $pdfBase = [System.IO.Path]::GetFileNameWithoutExtension($outputPDF)
  $pdfExtension = [System.IO.Path]::GetExtension($outputPDF)
  $pdfNumber = 2
  do {
    $finalPDF = Join-Path $pdfFolder "$pdfBase ($pdfNumber)$pdfExtension"
    $pdfNumber++
  } while (Test-Path -LiteralPath $finalPDF)
}
Move-Item -LiteralPath $outputPDFFinal -Destination $finalPDF -Force
[System.IO.File]::WriteAllText((Join-Path $Log_folder 'lastpdf.txt'), $finalPDF)
Write-Host "PDF written: $finalPDF"
#Get-ChildItem -LiteralPath $TempFolder -File |
#  Where-Object { $_.Extension -ne '.log' } |
#  Remove-Item -Force

 