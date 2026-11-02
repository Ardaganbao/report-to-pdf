
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
Move-Item $outputPDFFinal  $outputPDF -Force
#Get-ChildItem -LiteralPath $TempFolder -File |
#  Where-Object { $_.Extension -ne '.log' } |
#  Remove-Item -Force

 