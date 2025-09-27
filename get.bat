@echo off
REM ============================================================
REM  Typhoon latest downloader + KMZ->GeoJSON converter
REM  - Download latest W-C0034-005.json  (overwrite)
REM  - Download latest W-C0034-002.kmz   (overwrite)
REM  - Convert KMZ (KML inside) to W-C0034-002.json (GeoJSON, UTF-8)
REM  Requirements: Windows 10/11 (curl, PowerShell)
REM ============================================================

chcp 65001 >nul

REM === Config ======================================================
set "DATASET_JSON=W-C0034-005"
set "DATASET_KMZ=W-C0034-002"
REM 建議只填純金鑰：rdec-key-123-45678-011121314
set "AUTH_KEY=rdec-key-123-45678-011121314"

set "BASE_FILEAPI=https://opendata.cwa.gov.tw/fileapi/v1/opendataapi"

REM Handle accidental inclusion of format in AUTH_KEY
set "FMT_JSON=&format=JSON"
set "FMT_KMZ=&format=KMZ"
echo %AUTH_KEY% | findstr /I "format=" >nul
if %errorlevel%==0 (
  set "FMT_JSON="
  set "FMT_KMZ="
)

set "URL_JSON=%BASE_FILEAPI%/%DATASET_JSON%?Authorization=%AUTH_KEY%%FMT_JSON%"
set "URL_KMZ=%BASE_FILEAPI%/%DATASET_KMZ%?Authorization=%AUTH_KEY%%FMT_KMZ%"

set "OUT_JSON=%DATASET_JSON%.json"
set "OUT_KMZ=%DATASET_KMZ%.kmz"
set "OUT_KMZ_JSON=%DATASET_KMZ%.json"

echo [1/4] Downloading %OUT_JSON% ...
curl -sSL --retry 3 --retry-delay 2 --fail "%URL_JSON%" -o "%OUT_JSON%"
if errorlevel 1 (
  echo [ERROR] Failed to download %OUT_JSON%. Check network/key.
  echo URL: %URL_JSON%
  pause
  exit /b 1
)

echo [2/4] Downloading %OUT_KMZ% ...
curl -sSL --retry 3 --retry-delay 2 --fail "%URL_KMZ%" -o "%OUT_KMZ%"
if errorlevel 1 (
  echo [ERROR] Failed to download %OUT_KMZ%. Check network/key.
  echo URL: %URL_KMZ%
  pause
  exit /b 1
)

echo [3/4] Extracting KMZ and converting KML -> GeoJSON ...

REM --- PowerShell: KMZ -> KML -> GeoJSON (FeatureCollection) ------
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$ErrorActionPreference='Stop';" ^
  "$kmz = Join-Path (Get-Location) '%OUT_KMZ%';" ^
  "$tmp = Join-Path (Get-Location) 'tmp_kmz_extract';" ^
  "$zip = Join-Path (Get-Location) 'tmp_kmz.zip';" ^
  "if (Test-Path $tmp) { Remove-Item $tmp -Recurse -Force };" ^
  "if (Test-Path $zip) { Remove-Item $zip -Force };" ^
  "Copy-Item -Path $kmz -Destination $zip -Force;" ^
  "Expand-Archive -Path $zip -DestinationPath $tmp -Force;" ^
  "$kmlFile = Get-ChildItem -Path $tmp -Recurse -Filter *.kml | Select-Object -First 1;" ^
  "if (-not $kmlFile) { throw 'No KML found inside KMZ.' };" ^
  "$xml = New-Object System.Xml.XmlDocument;" ^
  "$xml.XmlResolver = $null; $xml.Load($kmlFile.FullName);" ^
  "$nsm = New-Object System.Xml.XmlNamespaceManager($xml.NameTable);" ^
  "$nsm.AddNamespace('kml','http://www.opengis.net/kml/2.2');" ^
  "$nsm.AddNamespace('gx','http://www.google.com/kml/ext/2.2');" ^
  "$placemarks = $xml.SelectNodes('//kml:Placemark',$nsm);" ^
  "$features = @();" ^
  "function Parse-Coords([string]$coordText) {" ^
  "  $coordText = ($coordText -replace '\s+', ' ').Trim();" ^
  "  $pairs = $coordText -split '\s+' | Where-Object { $_ -ne '' };" ^
  "  $pairs | ForEach-Object {" ^
  "    $p = $_.Split(',');" ^
  "    [double]$lon = $p[0]; [double]$lat = $p[1];" ^
  "    if ($p.Count -ge 3) { [double]$alt = $p[2] } else { $alt = $null }" ^
  "    if ($alt -ne $null) { ,@($lon,$lat,$alt) } else { ,@($lon,$lat) }" ^
  "  }" ^
  "}" ^
  "foreach ($pm in $placemarks) {" ^
  "  $nameNode = $pm.SelectSingleNode('kml:name',$nsm);" ^
  "  $descNode = $pm.SelectSingleNode('kml:description',$nsm);" ^
  "  $name = if ($nameNode) { $nameNode.InnerText } else { $null };" ^
  "  $desc = if ($descNode) { $descNode.InnerText } else { $null };" ^
  "  $geom = $null; $type = $null;" ^
  "  $pt = $pm.SelectSingleNode('.//kml:Point/kml:coordinates',$nsm);" ^
  "  if ($pt) {" ^
  "    $coords = Parse-Coords $pt.InnerText;" ^
  "    if ($coords.Count -ge 1) { $type='Point'; $geom=@{ type=$type; coordinates=$coords[0] } }" ^
  "  } else {" ^
  "    $ls = $pm.SelectSingleNode('.//kml:LineString/kml:coordinates',$nsm);" ^
  "    if ($ls) {" ^
  "      $coords = Parse-Coords $ls.InnerText;" ^
  "      if ($coords.Count -ge 2) { $type='LineString'; $geom=@{ type=$type; coordinates=$coords } }" ^
  "    } else {" ^
  "      $poly = $pm.SelectSingleNode('.//kml:Polygon/kml:outerBoundaryIs/kml:LinearRing/kml:coordinates',$nsm);" ^
  "      if ($poly) {" ^
  "        $coords = Parse-Coords $poly.InnerText;" ^
  "        if ($coords.Count -ge 4) { $type='Polygon'; $geom=@{ type=$type; coordinates=@(,$coords) } }" ^
  "      } else {" ^
  "        $gxTrack = $pm.SelectNodes('.//gx:Track/gx:coord',$nsm);" ^
  "        if ($gxTrack -and $gxTrack.Count -gt 0) {" ^
  "          $coordArr = @();" ^
  "          foreach ($c in $gxTrack) {" ^
  "            $parts = ($c.InnerText -replace '\s+',' ').Trim().Split(' ');" ^
  "            if ($parts.Count -ge 2) {" ^
  "              [double]$lon=$parts[0]; [double]$lat=$parts[1];" ^
  "              if ($parts.Count -ge 3) { [double]$alt=$parts[2] } else { $alt=$null }" ^
  "              if ($alt -ne $null) { $coordArr += ,@($lon,$lat,$alt) } else { $coordArr += ,@($lon,$lat) }" ^
  "            }" ^
  "          }" ^
  "          if ($coordArr.Count -ge 2) { $type='LineString'; $geom=@{ type=$type; coordinates=$coordArr } }" ^
  "        }" ^
  "      }" ^
  "    }" ^
  "  }" ^
  "  if ($geom -ne $null) {" ^
  "    $feature = @{ type='Feature'; geometry=$geom; properties=@{ name=$name; description=$desc } };" ^
  "    $features += ,$feature;" ^
  "  }" ^
  "}" ^
  "$fc = @{ type='FeatureCollection'; features=$features };" ^
  "($fc | ConvertTo-Json -Depth 100) | Set-Content -Path '%OUT_KMZ_JSON%' -Encoding UTF8;" ^
  "Remove-Item $tmp -Recurse -Force -ErrorAction SilentlyContinue;" ^
  "Remove-Item $zip -Force -ErrorAction SilentlyContinue;"

if errorlevel 1 (
  echo [ERROR] KMZ -> GeoJSON failed.
  pause
  exit /b 1
)

for %%F in ("%OUT_JSON%") do set "SIZE_JSON=%%~zF"
for %%F in ("%OUT_KMZ%") do set "SIZE_KMZ=%%~zF"
for %%F in ("%OUT_KMZ_JSON%") do set "SIZE_KMZ_JSON=%%~zF"
echo [4/4] Done:
echo   %OUT_JSON%      (%SIZE_JSON% bytes)
echo   %OUT_KMZ%       (%SIZE_KMZ% bytes)
echo   %OUT_KMZ_JSON%  (%SIZE_KMZ_JSON% bytes, GeoJSON)
exit /b 0
