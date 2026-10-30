# Klickerns brygga till PowerPoint (30 oktober 2026). Läser ett kommando per rad på stdin och
# svarar med läget som en rad JSON, en rad per kommando:
#
#   lage    bara läget
#   fram    nästa klick i bildspelet, eller → till fönstret som ligger överst
#   bak     förra klicket, eller ←
#   start   startar bildspelet i den aktiva presentationen
#   svart   svart skärm av och på
#   bilder  läget och dessutom avsnitten och bildernas rubriker (innehall)
#   ga N    går till bild N; finns inget bildspel startar det där
#
# Bildspelet styrs genom PowerPoints automationsgränssnitt (COM), så PowerPoint behöver inte
# ligga överst. Utan bildspel blir fram och bak piltangenter, för en PDF eller en presentation i
# webbläsaren. KLICKERN_TORRT=1 skickar inga tangenter (proven).

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding $false
$utskrift = [Console]::Out
$torrt = $env:KLICKERN_TORRT -eq '1'

# Anteckningarna läses en gång per bild, nyckel sökväg|SlideID.
$anteckningar = @{}
$stadning = [Diagnostics.Stopwatch]::StartNew()
$formular = $false

function Program {
  # Startar aldrig PowerPoint: bara en som redan går.
  if ([Diagnostics.Process]::GetProcessesByName('POWERPNT').Length -eq 0) { return $null }
  try { return [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application') } catch {}
  # Office registrerar sig ibland först när fönstret har tappat fokus. PowerPoint har bara en
  # instans, så ett nytt objekt kopplas till den som går.
  try { return New-Object -ComObject PowerPoint.Application } catch { return $null }
}

function Tangent($tangent) {
  if (-not $torrt) {
    if (-not $script:formular) {
      Add-Type -AssemblyName System.Windows.Forms
      $script:formular = $true
    }
    [System.Windows.Forms.SendKeys]::SendWait($tangent)
  }
  return $tangent
}

function Anteckning($pres, $bild) {
  $nyckel = "$($pres.FullName)|$($bild.SlideID)"
  if ($anteckningar.ContainsKey($nyckel)) { return $anteckningar[$nyckel] }
  $text = ''
  try {
    foreach ($form in $bild.NotesPage.Shapes.Placeholders) {
      # 2 = ppPlaceholderBody, anteckningsfältet
      if ($form.PlaceholderFormat.Type -eq 2 -and $form.HasTextFrame) {
        $text = $form.TextFrame.TextRange.Text
        break
      }
    }
  } catch {}
  # PowerPoint skiljer stycken med \r och radbrytningar med \v.
  $text = ($text -replace "[\r\v]", "`n").Trim()
  $anteckningar[$nyckel] = $text
  return $text
}

function Presentation($app, $vy) {
  if ($vy) { return $app.SlideShowWindows.Item(1).Presentation }
  if ($app -and $app.Presentations.Count -gt 0) {
    try { return $app.ActivePresentation } catch { return $app.Presentations.Item(1) }
  }
  return $null
}

# Avsnitten och bildernas rubriker, till telefonens meny. Rubriken är rubrikfältet, eller annars
# första texten på bilden.
function Innehall($pres) {
  $bilder = New-Object System.Collections.ArrayList
  foreach ($bild in $pres.Slides) {
    $rubrik = ''
    try { if ($bild.Shapes.HasTitle) { $rubrik = $bild.Shapes.Title.TextFrame.TextRange.Text } } catch {}
    if (-not $rubrik.Trim()) {
      try {
        foreach ($form in $bild.Shapes) {
          if ($form.HasTextFrame -and $form.TextFrame.HasText) { $rubrik = $form.TextFrame.TextRange.Text; break }
        }
      } catch {}
    }
    $rubrik = ($rubrik -replace "[\r\v\n]+", ' ').Trim()
    if ($rubrik.Length -gt 90) { $rubrik = $rubrik.Substring(0, 89) + '…' }
    [void]$bilder.Add([ordered]@{ nr = [int]$bild.SlideIndex; rubrik = $rubrik; dold = ($bild.SlideShowTransition.Hidden -eq -1) })
  }
  $avsnitt = New-Object System.Collections.ArrayList
  try {
    $delar = $pres.SectionProperties
    for ($i = 1; $i -le $delar.Count; $i++) {
      [void]$avsnitt.Add([ordered]@{ namn = $delar.Name($i); forsta = [int]$delar.FirstSlide($i); antal = [int]$delar.SlidesCount($i) })
    }
  } catch {}
  return [ordered]@{ avsnitt = @($avsnitt); bilder = @($bilder) }
}

function Namn($pres) {
  return [IO.Path]::GetFileNameWithoutExtension($pres.Name)
}

function Lage($app) {
  $r = [ordered]@{ lage = 'ingen' }
  if ($null -eq $app) { return $r }
  if ($app.SlideShowWindows.Count -gt 0) {
    $fonster = $app.SlideShowWindows.Item(1)
    $vy = $fonster.View
    $pres = $fonster.Presentation
    $r.lage = 'bildspel'
    $r.titel = Namn $pres
    $r.antal = [int]$pres.Slides.Count
    # 1 visas, 2 pausad, 3 svart, 4 vit, 5 slut
    $r.status = [int]$vy.State
    if ($r.status -eq 5) {
      $r.slut = $true
    } else {
      $bild = $vy.Slide
      $r.bild = [int]$bild.SlideIndex
      $r.klick = [int]$vy.GetClickIndex()
      $r.klickAntal = [int]$vy.GetClickCount()
      $r.anteckningar = Anteckning $pres $bild
    }
    return $r
  }
  $anteckningar.Clear()
  if ($app.Presentations.Count -gt 0) {
    $r.lage = 'redigering'
    try { $r.titel = Namn $app.ActivePresentation } catch { $r.titel = Namn $app.Presentations.Item(1) }
  }
  return $r
}

while ($true) {
  $rad = [Console]::In.ReadLine()
  if ($null -eq $rad) { break }
  $kommando = $rad.Trim()
  $svar = 'ok'
  $tangent = $null
  $innehall = $null
  $app = $null
  $vy = $null
  try {
    $app = Program
    if ($app -and $app.SlideShowWindows.Count -gt 0) { $vy = $app.SlideShowWindows.Item(1).View }
    if ($kommando -match '^ga (\d{1,4})$') {
      $nr = [int]$Matches[1]
      $pres = Presentation $app $vy
      if (-not $pres) { $svar = 'ingen presentation' }
      elseif ($nr -lt 1 -or $nr -gt $pres.Slides.Count) { $svar = 'ingen sådan bild' }
      else {
        if (-not $vy) {
          $null = $pres.SlideShowSettings.Run()
          $vy = $app.SlideShowWindows.Item(1).View
        }
        $vy.GotoSlide($nr)
      }
    } elseif ($kommando -eq 'bilder') {
      $pres = Presentation $app $vy
      if ($pres) { $innehall = Innehall $pres } else { $svar = 'ingen presentation' }
    } else {
      switch ($kommando) {
        'lage' {}
        'fram' { if ($vy) { $vy.Next() } else { $tangent = Tangent '{RIGHT}' } }
        'bak' { if ($vy) { $vy.Previous() } else { $tangent = Tangent '{LEFT}' } }
        'start' {
          $pres = Presentation $app $vy
          if ($pres -and -not $vy) { $null = $pres.SlideShowSettings.Run() } else { $svar = 'inget att starta' }
        }
        'svart' { if ($vy) { $vy.State = $(if ($vy.State -eq 3 -or $vy.State -eq 4) { 1 } else { 3 }) } else { $svar = 'inget bildspel' } }
        default { $svar = 'okänt kommando' }
      }
    }
    $r = Lage $app
  } catch {
    # Till exempel en dialogruta i PowerPoint: anropen nekas tills den är stängd.
    $r = [ordered]@{ lage = 'upptagen'; fel = $_.Exception.Message }
  }
  $r.svar = $svar
  if ($tangent) { $r.tangent = $tangent }
  if ($innehall) { $r.innehall = $innehall }
  # Djupet räcker till innehall.bilder[i].nr.
  $utskrift.WriteLine(($r | ConvertTo-Json -Compress -Depth 5))
  $pres = $null
  # Släpp referenserna till PowerPoint ofta, så att den kan stängas som vanligt.
  $app = $null
  $vy = $null
  $fonster = $null
  if ($stadning.ElapsedMilliseconds -gt 2000) {
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    $stadning.Restart()
  }
}
