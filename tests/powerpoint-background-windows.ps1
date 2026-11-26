# Run on a Windows desktop with PowerPoint installed. Only the temporary fixture is edited.
$ErrorActionPreference = 'Stop'
$folder = Join-Path ([IO.Path]::GetTempPath()) ('triobridge-ppt-' + [guid]::NewGuid())
[void][IO.Directory]::CreateDirectory($folder)
$helper = Join-Path $folder 'Office.exe'
$deckPath = Join-Path $folder 'Background-20.pptx'
$deck = $null
$office = $null
try {
    $compiler = Join-Path $env:SystemRoot 'Microsoft.NET/Framework/v4.0.30319/csc.exe'
    & $compiler /nologo /utf8output /codepage:65001 /target:exe "/out:$helper" /reference:Microsoft.CSharp.dll /reference:System.Web.Extensions.dll /reference:System.Core.dll (Join-Path $PSScriptRoot '../offline/windows/Office.cs')
    if ($LASTEXITCODE) { throw 'Could not compile the native PowerPoint bridge.' }
    $office = New-Object -ComObject PowerPoint.Application
    $deck = $office.Presentations.Add(-1)
    for ($i = 1; $i -le 20; $i++) {
        $slide = $deck.Slides.Add($i, 12)
        $slide.Shapes.AddTextbox(1, 30, 20, 360, 60).TextFrame.TextRange.Text = "Quarter $i"
        if ($i -eq 5) { $null = $slide.Shapes.AddChart2(-1, 51, 30, 90, 360, 220) }
    }
    $deck.SaveAs($deckPath, 24)
    function Invoke-Bridge($payload) {
        # Explicit UTF-8 preserves Chinese labels across the PowerShell 5.1 stdin pipe.
        $OutputEncoding = [Text.UTF8Encoding]::new($false)
        $json = $payload | ConvertTo-Json -Depth 30 -Compress
        $result = ($json | & $helper) | ConvertFrom-Json
        if (!$result.ok) { throw $result.message }
        return $result.value
    }
    $before = Invoke-Bridge @{ host = 'powerpoint'; op = 'presentation'; scope = 'all' }
    if ($before.guard.name -ne $deckPath -or $before.snapshot.slides.Count -ne 20) { throw 'Wrong test deck.' }
    $actions = foreach ($slide in $before.snapshot.slides) {
        $color = if ($slide.index -eq 4) { '#17365D' } else { '#F5F1E8' }
        @{ kind = 'background'; slideId = $slide.id; fill = @{ kind = 'solid'; color = $color } }
    }
    $null = Invoke-Bridge @{ host = 'powerpoint'; op = 'edit'; guard = $before.guard; snapshot = $before.snapshot; proposal = @{ summary = '20-slide acceptance'; changes = @(); actions = @($actions) } }
    $deck.Save()
    $deck.Close()
    $deck = $null
    $deck = $office.Presentations.Open($deckPath, 0, 0, -1)
    $after = Invoke-Bridge @{ host = 'powerpoint'; op = 'presentation'; scope = 'all' }
    if ($after.guard.name -ne $deckPath -or $after.snapshot.slides.Count -ne 20) { throw 'Wrong reopened deck.' }
    foreach ($slide in $after.snapshot.slides) {
        $expected = if ($slide.index -eq 4) { '#17365D' } else { '#F5F1E8' }
        if ($slide.background.kind -ne 'solid' -or $slide.background.color -ne $expected) { throw "Unexpected background on slide $($slide.index + 1)." }
    }
    if (!$deck.Slides.Item(5).Shapes.Item(2).HasChart) { throw 'The editable chart was lost.' }
    Write-Output 'Passed: 19 cream slides, dark-blue slide 5 and the original editable chart after reopen.'
} finally {
    if ($null -ne $deck) { $deck.Saved = -1; $deck.Close() }
    if ($null -ne $office) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($office) }
    if ([IO.Directory]::Exists($folder)) { [IO.Directory]::Delete($folder, $true) }
}
