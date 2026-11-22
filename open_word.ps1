param([Parameter(Mandatory=$true)][string]$DocumentPath)

$ErrorActionPreference = 'Stop'
try {
    $word = [Runtime.InteropServices.Marshal]::GetActiveObject('Word.Application')
} catch {
    $word = New-Object -ComObject Word.Application
}

$word.Visible = $true
$word.Options.CheckSpellingAsYouType = $true
$word.Options.CheckGrammarAsYouType = $true
$word.Options.CheckGrammarWithSpelling = $true

$target = [IO.Path]::GetFullPath($DocumentPath)
foreach ($document in $word.Documents) {
    if ($document.FullName -eq $target) {
        $document.Activate()
        exit 0
    }
}

$null = $word.Documents.Open($target)
$word.Activate()
