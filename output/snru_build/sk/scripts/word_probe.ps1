param($doc,$out)
# Writes one line per body paragraph (table-cell paragraphs are skipped):
#   <displayed page number>\t<physical page>\t<first 60 chars>
# Used by finalize_thesis.py / check_thesis.py to learn Word's real pagination.
$w=New-Object -ComObject Word.Application; $w.Visible=$false; $w.DisplayAlerts=0
try {
  $d=$w.Documents.Open($doc,$false,$true)
  $d.Repaginate()
  $lines=New-Object System.Collections.Generic.List[string]
  foreach($p in $d.Paragraphs){
    $r=$p.Range
    if($r.Information(12)){ continue }
    $t=$r.Text
    if($t.Length -gt 60){$t=$t.Substring(0,60)}
    $t=$t -replace "`t",' ' -replace "`r",'' -replace "`n",' ' -replace [char]11,' ' -replace [char]12,''
    $s=$d.Range($r.Start,$r.Start)
    $lines.Add([string]$s.Information(1)+"`t"+[string]$s.Information(3)+"`t"+$t)
  }
  $d.Close($false)
} finally { $w.Quit() }
[IO.File]::WriteAllLines($out,$lines,[Text.UTF8Encoding]::new($false))
