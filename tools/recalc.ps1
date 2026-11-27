param([string]$Dir)
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
try {
  Get-ChildItem -LiteralPath $Dir -Filter *.xlsx | ForEach-Object {
    $wb = $xl.Workbooks.Open($_.FullName)
    $xl.CalculateFull()
    $formulas = 0; $errs = @()
    foreach ($ws in $wb.Worksheets) {
      foreach ($c in $ws.UsedRange.Cells) {
        if ($c.HasFormula) {
          $formulas++
          if ($c.Text -match '^#') { $errs += "$($ws.Name)!$($c.Address(0,0))=$($c.Text) [$($c.Formula)]" }
        }
      }
    }
    "{0}: formulas={1} errors={2}" -f $_.Name, $formulas, $errs.Count
    $errs | Select-Object -First 10
    $wb.Save()
    $wb.Close($false)
  }
} finally { $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) }
