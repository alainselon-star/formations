param(
    [string]$Ids = "",
    [switch]$Visible
)
# ---------------------------------------------------------------------------
#  Generer_captures.ps1 - Test de positionnement Excel (Atelier Martin)
#  Ouvre Decors_captures_test_Excel.xlsx dans Excel, photographie chaque plage
#  listee dans l'onglet Index (lignes "Script") et enregistre une image PNG
#  par question dans le sous-dossier "captures" (ex. captures\Q8.png).
#  Toutes les images ont le meme format 3:2 (1800 x 1200 px), fond blanc.
#  Usage : double-clic sur Generer_captures.bat (toutes les captures)
#          ou Tester_3_captures.bat (Q8, Q19, R4 seulement).
# ---------------------------------------------------------------------------
$ErrorActionPreference = "Stop"
$base   = Split-Path -Parent $MyInvocation.MyCommand.Path
$source = Join-Path $base "Decors_captures_test_Excel.xlsx"
$sortie = Join-Path $base "captures"
New-Item -ItemType Directory -Force -Path $sortie | Out-Null

$filtre = @()
if ($Ids -ne "") { $filtre = $Ids.Split(",") | ForEach-Object { $_.Trim() } }

# Format de l'image finale (en points ; export = points x 4/3 en pixels)
$LARG = 1350
$HAUT = 900
$MARGE = 20

Write-Host ""
Write-Host "Captures du test de positionnement Excel" -ForegroundColor Cyan
Write-Host "Classeur : $source"
Write-Host "Dossier  : $sortie"
Write-Host ""

if (-not (Test-Path $source)) { Write-Host "Classeur des decors introuvable." -ForegroundColor Red; exit 1 }

$xl = New-Object -ComObject Excel.Application
$xl.Visible = [bool]$Visible
$xl.DisplayAlerts = $false
$wb = $xl.Workbooks.Open($source, 0, $true)   # lecture seule : l'original n'est jamais modifie

$index = $wb.Worksheets.Item("Index")
$tmp = $wb.Worksheets.Add()
$tmp.Name = "Montage"
$ok = 0; $ko = @(); $sansEntetes = @()
$ligne = 2
while ($index.Cells.Item($ligne, 1).Text -ne "") {
    $id      = $index.Cells.Item($ligne, 1).Text
    $feuille = $index.Cells.Item($ligne, 2).Text
    $plage   = $index.Cells.Item($ligne, 3).Text
    $type    = $index.Cells.Item($ligne, 4).Text
    $quadr   = $index.Cells.Item($ligne, 5).Text
    $prod    = $index.Cells.Item($ligne, 6).Text
    $ligne++

    if ($prod -ne "Script") { continue }
    if ($filtre.Count -gt 0 -and -not ($filtre -contains $id)) { continue }

    Write-Host ("  {0,-4} {1,-10} {2,-8} ... " -f $id, $type, $plage) -NoNewline
    try {
        $ws = $wb.Worksheets.Item($feuille)
        $null = $ws.Activate()

        # Tableau croise dynamique cree a la volee
        if ($type -eq "TCD") {
            $src  = $index.Cells.Item($ligne - 1, 7).Text
            $dest = $index.Cells.Item($ligne - 1, 8).Text
            $lig  = $index.Cells.Item($ligne - 1, 9).Text
            $val  = $index.Cells.Item($ligne - 1, 10).Text
            $pc = $wb.PivotCaches().Create(1, $ws.Range($src))
            $pt = $pc.CreatePivotTable($ws.Range($dest), "TCD_" + $id)
            $pt.PivotFields($lig).Orientation = 1
            $df = $pt.AddDataField($pt.PivotFields($val), "Somme de " + $val, -4157)
            $df.NumberFormat = "# ##0 €"
        }

        # En-tetes de lignes et colonnes (A, B, C / 1, 2, 3) et quadrillage
        $avecEntetes = $true
        try {
            $xl.PrintCommunication = $false
            $ws.PageSetup.PrintHeadings  = $true
            $ws.PageSetup.PrintGridlines = ($quadr -eq "Oui")
            $xl.PrintCommunication = $true
        } catch { $avecEntetes = $false; try { $xl.PrintCommunication = $true } catch {} }

        # Copie de la plage en image vectorielle (fonction locale)
        function Copier-Plage {
            for ($t = 1; $t -le 3; $t++) {
                try {
                    if ($script:avecEntetes) { $null = $ws.Range($plage).CopyPicture(2, -4147) }   # xlPrinter : avec en-tetes
                    else                     { $null = $ws.Range($plage).CopyPicture(1, -4147) }   # xlScreen  : sans en-tetes
                    return
                } catch {
                    if ($script:avecEntetes -and $t -eq 2) { $script:avecEntetes = $false }
                    Start-Sleep -Milliseconds 400
                }
            }
            throw "copie de la plage impossible"
        }
        $script:avecEntetes = $avecEntetes

        # Le graphique de montage est cree sur une feuille a part, pour ne jamais masquer la plage photographiee
        $co = $tmp.ChartObjects().Add(0, 0, $LARG, $HAUT)
        $ch = $co.Chart
        $ch.ChartArea.Format.Fill.Visible = -1
        $ch.ChartArea.Format.Fill.ForeColor.RGB = 16777215
        $ch.ChartArea.Format.Line.Visible = 0

        # 1er passage : mesure de la taille naturelle de la plage
        Copier-Plage
        $null = $tmp.Activate()
        $null = $co.Activate()
        Start-Sleep -Milliseconds 200
        $null = $ch.Paste()
        $img = $ch.Shapes.Item($ch.Shapes.Count)
        $ratio = [Math]::Min(($LARG - 2 * $MARGE) / $img.Width, ($HAUT - 2 * $MARGE) / $img.Height)
        $null = $img.Delete()

        # 2e passage : agrandissement par le zoom d'impression (texte net, sans deformation)
        if ($script:avecEntetes) {
            $zoom = [Math]::Max(10, [Math]::Min(400, [Math]::Floor(100 * $ratio)))
            try {
                $xl.PrintCommunication = $false
                $ws.PageSetup.Zoom = $zoom
                $xl.PrintCommunication = $true
            } catch { try { $xl.PrintCommunication = $true } catch {} }
        }
        Copier-Plage
        $null = $tmp.Activate()
        $null = $co.Activate()
        Start-Sleep -Milliseconds 200
        $null = $ch.Paste()
        $img = $ch.Shapes.Item($ch.Shapes.Count)
        $img.LockAspectRatio = 0
        $w0 = $img.Width; $h0 = $img.Height
        $r2 = [Math]::Min(($LARG - 2 * $MARGE) / $w0, ($HAUT - 2 * $MARGE) / $h0)
        $img.Width  = $w0 * $r2
        $img.Height = $h0 * $r2
        $img.Left   = ($LARG - $img.Width) / 2
        $img.Top    = ($HAUT - $img.Height) / 2
        if (-not $script:avecEntetes) { $sansEntetes += $id }
        $fichier = Join-Path $sortie ($id + ".png")
        if (Test-Path $fichier) { Remove-Item $fichier -Force }
        $null = $ch.Export($fichier, "PNG")
        $null = $co.Delete()

        $ok++
        Write-Host "OK" -ForegroundColor Green
    } catch {
        $ko += "$id : $($_.Exception.Message)"
        Write-Host "ERREUR" -ForegroundColor Red
    }
}

$wb.Close($false)
$xl.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null

Write-Host ""
Write-Host "$ok capture(s) enregistree(s) dans $sortie" -ForegroundColor Cyan
if ($sansEntetes.Count -gt 0) {
    Write-Host ("Sans lettres de colonnes (aucune imprimante disponible) : " + ($sansEntetes -join ", ")) -ForegroundColor Yellow
}
if ($ko.Count -gt 0) {
    Write-Host "Problemes :" -ForegroundColor Red
    $ko | ForEach-Object { Write-Host "  $_" -ForegroundColor Red }
}
Start-Process explorer.exe $sortie
