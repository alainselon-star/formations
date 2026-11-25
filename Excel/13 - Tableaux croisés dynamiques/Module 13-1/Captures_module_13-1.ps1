param([switch]$Visible)
# Excel doit etre visible pour le rendu "ecran" : le lanceur passe -Visible
# ---------------------------------------------------------------------------
#  Captures_module_13-1.ps1 - Module 13.1 "La reunion de 10 h" (TCD)
#  Ouvre Excel_13_TCD_factures_2026.xlsx EN LECTURE SEULE (dossier parent),
#  fait les manipulations en memoire (filtres, formule, TCD) sans jamais
#  enregistrer, et exporte une image PNG par ecran dans le sous-dossier
#  "captures" (C01.png ... C11.png), rendu identique a l'ecran, fond blanc.
# ---------------------------------------------------------------------------
$ErrorActionPreference = "Stop"
$base   = Split-Path -Parent $MyInvocation.MyCommand.Path
$source = Join-Path (Split-Path -Parent $base) "Excel_13_TCD_factures_2026.xlsx"
$sortie = Join-Path $base "captures"
New-Item -ItemType Directory -Force -Path $sortie | Out-Null
Get-ChildItem $sortie -Filter "*_ecran.png" | Remove-Item -Force   # anciennes versions de test
$LARG = 1350; $HAUT = 900; $MARGE = 20

Write-Host ""
Write-Host "Captures du module 13.1 - La reunion de 10 h" -ForegroundColor Cyan
Write-Host "Classeur : $source"
Write-Host "Dossier  : $sortie"
Write-Host ""
if (-not (Test-Path $source)) { Write-Host "Classeur introuvable." -ForegroundColor Red; exit 1 }

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $true   # obligatoire pour la copie "ecran"
$xl.DisplayAlerts = $false
$wb = $xl.Workbooks.Open($source, 0, $true)   # lecture seule
$tmp = $wb.Worksheets.Add()
$tmp.Name = "Montage"
$xl.WindowState = -4137   # plein ecran pour un rendu complet
$script:ok = 0; $script:ko = @()

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

function Capture($id, $ws, $rg, $quadr) {
    # Photo "ecran" en image bitmap, recuperee directement dans le presse-papiers :
    # aucune mise a l'echelle, donc aucune deformation possible. Marge blanche de 20 px.
    Write-Host ("  {0,-4} {1,-18} {2,-12} ... " -f $id, $ws.Name, $rg) -NoNewline
    try {
        $null = $ws.Activate()
        $xl.ActiveWindow.DisplayGridlines = [bool]$quadr
        [System.Windows.Forms.Clipboard]::Clear()
        $bmp = $null
        for ($t = 1; $t -le 5 -and $bmp -eq $null; $t++) {
            $null = $ws.Range($rg).CopyPicture(1, 2)          # xlScreen, xlBitmap
            Start-Sleep -Milliseconds 300
            $bmp = [System.Windows.Forms.Clipboard]::GetImage()
        }
        if ($bmp -eq $null) { throw "image absente du presse-papiers" }
        $m = 20
        $out = New-Object System.Drawing.Bitmap ($bmp.Width + 2*$m), ($bmp.Height + 2*$m)
        $g = [System.Drawing.Graphics]::FromImage($out)
        $g.Clear([System.Drawing.Color]::White)
        $g.DrawImageUnscaled($bmp, $m, $m)
        $g.Dispose()
        $f = Join-Path $sortie ($id + ".png")
        if (Test-Path $f) { Remove-Item $f -Force }
        $out.Save($f, [System.Drawing.Imaging.ImageFormat]::Png)
        $script:ok++
        Write-Host ("OK  ({0} x {1} px)" -f $out.Width, $out.Height) -ForegroundColor Green
        $out.Dispose(); $bmp.Dispose()
    } catch {
        $script:ko += "$id : $($_.Exception.Message)"
        Write-Host "ERREUR" -ForegroundColor Red
    }
}

$fact  = $wb.Worksheets.Item("Factures")
$essai = $wb.Worksheets.Item("Essai à la main")
$lo    = $fact.ListObjects.Item("Factures")
$sep   = $xl.International(4)            # separateur de milliers de Windows
$fmt   = "#" + $sep + "##0 €"

# --- Sequence 1 : le fichier et la grille vide ---------------------------
Capture "C01" $fact  "A1:H16" $true      # ecran 1.3 : les premieres factures
Capture "C02" $essai "A4:N10" $true      # ecran 1.4 : la grille 5 x 12 a remplir

# --- Sequence 2 : Solution 1 ---------------------------------------------
try {
    $null = $lo.Range.AutoFilter(5, "Peinture")
    Capture "C03" $fact "A1:H80" $true   # ecran 2.1 : filtre Peinture
    $null = $lo.Range.AutoFilter(1, ">=46023", 1, "<=46053")   # 01/01/2026 -> 31/01/2026
    Capture "C04" $fact "A1:H40" $true   # ecran 2.1 : Peinture + janvier (5 factures, 14 920 EUR)
    try { $lo.AutoFilter.ShowAllData() } catch {}
} catch { $script:ko += "Filtres : $($_.Exception.Message)" }

try {
    $frm = '=SUMIFS(Factures[Montant HT],Factures[Type de travaux],"Peinture",Factures[Date],">="&DATE(2026,1,1),Factures[Date],"<="&DATE(2026,1,31))'
    $essai.Range("B8").Formula = $frm
    $essai.Range("B8").NumberFormatLocal = $fmt
    $essai.Range("A12").Value2 = "B8 :  =SOMME.SI.ENS(Factures[Montant HT];Factures[Type de travaux];""Peinture"";Factures[Date];"">=01/01/2026"";Factures[Date];""<=31/01/2026"")"
    $essai.Range("A12").Font.Name = "Consolas"
    $essai.Range("A12").Font.Size = 10
    Capture "C05" $essai "A4:N12" $true  # ecran 2.2 : une case juste avec SOMME.SI.ENS
    $essai.Range("C8").Formula = $frm
    $essai.Range("C8").NumberFormatLocal = $fmt
    $essai.Range("B8:C8").Interior.Color = 13431551   # jaune pale
    $essai.Range("A13").Value2 = "C8 (copiée en février) : même résultat, 14 920 €, au lieu de 31 260 €"
    $essai.Range("A13").Font.Color = 255
    Capture "C06" $essai "A4:N13" $true  # ecran 2.3 : le piege de la copie
} catch { $script:ko += "Formules : $($_.Exception.Message)" }

# --- Sequence 3 : Solution 2.0, le TCD -----------------------------------
try {
    $tcdWs = $wb.Worksheets.Add()
    $tcdWs.Name = "TCD"
    $pc = $wb.PivotCaches().Create(1, "Factures")
    $pt = $pc.CreatePivotTable($tcdWs.Range("A3"), "TCD_Sophie")

    $pt.PivotFields("Type de travaux").Orientation = 1          # Lignes
    Capture "C07" $tcdWs $pt.TableRange1.Address() $true         # ecran 3.5 : les types en lignes

    $df = $pt.AddDataField($pt.PivotFields("Montant HT"), "Somme de Montant HT", -4157)
    function Format-TCD { try { $df.NumberFormat = $fmt } catch { $script:ko += "Format TCD : $($_.Exception.Message)" } }
    Format-TCD
    Capture "C08" $tcdWs $pt.TableRange1.Address() $true         # ecran 3.5 : + les montants (1 176 350 EUR)

    $pt.PivotFields("Date").Orientation = 2                      # Colonnes
    try {
        $cel = $pt.PivotFields("Date").DataRange.Cells.Item(1)
        $per = [object[]]@($false, $false, $false, $false, $true, $false, $false)
        $null = $cel.Group($true, $true, [System.Reflection.Missing]::Value, $per)
    } catch {}
    Format-TCD
    $tcdWs.Columns.AutoFit() | Out-Null
    Capture "C09" $tcdWs $pt.TableRange1.Address() $true         # ecran 3.6 : le TCD complet 5 x 12

    $tr = $pt.TableRange1
    $coin = $tcdWs.Range($tr.Cells.Item(1, 1), $tr.Cells.Item($tr.Rows.Count, 4)).Address()
    Capture "C10" $tcdWs $coin $true                             # ecran 3.1 : un coin du tableau croise

    try { $pt.PivotFields("Type de travaux").Orientation = 0 } catch {}
    foreach ($n in @("Date", "Mois", "Trimestres", "Années")) { try { $pt.PivotFields($n).Orientation = 0 } catch {} }
    $pt.PivotFields("Ville").Orientation = 1
    try { $pt.PivotFields("Ville").AutoSort(2, "Somme de Montant HT") } catch {}
    Format-TCD
    $tcdWs.Columns.AutoFit() | Out-Null
    Capture "C11" $tcdWs $pt.TableRange1.Address() $true         # ecran 3.2 / 3.6 : "Et par ville ?"
} catch { $script:ko += "TCD : $($_.Exception.Message)" }

$wb.Close($false)   # rien n'est enregistre
$xl.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null

Write-Host ""
Write-Host "$($script:ok) capture(s) enregistree(s) dans $sortie" -ForegroundColor Cyan
if ($script:ko.Count -gt 0) {
    Write-Host "Problemes :" -ForegroundColor Red
    $script:ko | ForEach-Object { Write-Host "  $_" -ForegroundColor Red }
}
Write-Host ""
Write-Host "A faire a la main (3 captures) : voir A_capturer_a_la_main.txt" -ForegroundColor Yellow
Start-Process explorer.exe $sortie
