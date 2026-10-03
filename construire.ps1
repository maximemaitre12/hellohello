# Fabrique les deux applications Adresses Outlook, prêtes à double-cliquer :
#
#   dist\Adresses Outlook.exe           Windows : un seul fichier, qui contient
#                                       adresses-outlook.ps1 et le fait tourner
#                                       en lui-même (pas de console, sa propre
#                                       icône dans la barre des tâches).
#   mac\natif\AppIcon.icns             Mac : l'icône. L'app Mac elle-même se
#                                       fabrique sur un Mac de GitHub (voir
#                                       .github\workflows\mac.yml).
#
#   powershell -NoProfile -ExecutionPolicy Bypass -File outils\construire.ps1

param([string]$Out = (Join-Path $PSScriptRoot "dist"))
$ErrorActionPreference = "Stop"
New-Item -ItemType Directory -Force -Path $Out | Out-Null
Add-Type -AssemblyName System.Drawing

# ------------------------------------------------------------------ icône
# Le « @ » blanc sur carré bleu Aether, comme dans l'app.

function Icone-Png([int]$size) {
  $bmp = New-Object System.Drawing.Bitmap $size, $size
  $g = [System.Drawing.Graphics]::FromImage($bmp)
  $g.SmoothingMode = "AntiAlias"; $g.TextRenderingHint = "AntiAliasGridFit"; $g.InterpolationMode = "HighQualityBicubic"
  $m = [Math]::Round($size * 0.08); $w = $size - 2 * $m; $r = [Math]::Max(2, [Math]::Round($w * 0.24))
  $path = New-Object System.Drawing.Drawing2D.GraphicsPath
  $path.AddArc($m, $m, $r, $r, 180, 90); $path.AddArc($m + $w - $r, $m, $r, $r, 270, 90)
  $path.AddArc($m + $w - $r, $m + $w - $r, $r, $r, 0, 90); $path.AddArc($m, $m + $w - $r, $r, $r, 90, 90); $path.CloseFigure()
  $g.FillPath((New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(30, 77, 140))), $path)
  $font = New-Object System.Drawing.Font("Segoe UI", [float]($size * 0.5), [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Pixel)
  $fmt = New-Object System.Drawing.StringFormat; $fmt.Alignment = "Center"; $fmt.LineAlignment = "Center"
  $g.DrawString("@", $font, [System.Drawing.Brushes]::White, (New-Object System.Drawing.RectangleF(0, [float](-$size * 0.04), $size, $size)), $fmt)
  $g.Dispose()
  $ms = New-Object System.IO.MemoryStream
  $bmp.Save($ms, [System.Drawing.Imaging.ImageFormat]::Png); $bmp.Dispose()
  return ,$ms.ToArray()
}

function Write-BE([System.IO.BinaryWriter]$w, [uint32]$v) { $w.Write([byte[]]@((($v -shr 24) -band 255), (($v -shr 16) -band 255), (($v -shr 8) -band 255), ($v -band 255))) }

# Windows : un .ico fait d'images PNG.
$icoPath = Join-Path $Out "icone.ico"
$sizes = 16, 24, 32, 48, 64, 128, 256
$pngs = @($sizes | ForEach-Object { ,(Icone-Png $_) })
$fs = [System.IO.File]::Create($icoPath); $w = New-Object System.IO.BinaryWriter $fs
$w.Write([uint16]0); $w.Write([uint16]1); $w.Write([uint16]$sizes.Count)
$offset = 6 + 16 * $sizes.Count
for ($i = 0; $i -lt $sizes.Count; $i++) {
  $s = $sizes[$i]; $b = if ($s -ge 256) { 0 } else { $s }
  $w.Write([byte]$b); $w.Write([byte]$b); $w.Write([byte]0); $w.Write([byte]0)
  $w.Write([uint16]1); $w.Write([uint16]32); $w.Write([uint32]$pngs[$i].Length); $w.Write([uint32]$offset)
  $offset += $pngs[$i].Length
}
foreach ($p in $pngs) { $w.Write($p) }
$w.Close()

# Mac : un .icns, PNG lui aussi.
$icns = New-Object System.IO.MemoryStream; $w = New-Object System.IO.BinaryWriter $icns
$parts = @(@("ic07", 128), @("ic08", 256), @("ic09", 512), @("ic10", 1024), @("ic11", 32), @("ic12", 64), @("ic13", 256), @("ic14", 512))
$body = New-Object System.IO.MemoryStream; $bw = New-Object System.IO.BinaryWriter $body
foreach ($pt in $parts) {
  $png = Icone-Png $pt[1]
  $bw.Write([System.Text.Encoding]::ASCII.GetBytes($pt[0])); Write-BE $bw ([uint32]($png.Length + 8)); $bw.Write($png)
}
$bw.Flush()
$w.Write([System.Text.Encoding]::ASCII.GetBytes("icns")); Write-BE $w ([uint32]($body.Length + 8)); $w.Write($body.ToArray()); $w.Flush()
$icnsBytes = $icns.ToArray()

# ------------------------------------------------------------- Windows

$smaPath = [PSObject].Assembly.Location
$lanceur = @'
using System;
using System.IO;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using System.Management.Automation;
using System.Management.Automation.Runspaces;

[assembly: AssemblyTitle("Adresses Outlook")]
[assembly: AssemblyProduct("Adresses Outlook")]
[assembly: AssemblyCompany("Aether")]
[assembly: AssemblyDescription("Relève les adresses que propose Outlook ou Gmail. N'envoie jamais rien.")]
[assembly: AssemblyVersion("2.0.0.0")]

// Fait tourner adresses-outlook.ps1, rangé dans l'exe, dans ce processus :
// pas de console, et l'icône de l'app dans la barre des tâches.
static class Lanceur {
  [STAThread]
  static void Main() {
    bool seul;
    using (Mutex m = new Mutex(true, "Aether.AdressesOutlook", out seul)) {
      if (!seul) { MessageBox.Show("Adresses Outlook est déjà ouvert.", "Adresses Outlook"); return; }
      string script;
      using (Stream s = Assembly.GetExecutingAssembly().GetManifestResourceStream("app.ps1"))
      using (StreamReader r = new StreamReader(s, Encoding.UTF8)) script = r.ReadToEnd();
      DateTime debut = DateTime.Now;
      try {
        using (Runspace rs = RunspaceFactory.CreateRunspace(InitialSessionState.CreateDefault())) {
          rs.ApartmentState = ApartmentState.STA;
          rs.ThreadOptions = PSThreadOptions.UseCurrentThread;
          rs.Open();
          using (PowerShell ps = PowerShell.Create()) {
            ps.Runspace = rs;
            ps.AddScript(script);
            ps.Invoke();
            // Une erreur qui empêche la fenêtre de s'ouvrir arrive dans les
            // premières secondes ; plus tard, la fenêtre l'a déjà affichée.
            if (ps.Streams.Error.Count > 0 && (DateTime.Now - debut).TotalSeconds < 5)
              Erreur(ps.Streams.Error[0].ToString());
          }
        }
      } catch (Exception e) { Erreur(e.Message); }
    }
  }
  static void Erreur(string message) {
    MessageBox.Show("Adresses Outlook n'a pas pu démarrer :\n\n" + message, "Adresses Outlook", MessageBoxButtons.OK, MessageBoxIcon.Error);
  }
}
'@
$cs = Join-Path $Out "lanceur.cs"
[System.IO.File]::WriteAllText($cs, $lanceur, (New-Object System.Text.UTF8Encoding $true))
$exe = Join-Path $Out "Adresses Outlook.exe"
$csc = Join-Path ([Runtime.InteropServices.RuntimeEnvironment]::GetRuntimeDirectory()) "csc.exe"
$ps1 = Join-Path $PSScriptRoot "adresses-outlook.ps1"
& $csc /nologo /target:winexe /optimize+ /codepage:65001 "/out:$exe" "/win32icon:$icoPath" "/resource:$ps1,app.ps1" `
  "/reference:$smaPath" /reference:System.Windows.Forms.dll $cs
if ($LASTEXITCODE -ne 0) { throw "La compilation de l'exe a échoué." }
Remove-Item $cs

# ----------------------------------------------------------------- Mac
# L'app Mac se fabrique sur un Mac (mac/natif/construire.sh, lancé par la
# construction automatique de GitHub). D'ici, on lui fournit son icône.

$icnsPath = Join-Path $PSScriptRoot "mac\natif\AppIcon.icns"
[System.IO.File]::WriteAllBytes($icnsPath, $icnsBytes)
Remove-Item $icoPath

Get-Item $exe, $icnsPath | ForEach-Object { "{0}  ({1:N0} ko)" -f $_.FullName, ($_.Length / 1KB) }
