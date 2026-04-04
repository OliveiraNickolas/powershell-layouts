# ============================================================
#  SnapLayout v2 - Gestor de janelas com Spaces, Layers e Canvas
#  Sem dependencias externas. Requer PowerShell 5+ no Windows 11.
#  Execucao: powershell -ExecutionPolicy Bypass -File layouts.ps1
# ============================================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

if (-not ([System.Management.Automation.PSTypeName]'WinAPI').Type) { Add-Type @"
using System;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using System.Text;

public class WinAPI {
    [DllImport("user32.dll")] public static extern bool SetWindowPos(
        IntPtr hWnd, IntPtr hWndInsertAfter,
        int X, int Y, int cx, int cy, uint uFlags);

    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern bool IsIconic(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int count);
    [DllImport("user32.dll")] public static extern int GetWindowTextLength(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] public static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
    [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern short GetAsyncKeyState(int vKey);
    [DllImport("user32.dll")] public static extern bool IsWindow(IntPtr hWnd);

    [StructLayout(LayoutKind.Sequential)]
    public struct RECT { public int Left, Top, Right, Bottom; }

    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    public const int SW_RESTORE   = 9;
    public const uint SWP_SHOWWINDOW = 0x0040;
    public const uint SWP_NOZORDER  = 0x0004;

    public static List<IntPtr> GetVisibleWindows() {
        var list = new List<IntPtr>();
        EnumWindows((hWnd, _) => {
            if (!IsWindowVisible(hWnd) || IsIconic(hWnd)) return true;
            int len = GetWindowTextLength(hWnd);
            if (len == 0) return true;
            var sb = new StringBuilder(len + 1);
            GetWindowText(hWnd, sb, sb.Capacity);
            string title = sb.ToString();
            if (title.Length > 1) list.Add(hWnd);
            return true;
        }, IntPtr.Zero);
        return list;
    }

    public static string GetTitle(IntPtr hWnd) {
        int len = GetWindowTextLength(hWnd);
        if (len == 0) return "";
        var sb = new StringBuilder(len + 1);
        GetWindowText(hWnd, sb, sb.Capacity);
        return sb.ToString();
    }

    public static void MoveWindow(IntPtr hWnd, int x, int y, int w, int h) {
        ShowWindow(hWnd, SW_RESTORE);
        SetWindowPos(hWnd, IntPtr.Zero, x, y, w, h, SWP_SHOWWINDOW | SWP_NOZORDER);
    }

    public static RECT GetRect(IntPtr hWnd) {
        RECT r;
        GetWindowRect(hWnd, out r);
        return r;
    }
}
"@ }

# ============================================================
#  CONSTANTES E ESTADO GLOBAL
# ============================================================

$screen = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
$SW     = $screen.Width
$SH     = $screen.Height
$SX     = $screen.X
$SY     = $screen.Y

$script:savePath = Join-Path $env:APPDATA "SnapLayout"
$script:saveFile = Join-Path $script:savePath "layouts.json"

# Layouts salvos pelo usuario: name -> @{ Spaces=@(...); Shortcut="" }
$script:savedLayouts = [ordered]@{}
# Layout atualmente selecionado (array de spaces com Layers)
$script:currentSpaces = @()
$script:currentSource = ""
$script:currentName   = ""
# Hotkey bindings
$script:hotkeyBindings = @()
$script:hotkeyCooldown = $false
# Space hotkey bindings
$script:spaceHotkeyBindings = @()
$script:spaceHotkeyCooldown = $false
# Index do space destacado no canvas
$script:highlightedSpaceIndex = -1
# Estado de drag-resize no canvas
$script:dragSpaceIdx  = -1
$script:dragEdge      = ""
$script:dragStartZone = @()

# ============================================================
#  FUNCOES UTILITARIAS
# ============================================================

function ConvertZone($zone) {
    return @{
        X = $SX + [int]($SW * $zone[0] / 100)
        Y = $SY + [int]($SH * $zone[1] / 100)
        W = [int]($SW * $zone[2] / 100)
        H = [int]($SH * $zone[3] / 100)
    }
}

function Get-VisibleWindows {
    $handles = [WinAPI]::GetVisibleWindows()
    $result  = @()
    foreach ($h in $handles) {
        # Verificar se janela está minimizada ou oculta
        if ([WinAPI]::IsIconic($h) -or -not [WinAPI]::IsWindowVisible($h)) {
            continue
        }
        $title = [WinAPI]::GetTitle($h)
        if ($title -and $title.Length -gt 1) {
            $result += [PSCustomObject]@{ Handle = $h; Title = $title }
        }
    }
    return $result | Sort-Object Title
}

# ============================================================
#  PERSISTENCIA (JSON)
# ============================================================

function Save-AllLayouts {
    if (-not (Test-Path $script:savePath)) {
        New-Item -Path $script:savePath -ItemType Directory -Force | Out-Null
    }
    $data = @{ version = 1; layouts = @() }
    foreach ($name in $script:savedLayouts.Keys) {
        $layout = $script:savedLayouts[$name]
        $spacesData = @()
        foreach ($s in $layout.Spaces) {
            $titles = @()
            foreach ($l in $s.Layers) { $titles += $l.Title }
            $spacesData += @{
                name         = $s.Name
                zone         = @($s.Zone)
                windowTitles = $titles
                shortcut     = if ($s.Shortcut) { $s.Shortcut } else { "" }
            }
        }
        $data.layouts += @{
            name     = $name
            shortcut = $layout.Shortcut
            spaces   = $spacesData
        }
    }
    $json = $data | ConvertTo-Json -Depth 6
    $tmpFile = $script:saveFile + ".tmp"
    $json | Set-Content $tmpFile -Encoding UTF8
    Move-Item $tmpFile $script:saveFile -Force
}

function Load-AllLayouts {
    if (-not (Test-Path $script:saveFile)) { return }
    try {
        $json = Get-Content $script:saveFile -Raw | ConvertFrom-Json
    } catch { return }
    foreach ($l in $json.layouts) {
        $spaces = @()
        foreach ($s in $l.spaces) {
            $sp = @{
                Name     = $s.name
                Zone     = @($s.zone)
                Layers   = [System.Collections.ArrayList]::new()
                Shortcut = if ($s.shortcut) { $s.shortcut } else { "" }
            }
            # Tentar encontrar janelas por titulo
            if ($s.windowTitles) {
                foreach ($wt in $s.windowTitles) {
                    if (-not $wt) { continue }
                    $match = Get-VisibleWindows | Where-Object { $_.Title -like "*$wt*" } | Select-Object -First 1
                    if ($match) {
                        [void]$sp.Layers.Add(@{ Handle = $match.Handle; Title = $match.Title })
                    }
                }
            }
            $spaces += $sp
        }
        $shortcut = if ($l.shortcut) { $l.shortcut } else { "" }
        $script:savedLayouts[$l.name] = @{ Spaces = $spaces; Shortcut = $shortcut }
    }
}

# ============================================================
#  CORES
# ============================================================

$cAccent  = [System.Drawing.Color]::FromArgb(100, 160, 255)
$cSurface = [System.Drawing.Color]::FromArgb(38, 38, 48)
$cBg      = [System.Drawing.Color]::FromArgb(24, 24, 30)
$cBorder  = [System.Drawing.Color]::FromArgb(55, 55, 68)
$cText    = [System.Drawing.Color]::FromArgb(230, 230, 240)
$cMuted   = [System.Drawing.Color]::FromArgb(130, 130, 150)
$cGreen   = [System.Drawing.Color]::FromArgb(72, 200, 120)
$cOrange  = [System.Drawing.Color]::FromArgb(255, 130, 50)
$cRed     = [System.Drawing.Color]::FromArgb(210, 55, 55)

$script:spaceColors = @(
    @{ Fill=[System.Drawing.Color]::FromArgb(40, 220, 60, 60);   Stroke=[System.Drawing.Color]::FromArgb(220, 60, 60) },
    @{ Fill=[System.Drawing.Color]::FromArgb(40, 80, 140, 255);  Stroke=[System.Drawing.Color]::FromArgb(80, 140, 255) },
    @{ Fill=[System.Drawing.Color]::FromArgb(40, 80, 210, 130);  Stroke=[System.Drawing.Color]::FromArgb(80, 210, 130) },
    @{ Fill=[System.Drawing.Color]::FromArgb(40, 255, 180, 40);  Stroke=[System.Drawing.Color]::FromArgb(255, 180, 40) },
    @{ Fill=[System.Drawing.Color]::FromArgb(40, 200, 80, 200);  Stroke=[System.Drawing.Color]::FromArgb(200, 80, 200) },
    @{ Fill=[System.Drawing.Color]::FromArgb(40, 80, 210, 210);  Stroke=[System.Drawing.Color]::FromArgb(80, 210, 210) }
)

# ============================================================
#  FORMULARIO PRINCIPAL
# ============================================================

$form = New-Object System.Windows.Forms.Form
$form.Text            = "SnapLayout v2"
$form.Size            = New-Object System.Drawing.Size(1080, 575)
$form.StartPosition   = "CenterScreen"
$form.FormBorderStyle = "FixedSingle"
$form.MaximizeBox     = $false
$form.BackColor       = $cBg
$form.ForeColor       = $cText
$form.Font            = New-Object System.Drawing.Font("Segoe UI", 9)

# -- Titulo
$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Text      = "SnapLayout"
$lblTitle.Font      = New-Object System.Drawing.Font("Segoe UI", 13, [System.Drawing.FontStyle]::Bold)
$lblTitle.ForeColor = $cAccent
$lblTitle.Location  = New-Object System.Drawing.Point(14, 12)
$lblTitle.Size      = New-Object System.Drawing.Size(190, 26)
$form.Controls.Add($lblTitle)

$lblSub = New-Object System.Windows.Forms.Label
$lblSub.Text      = "Spaces · Layers · Atalhos"
$lblSub.ForeColor = $cMuted
$lblSub.Font      = New-Object System.Drawing.Font("Segoe UI", 8)
$lblSub.Location  = New-Object System.Drawing.Point(14, 40)
$lblSub.Size      = New-Object System.Drawing.Size(190, 16)
$form.Controls.Add($lblSub)

# Barra accent no rodape do header
$sep = New-Object System.Windows.Forms.Panel
$sep.Location  = New-Object System.Drawing.Point(0, 62)
$sep.Size      = New-Object System.Drawing.Size(1080, 2)
$sep.BackColor = $cAccent
$form.Controls.Add($sep)

# Separadores verticais entre colunas (64 ate o sepBottom em 480)
$sepV1 = New-Object System.Windows.Forms.Panel
$sepV1.Location  = New-Object System.Drawing.Point(210, 64)
$sepV1.Size      = New-Object System.Drawing.Size(1, 416)
$sepV1.BackColor = $cBorder
$form.Controls.Add($sepV1)

$sepV2 = New-Object System.Windows.Forms.Panel
$sepV2.Location  = New-Object System.Drawing.Point(782, 64)
$sepV2.Size      = New-Object System.Drawing.Size(1, 416)
$sepV2.BackColor = $cBorder
$form.Controls.Add($sepV2)

# ============================================================
#  COLUNA ESQUERDA
# ============================================================

# -- Coluna esquerda: w=210 (x=0..210)
$btnSnapshot = New-Object System.Windows.Forms.Button
$btnSnapshot.Text      = "Snapshot do Desktop"
$btnSnapshot.Location  = New-Object System.Drawing.Point(8, 76)
$btnSnapshot.Size      = New-Object System.Drawing.Size(194, 32)
$btnSnapshot.FlatStyle = "Flat"
$btnSnapshot.BackColor = $cSurface
$btnSnapshot.ForeColor = $cText
$btnSnapshot.FlatAppearance.BorderColor = $cBorder
$btnSnapshot.Font      = New-Object System.Drawing.Font("Segoe UI", 9)
$form.Controls.Add($btnSnapshot)

$sepSnap = New-Object System.Windows.Forms.Panel
$sepSnap.Location  = New-Object System.Drawing.Point(8, 116)
$sepSnap.Size      = New-Object System.Drawing.Size(194, 1)
$sepSnap.BackColor = $cBorder
$form.Controls.Add($sepSnap)

$lblSaved = New-Object System.Windows.Forms.Label
$lblSaved.Text      = "LAYOUTS SALVOS"
$lblSaved.Font      = New-Object System.Drawing.Font("Segoe UI", 7, [System.Drawing.FontStyle]::Bold)
$lblSaved.ForeColor = $cMuted
$lblSaved.Location  = New-Object System.Drawing.Point(8, 124)
$lblSaved.Size      = New-Object System.Drawing.Size(194, 13)
$form.Controls.Add($lblSaved)

$lstSaved = New-Object System.Windows.Forms.ListBox
$lstSaved.Location    = New-Object System.Drawing.Point(8, 140)
$lstSaved.Size        = New-Object System.Drawing.Size(194, 330)
$lstSaved.BackColor   = $cSurface
$lstSaved.ForeColor   = $cText
$lstSaved.BorderStyle = "None"
$lstSaved.Font        = New-Object System.Drawing.Font("Segoe UI", 9)
$form.Controls.Add($lstSaved)

$btnDeleteSaved = New-Object System.Windows.Forms.Button
$btnDeleteSaved.Text      = "Excluir"
$btnDeleteSaved.Location  = New-Object System.Drawing.Point(8, 483)
$btnDeleteSaved.Size      = New-Object System.Drawing.Size(95, 34)
$btnDeleteSaved.FlatStyle = "Flat"
$btnDeleteSaved.BackColor = $cSurface
$btnDeleteSaved.ForeColor = $cRed
$btnDeleteSaved.FlatAppearance.BorderColor = $cRed
$btnDeleteSaved.Font      = New-Object System.Drawing.Font("Segoe UI", 8)
$form.Controls.Add($btnDeleteSaved)

$btnSetShortcut = New-Object System.Windows.Forms.Button
$btnSetShortcut.Text      = "Atalho..."
$btnSetShortcut.Location  = New-Object System.Drawing.Point(107, 483)
$btnSetShortcut.Size      = New-Object System.Drawing.Size(95, 34)
$btnSetShortcut.FlatStyle = "Flat"
$btnSetShortcut.BackColor = $cSurface
$btnSetShortcut.ForeColor = $cMuted
$btnSetShortcut.FlatAppearance.BorderColor = $cBorder
$btnSetShortcut.Font      = New-Object System.Drawing.Font("Segoe UI", 8)
$form.Controls.Add($btnSetShortcut)

# ============================================================
#  PRE-VISUALIZACAO — centro: x=212..782 (w=570)
# ============================================================

$lblPreview = New-Object System.Windows.Forms.Label
$lblPreview.Text      = "PREVIEW"
$lblPreview.Font      = New-Object System.Drawing.Font("Segoe UI", 7, [System.Drawing.FontStyle]::Bold)
$lblPreview.ForeColor = $cMuted
$lblPreview.Location  = New-Object System.Drawing.Point(220, 74)
$lblPreview.Size      = New-Object System.Drawing.Size(100, 13)
$form.Controls.Add($lblPreview)

$lblRes = New-Object System.Windows.Forms.Label
$lblRes.Text      = "${SW} × ${SH}"
$lblRes.ForeColor = $cMuted
$lblRes.Font      = New-Object System.Drawing.Font("Segoe UI", 7)
$lblRes.Location  = New-Object System.Drawing.Point(640, 75)
$lblRes.Size      = New-Object System.Drawing.Size(134, 13)
$lblRes.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$form.Controls.Add($lblRes)

$pnlPreview = New-Object System.Windows.Forms.Panel
$pnlPreview.Location    = New-Object System.Drawing.Point(220, 92)
$pnlPreview.Size        = New-Object System.Drawing.Size(554, 378)
$pnlPreview.BackColor   = $cSurface
$pnlPreview.BorderStyle = "None"
$form.Controls.Add($pnlPreview)

# ============================================================
#  SPACES / LAYERS — direita: x=784..1072 (w=288)
# ============================================================

$lblSpaces = New-Object System.Windows.Forms.Label
$lblSpaces.Text      = "SPACES & LAYERS"
$lblSpaces.Font      = New-Object System.Drawing.Font("Segoe UI", 7, [System.Drawing.FontStyle]::Bold)
$lblSpaces.ForeColor = $cMuted
$lblSpaces.Location  = New-Object System.Drawing.Point(792, 74)
$lblSpaces.Size      = New-Object System.Drawing.Size(280, 13)
$form.Controls.Add($lblSpaces)

$pnlSpaces = New-Object System.Windows.Forms.Panel
$pnlSpaces.Location            = New-Object System.Drawing.Point(792, 92)
$pnlSpaces.Size                = New-Object System.Drawing.Size(270, 378)
$pnlSpaces.BackColor           = $cSurface
$pnlSpaces.AutoScroll          = $true
$pnlSpaces.AutoScrollMinSize   = New-Object System.Drawing.Size(1, 1)
$form.Controls.Add($pnlSpaces)

# ============================================================
#  BARRA DE ACOES (baixo)
# ============================================================

$sepBottom = New-Object System.Windows.Forms.Panel
$sepBottom.Location  = New-Object System.Drawing.Point(0, 480)
$sepBottom.Size      = New-Object System.Drawing.Size(1080, 1)
$sepBottom.BackColor = $cBorder
$form.Controls.Add($sepBottom)

$btnApply = New-Object System.Windows.Forms.Button
$btnApply.Text      = "Aplicar Layout"
$btnApply.Location  = New-Object System.Drawing.Point(318, 483)
$btnApply.Size      = New-Object System.Drawing.Size(190, 34)
$btnApply.FlatStyle = "Flat"
$btnApply.BackColor = $cAccent
$btnApply.ForeColor = [System.Drawing.Color]::White
$btnApply.FlatAppearance.BorderSize = 0
$btnApply.Font      = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($btnApply)

$btnSaveCurrent = New-Object System.Windows.Forms.Button
$btnSaveCurrent.Text      = "Salvar Layout"
$btnSaveCurrent.Location  = New-Object System.Drawing.Point(516, 483)
$btnSaveCurrent.Size      = New-Object System.Drawing.Size(160, 34)
$btnSaveCurrent.FlatStyle = "Flat"
$btnSaveCurrent.BackColor = $cSurface
$btnSaveCurrent.ForeColor = $cAccent
$btnSaveCurrent.FlatAppearance.BorderColor = $cAccent
$btnSaveCurrent.Font      = New-Object System.Drawing.Font("Segoe UI", 9)
$form.Controls.Add($btnSaveCurrent)

# -- Status: linha discreta abaixo dos botoes, sem separador
$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Text      = ""
$lblStatus.ForeColor = $cMuted
$lblStatus.Font      = New-Object System.Drawing.Font("Segoe UI", 7.5)
$lblStatus.Location  = New-Object System.Drawing.Point(12, 520)
$lblStatus.Size      = New-Object System.Drawing.Size(1056, 14)
$form.Controls.Add($lblStatus)

# ============================================================
#  FUNCOES DE REFRESH
# ============================================================

function Refresh-SavedList {
    $lstSaved.Items.Clear()
    foreach ($name in $script:savedLayouts.Keys) {
        $sc = $script:savedLayouts[$name].Shortcut
        $display = if ($sc) { "$name [$sc]" } else { $name }
        [void]$lstSaved.Items.Add($display)
    }
}

# ============================================================
#  PREVIEW - DESENHO COM SPACES, LAYERS E BOTAO [+]
# ============================================================

$pnlPreview.add_Paint({
    param($s, $e)
    $g = $e.Graphics
    $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $g.Clear($cSurface)

    if ($script:currentSpaces.Count -eq 0) {
        $f = New-Object System.Drawing.Font("Segoe UI", 10)
        $g.DrawString("Selecione um layout", $f, [System.Drawing.Brushes]::Gray, 200, 170)
        return
    }

    $pw = $pnlPreview.Width  - 8
    $ph = $pnlPreview.Height - 8
    $i  = 0

    foreach ($space in $script:currentSpaces) {
        $zp = $space.Zone
        $rx = 4 + [int]($pw * $zp[0] / 100)
        $ry = 4 + [int]($ph * $zp[1] / 100)
        $rw = [int]($pw * $zp[2] / 100) - 4
        $rh = [int]($ph * $zp[3] / 100) - 4

        $ci = $i % $script:spaceColors.Count
        $brush = New-Object System.Drawing.SolidBrush($script:spaceColors[$ci].Fill)
        $pen   = New-Object System.Drawing.Pen($script:spaceColors[$ci].Stroke, 2)
        $g.FillRectangle($brush, $rx, $ry, $rw, $rh)
        $g.DrawRectangle($pen, $rx, $ry, $rw, $rh)

        # Desenhar highlight ao redor do space destacado
        if ($i -eq $script:highlightedSpaceIndex) {
            $highlightPen = New-Object System.Drawing.Pen([System.Drawing.Color]::Yellow, 3)
            $g.DrawRectangle($highlightPen, ($rx - 2), ($ry - 2), ($rw + 4), ($rh + 4))
        }

        # Nome do space
        $fontBold = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
        $fontSmall = New-Object System.Drawing.Font("Segoe UI", 7)
        $textBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::White)
        $mutedBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(180, 180, 195))

        $g.DrawString($space.Name, $fontBold, $textBrush, ($rx + 6), ($ry + 4))

        # Layers
        $ly = $ry + 22
        $layerIdx = 1
        foreach ($layer in $space.Layers) {
            if (($ly + 14) -gt ($ry + $rh - 20)) {
                $g.DrawString("...", $fontSmall, $mutedBrush, ($rx + 10), $ly)
                break
            }
            $lt = if ($layer.Title.Length -gt 20) { $layer.Title.Substring(0, 17) + "..." } else { $layer.Title }
            $g.DrawString("L${layerIdx}: $lt", $fontSmall, $mutedBrush, ($rx + 10), $ly)
            $ly += 14
            $layerIdx++
        }

        # Atalho do space (canto inferior direito)
        if ($space.Shortcut) {
            $fontShortcut = New-Object System.Drawing.Font("Segoe UI", 7, [System.Drawing.FontStyle]::Bold)
            $scBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(200, 180, 220, 255))
            $scText  = $space.Shortcut
            $scSize  = $g.MeasureString($scText, $fontShortcut)
            $g.DrawString($scText, $fontShortcut, $scBrush, ($rx + $rw - $scSize.Width - 4), ($ry + $rh - $scSize.Height - 2))
        }

        $i++
    }

    # Borda do preview
    $borderPen = New-Object System.Drawing.Pen($cBorder, 1)
    $g.DrawRectangle($borderPen, 0, 0, ($pnlPreview.Width - 1), ($pnlPreview.Height - 1))
})

# -- Double-click no preview para resetar highlight
$pnlPreview.add_MouseDoubleClick({
    param($s, $e)
    $script:highlightedSpaceIndex = -1
    $pnlPreview.Invalidate()
})

# -- Drag-resize: arrastar borda de um space para redimensionar
$pnlPreview.add_MouseMove({
    param($s, $e)
    if ($script:currentSpaces.Count -eq 0) { return }
    $pw  = $pnlPreview.Width  - 8
    $ph  = $pnlPreview.Height - 8
    $hit = 6

    if ($script:dragSpaceIdx -ge 0) {
        $space = $script:currentSpaces[$script:dragSpaceIdx]
        $oz = $script:dragStartZone
        $curX = [Math]::Round($e.X / $pw * 100)
        $curY = [Math]::Round($e.Y / $ph * 100)
        if ($script:dragEdge -match "right") {
            $space.Zone[2] = [Math]::Max(5, [Math]::Min(100 - $oz[0], $curX - $oz[0]))
        }
        if ($script:dragEdge -match "bottom") {
            $space.Zone[3] = [Math]::Max(5, [Math]::Min(100 - $oz[1], $curY - $oz[1]))
        }
        if ($script:dragEdge -match "left") {
            $newX = [Math]::Max(0, [Math]::Min($oz[0] + $oz[2] - 5, $curX))
            $space.Zone[2] = $oz[0] + $oz[2] - $newX
            $space.Zone[0] = $newX
        }
        if ($script:dragEdge -match "top") {
            $newY = [Math]::Max(0, [Math]::Min($oz[1] + $oz[3] - 5, $curY))
            $space.Zone[3] = $oz[1] + $oz[3] - $newY
            $space.Zone[1] = $newY
        }
        $pnlPreview.Invalidate()
        return
    }

    $found = $false
    foreach ($space in $script:currentSpaces) {
        $zp = $space.Zone
        $rx = 4 + [int]($pw * $zp[0] / 100)
        $ry = 4 + [int]($ph * $zp[1] / 100)
        $rw = [int]($pw * $zp[2] / 100) - 4
        $rh = [int]($ph * $zp[3] / 100) - 4
        $ex = $rx + $rw; $ey = $ry + $rh
        $nr = [Math]::Abs($e.X - $ex) -le $hit -and $e.Y -ge $ry -and $e.Y -le $ey
        $nb = [Math]::Abs($e.Y - $ey) -le $hit -and $e.X -ge $rx -and $e.X -le $ex
        $nl = [Math]::Abs($e.X - $rx) -le $hit -and $e.Y -ge $ry -and $e.Y -le $ey
        $nt = [Math]::Abs($e.Y - $ry) -le $hit -and $e.X -ge $rx -and $e.X -le $ex
        $cursor = if     ($nr -and $nb) { [System.Windows.Forms.Cursors]::SizeNWSE }
                  elseif ($nl -and $nt) { [System.Windows.Forms.Cursors]::SizeNWSE }
                  elseif ($nr -or $nl)  { [System.Windows.Forms.Cursors]::SizeWE }
                  elseif ($nb -or $nt)  { [System.Windows.Forms.Cursors]::SizeNS }
                  else                  { $null }
        if ($cursor) { $pnlPreview.Cursor = $cursor; $found = $true; break }
    }
    if (-not $found) { $pnlPreview.Cursor = [System.Windows.Forms.Cursors]::Default }
})

$pnlPreview.add_MouseDown({
    param($s, $e)
    if ($e.Button -ne [System.Windows.Forms.MouseButtons]::Left) { return }
    if ($script:currentSpaces.Count -eq 0) { return }
    $pw  = $pnlPreview.Width  - 8
    $ph  = $pnlPreview.Height - 8
    $hit = 6
    $idx = 0
    foreach ($space in $script:currentSpaces) {
        $zp = $space.Zone
        $rx = 4 + [int]($pw * $zp[0] / 100)
        $ry = 4 + [int]($ph * $zp[1] / 100)
        $rw = [int]($pw * $zp[2] / 100) - 4
        $rh = [int]($ph * $zp[3] / 100) - 4
        $ex = $rx + $rw; $ey = $ry + $rh
        $nr = [Math]::Abs($e.X - $ex) -le $hit -and $e.Y -ge $ry -and $e.Y -le $ey
        $nb = [Math]::Abs($e.Y - $ey) -le $hit -and $e.X -ge $rx -and $e.X -le $ex
        $nl = [Math]::Abs($e.X - $rx) -le $hit -and $e.Y -ge $ry -and $e.Y -le $ey
        $nt = [Math]::Abs($e.Y - $ry) -le $hit -and $e.X -ge $rx -and $e.X -le $ex
        $edge = ""
        if ($nr) { $edge += "right" }
        if ($nb) { $edge += "bottom" }
        if ($nl) { $edge += "left" }
        if ($nt) { $edge += "top" }
        if ($edge) {
            $script:dragSpaceIdx  = $idx
            $script:dragStartZone = @($zp[0], $zp[1], $zp[2], $zp[3])
            $script:dragEdge      = $edge
            $pnlPreview.Capture   = $true
            break
        }
        $idx++
    }
})

$pnlPreview.add_MouseUp({
    param($s, $e)
    if ($script:dragSpaceIdx -ge 0) {
        $script:dragSpaceIdx = -1
        $script:dragEdge     = ""
        $pnlPreview.Capture  = $false
        $pnlPreview.Cursor   = [System.Windows.Forms.Cursors]::Default
        Build-SpacePanel
        $pnlPreview.Invalidate()
    }
})

# ============================================================
#  PAINEL DE SPACES COM LAYERS
# ============================================================

function Build-SpacePanel {
    $pnlSpaces.Controls.Clear()
    $y = 6
    $i = 0

    foreach ($space in $script:currentSpaces) {
        $ci = $i % $script:spaceColors.Count
        $strokeColor = $script:spaceColors[$ci].Stroke

        # Indicador de cor + Nome + botoes
        $pnlHeader = New-Object System.Windows.Forms.Panel
        $pnlHeader.Location  = New-Object System.Drawing.Point(2, $y)
        $pnlHeader.Size      = New-Object System.Drawing.Size(250, 26)
        $pnlHeader.BackColor = $cSurface
        $pnlHeader.Cursor    = [System.Windows.Forms.Cursors]::Hand
        $pnlHeader.Tag       = $i
        # Click no header destaca o space no canvas
        $pnlHeader.add_Click({
            param($s, $e)
            if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Left) {
                $script:highlightedSpaceIndex = $s.Tag
                $pnlPreview.Invalidate()
            }
        })
        $pnlSpaces.Controls.Add($pnlHeader)

        $colorBar = New-Object System.Windows.Forms.Panel
        $colorBar.Location  = New-Object System.Drawing.Point(0, 0)
        $colorBar.Size      = New-Object System.Drawing.Size(5, 26)
        $colorBar.BackColor = $strokeColor
        $pnlHeader.Controls.Add($colorBar)

        # Label clicavel para renomear inline
        $lblName = New-Object System.Windows.Forms.Label
        $lblName.Text      = $space.Name
        $lblName.ForeColor = $cText
        $lblName.Font      = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Bold)
        $lblName.Location  = New-Object System.Drawing.Point(10, 4)
        $lblName.Size      = New-Object System.Drawing.Size(130, 18)
        $lblName.Cursor    = [System.Windows.Forms.Cursors]::IBeam

        # TextBox inline (oculto por padrao)
        $txtRename = New-Object System.Windows.Forms.TextBox
        $txtRename.Text        = $space.Name
        $txtRename.Font        = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Bold)
        $txtRename.Location    = New-Object System.Drawing.Point(9, 3)
        $txtRename.Size        = New-Object System.Drawing.Size(130, 20)
        $txtRename.BackColor   = $cSurface
        $txtRename.ForeColor   = $cText
        $txtRename.BorderStyle = "FixedSingle"
        $txtRename.Visible     = $false
        $txtRename.Tag         = @{ SpaceIdx = $i; Label = $lblName }

        $lblName.Tag = $txtRename
        $lblName.add_Click({
            param($s, $e)
            $txt       = $s.Tag
            $s.Visible = $false
            $txt.Visible = $true
            $txt.SelectAll()
            $txt.Focus()
        })

        $txtRename.add_KeyDown({
            param($s, $ke)
            if ($ke.KeyCode -eq [System.Windows.Forms.Keys]::Return) {
                $ke.SuppressKeyPress = $true
                $newName = $s.Text.Trim()
                if ($newName) { $script:currentSpaces[$s.Tag.SpaceIdx].Name = $newName }
                Build-SpacePanel
                $pnlPreview.Invalidate()
            } elseif ($ke.KeyCode -eq [System.Windows.Forms.Keys]::Escape) {
                $s.Visible            = $false
                $s.Tag.Label.Visible  = $true
            }
        })

        $txtRename.add_LostFocus({
            param($s, $e)
            $s.Visible           = $false
            $s.Tag.Label.Visible = $true
        })

        $pnlHeader.Controls.Add($lblName)
        $pnlHeader.Controls.Add($txtRename)

        # Botoes de reordenar
        $btnUp = New-Object System.Windows.Forms.Button
        $btnUp.Text      = [char]0x25B2
        $btnUp.Location  = New-Object System.Drawing.Point(148, 2)
        $btnUp.Size      = New-Object System.Drawing.Size(18, 22)
        $btnUp.FlatStyle = "Flat"
        $btnUp.BackColor = $cSurface
        $btnUp.ForeColor = $cMuted
        $btnUp.FlatAppearance.BorderSize = 0
        $btnUp.Font      = New-Object System.Drawing.Font("Segoe UI", 6)
        $btnUp.Tag       = $i
        $btnUp.add_Click({
            param($s, $e)
            $idx = $s.Tag
            if ($idx -gt 0) {
                $tmp = $script:currentSpaces[$idx - 1]
                $script:currentSpaces[$idx - 1] = $script:currentSpaces[$idx]
                $script:currentSpaces[$idx] = $tmp
                Build-SpacePanel
                $pnlPreview.Invalidate()
            }
        })
        $pnlHeader.Controls.Add($btnUp)

        $btnDown = New-Object System.Windows.Forms.Button
        $btnDown.Text      = [char]0x25BC
        $btnDown.Location  = New-Object System.Drawing.Point(168, 2)
        $btnDown.Size      = New-Object System.Drawing.Size(18, 22)
        $btnDown.FlatStyle = "Flat"
        $btnDown.BackColor = $cSurface
        $btnDown.ForeColor = $cMuted
        $btnDown.FlatAppearance.BorderSize = 0
        $btnDown.Font      = New-Object System.Drawing.Font("Segoe UI", 6)
        $btnDown.Tag       = $i
        $btnDown.add_Click({
            param($s, $e)
            $idx = $s.Tag
            if ($idx -lt ($script:currentSpaces.Count - 1)) {
                $tmp = $script:currentSpaces[$idx + 1]
                $script:currentSpaces[$idx + 1] = $script:currentSpaces[$idx]
                $script:currentSpaces[$idx] = $tmp
                Build-SpacePanel
                $pnlPreview.Invalidate()
            }
        })
        $pnlHeader.Controls.Add($btnDown)

        # Botao de deletar space
        $btnDelSpace = New-Object System.Windows.Forms.Button
        $btnDelSpace.Text      = "X"
        $btnDelSpace.Location  = New-Object System.Drawing.Point(226, 2)
        $btnDelSpace.Size      = New-Object System.Drawing.Size(22, 22)
        $btnDelSpace.FlatStyle = "Flat"
        $btnDelSpace.BackColor = $cSurface
        $btnDelSpace.ForeColor = $cRed
        $btnDelSpace.FlatAppearance.BorderSize = 0
        $btnDelSpace.Font      = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
        $btnDelSpace.Tag       = $i
        $btnDelSpace.add_Click({
            param($s, $e)
            $spIdx = $s.Tag
            $script:currentSpaces = $script:currentSpaces | Where-Object { $_ -ne $script:currentSpaces[$spIdx] }
            if ($script:highlightedSpaceIndex -eq $spIdx) {
                $script:highlightedSpaceIndex = -1
            } elseif ($script:highlightedSpaceIndex -gt $spIdx) {
                $script:highlightedSpaceIndex--
            }
            Build-SpacePanel
            $pnlPreview.Invalidate()
        })
        $pnlHeader.Controls.Add($btnDelSpace)

        $y += 30

        # Layers
        $layerIdx = 0
        foreach ($layer in $space.Layers) {
            $lt = if ($layer.Title.Length -gt 28) { $layer.Title.Substring(0, 25) + "..." } else { $layer.Title }

            $lblLayer = New-Object System.Windows.Forms.Label
            $lblLayer.Text      = "  L$($layerIdx + 1): $lt"
            $lblLayer.ForeColor = $cMuted
            $lblLayer.Font      = New-Object System.Drawing.Font("Segoe UI", 8)
            $lblLayer.Location  = New-Object System.Drawing.Point(10, $y)
            $lblLayer.Size      = New-Object System.Drawing.Size(208, 16)
            $pnlSpaces.Controls.Add($lblLayer)

            # Botao [X] remover layer
            $btnRemove = New-Object System.Windows.Forms.Button
            $btnRemove.Text      = "X"
            $btnRemove.Location  = New-Object System.Drawing.Point(218, ($y - 1))
            $btnRemove.Size      = New-Object System.Drawing.Size(20, 18)
            $btnRemove.FlatStyle = "Flat"
            $btnRemove.BackColor = $cSurface
            $btnRemove.ForeColor = $cRed
            $btnRemove.FlatAppearance.BorderSize = 0
            $btnRemove.Font      = New-Object System.Drawing.Font("Segoe UI", 7, [System.Drawing.FontStyle]::Bold)
            $btnRemove.Tag       = @{ SpaceIdx = $i; LayerIdx = $layerIdx }
            $btnRemove.add_Click({
                param($sender, $ea)
                $tag = $sender.Tag
                $sp  = $script:currentSpaces[$tag.SpaceIdx]
                if ($tag.LayerIdx -lt $sp.Layers.Count) {
                    $sp.Layers.RemoveAt($tag.LayerIdx)
                }
                Build-SpacePanel
                $pnlPreview.Invalidate()
            })
            $pnlSpaces.Controls.Add($btnRemove)
            $y += 18
            $layerIdx++
        }

        # Botao "Add Layer"
        $btnAdd = New-Object System.Windows.Forms.Button
        $btnAdd.Text      = "+ Add Layer"
        $btnAdd.Location  = New-Object System.Drawing.Point(4, $y)
        $btnAdd.Size      = New-Object System.Drawing.Size(120, 22)
        $btnAdd.FlatStyle = "Flat"
        $btnAdd.BackColor = $cSurface
        $btnAdd.ForeColor = $cGreen
        $btnAdd.FlatAppearance.BorderColor = $cGreen
        $btnAdd.Font      = New-Object System.Drawing.Font("Segoe UI", 7.5)
        $btnAdd.Tag       = $i
        $btnAdd.add_Click({
            param($sender, $ea)
            $spIdx = $sender.Tag
            # Criar popup de selecao de janela
            $popup = New-Object System.Windows.Forms.Form
            $popup.Text            = "Selecionar Janela"
            $popup.Size            = New-Object System.Drawing.Size(400, 350)
            $popup.StartPosition   = "CenterParent"
            $popup.FormBorderStyle = "FixedToolWindow"
            $popup.BackColor       = $cBg
            $popup.ForeColor       = $cText

            $lstPopup = New-Object System.Windows.Forms.ListBox
            $lstPopup.Location    = New-Object System.Drawing.Point(10, 10)
            $lstPopup.Size        = New-Object System.Drawing.Size(370, 260)
            $lstPopup.BackColor   = $cSurface
            $lstPopup.ForeColor   = $cText
            $lstPopup.BorderStyle = "None"
            $lstPopup.Font        = New-Object System.Drawing.Font("Segoe UI", 9)
            $popup.Controls.Add($lstPopup)

            $popupHandles = @()
            $wins = Get-VisibleWindows | Where-Object { $_.Title -notlike "*SnapLayout*" -and $_.Title.Length -gt 1 }
            foreach ($w in $wins) {
                $short = if ($w.Title.Length -gt 55) { $w.Title.Substring(0, 52) + "..." } else { $w.Title }
                [void]$lstPopup.Items.Add($short)
                $popupHandles += $w
            }

            $btnOk = New-Object System.Windows.Forms.Button
            $btnOk.Text         = "Adicionar"
            $btnOk.Location     = New-Object System.Drawing.Point(10, 278)
            $btnOk.Size         = New-Object System.Drawing.Size(100, 30)
            $btnOk.FlatStyle    = "Flat"
            $btnOk.BackColor    = $cAccent
            $btnOk.ForeColor    = [System.Drawing.Color]::White
            $btnOk.DialogResult = [System.Windows.Forms.DialogResult]::OK
            $popup.Controls.Add($btnOk)
            $popup.AcceptButton = $btnOk

            if ($popup.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK -and $lstPopup.SelectedIndex -ge 0) {
                $selWin = $popupHandles[$lstPopup.SelectedIndex]
                $space = $script:currentSpaces[$spIdx]
                [void]$space.Layers.Add(@{ Handle = $selWin.Handle; Title = $selWin.Title })
                Build-SpacePanel
                $pnlPreview.Invalidate()
            }
        })
        $pnlSpaces.Controls.Add($btnAdd)

        # Botao "Atalho..." por space
        $btnSpaceShortcut = New-Object System.Windows.Forms.Button
        $btnSpaceShortcut.Text      = "Atalho..."
        $btnSpaceShortcut.Location  = New-Object System.Drawing.Point(126, $y)
        $btnSpaceShortcut.Size      = New-Object System.Drawing.Size(120, 22)
        $btnSpaceShortcut.FlatStyle = "Flat"
        $btnSpaceShortcut.BackColor = $cSurface
        $btnSpaceShortcut.ForeColor = $cAccent
        $btnSpaceShortcut.FlatAppearance.BorderColor = $cAccent
        $btnSpaceShortcut.Font      = New-Object System.Drawing.Font("Segoe UI", 7.5)
        $btnSpaceShortcut.Tag       = $i
        $btnSpaceShortcut.add_Click({
            param($sender, $ea)
            $sp = $script:currentSpaces[$sender.Tag]

            $dlg = New-Object System.Windows.Forms.Form
            $dlg.Text            = "Atalho - $($sp.Name)"
            $dlg.Size            = New-Object System.Drawing.Size(350, 160)
            $dlg.StartPosition   = "CenterParent"
            $dlg.FormBorderStyle = "FixedToolWindow"
            $dlg.BackColor       = $cBg
            $dlg.ForeColor       = $cText
            $dlg.KeyPreview      = $true

            $lblInfo = New-Object System.Windows.Forms.Label
            $lblInfo.Text     = "Pressione a combinacao de teclas desejada:"
            $lblInfo.Location = New-Object System.Drawing.Point(10, 15)
            $lblInfo.Size     = New-Object System.Drawing.Size(320, 18)
            $dlg.Controls.Add($lblInfo)

            $txtKey = New-Object System.Windows.Forms.TextBox
            $txtKey.Location  = New-Object System.Drawing.Point(10, 40)
            $txtKey.Size      = New-Object System.Drawing.Size(315, 24)
            $txtKey.BackColor = $cSurface
            $txtKey.ForeColor = $cAccent
            $txtKey.Font      = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
            $txtKey.ReadOnly  = $true
            $txtKey.Text      = $sp.Shortcut
            $dlg.Controls.Add($txtKey)

            $btnOk = New-Object System.Windows.Forms.Button
            $btnOk.Text         = "Salvar"
            $btnOk.Location     = New-Object System.Drawing.Point(10, 78)
            $btnOk.Size         = New-Object System.Drawing.Size(80, 30)
            $btnOk.FlatStyle    = "Flat"
            $btnOk.BackColor    = $cAccent
            $btnOk.ForeColor    = [System.Drawing.Color]::White
            $btnOk.DialogResult = [System.Windows.Forms.DialogResult]::OK
            $dlg.Controls.Add($btnOk)

            $btnClear = New-Object System.Windows.Forms.Button
            $btnClear.Text      = "Limpar"
            $btnClear.Location  = New-Object System.Drawing.Point(100, 78)
            $btnClear.Size      = New-Object System.Drawing.Size(80, 30)
            $btnClear.FlatStyle = "Flat"
            $btnClear.BackColor = $cSurface
            $btnClear.ForeColor = $cMuted
            $btnClear.FlatAppearance.BorderColor = $cBorder
            $btnClear.add_Click({ $txtKey.Text = "" })
            $dlg.Controls.Add($btnClear)

            $dlg.add_KeyDown({
                param($s2, $ke)
                $ke.SuppressKeyPress = $true
                $parts = @()
                if ($ke.Control) { $parts += "Ctrl" }
                if ($ke.Alt)     { $parts += "Alt" }
                if ($ke.Shift)   { $parts += "Shift" }
                $k = $ke.KeyCode.ToString()
                if ($k -notin @("ControlKey","ShiftKey","Menu","LMenu","RMenu")) { $parts += $k }
                if ($parts.Count -gt 0) { $txtKey.Text = ($parts -join "+") }
            })

            if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                $sp.Shortcut = $txtKey.Text
                Register-SpaceHotkeys
                $pnlPreview.Invalidate()
                $short = if ($txtKey.Text) { "'$($txtKey.Text)'" } else { "removido" }
                $lblStatus.Text      = "Atalho $short definido para '$($sp.Name)'."
                $lblStatus.ForeColor = $cGreen
            }
        })
        $pnlSpaces.Controls.Add($btnSpaceShortcut)

        $y += 30

        # Separador
        $sepSpace = New-Object System.Windows.Forms.Panel
        $sepSpace.Location  = New-Object System.Drawing.Point(2, $y)
        $sepSpace.Size      = New-Object System.Drawing.Size(250, 1)
        $sepSpace.BackColor = $cBorder
        $pnlSpaces.Controls.Add($sepSpace)
        $y += 8
        $i++
    }
    Register-SpaceHotkeys
}

# ============================================================
#  SELECAO DE LAYOUT
# ============================================================

function Select-Layout($name, $source) {
    $script:currentName   = $name
    $script:currentSource = $source

    # Clonar spaces do layout salvo
    $saved = $script:savedLayouts[$name]
    $script:currentSpaces = @()
    foreach ($s in $saved.Spaces) {
        $sp = @{
            Name   = $s.Name
            Zone   = @($s.Zone)
            Layers = [System.Collections.ArrayList]::new()
        }
        foreach ($l in $s.Layers) {
            [void]$sp.Layers.Add(@{ Handle = $l.Handle; Title = $l.Title })
        }
        $script:currentSpaces += $sp
    }

    Build-SpacePanel
    $pnlPreview.Invalidate()
    $lblStatus.Text      = "Layout: $name"
    $lblStatus.ForeColor = $cMuted
}

$lstSaved.add_SelectedIndexChanged({
    $sel = $lstSaved.SelectedItem
    if ($sel) {
        # Extrair nome (remover shortcut display)
        $name = ($sel -replace '\s*\[.*\]\s*$', '')
        Select-Layout $name "saved"
    }
})

# ============================================================
#  EVENTOS DOS BOTOES
# ============================================================

# -- Aplicar Layout
$btnApply.add_Click({
    if ($script:currentSpaces.Count -eq 0) {
        $lblStatus.Text      = "Nenhum layout selecionado."
        $lblStatus.ForeColor = $cOrange
        return
    }
    $applied = 0
    foreach ($space in $script:currentSpaces) {
        $z = ConvertZone $space.Zone
        $layerCount = $space.Layers.Count
        if ($layerCount -le 0) {
            continue
        } elseif ($layerCount -eq 1) {
            $layerW = $z.W
            $layerH = $z.H
        } else {
            $totalGap = ($layerCount - 1) * 1
            $layerW = [Math]::Max(100, ($z.W - $totalGap) / $layerCount)
            $layerH = [Math]::Max(50, ($z.H - $totalGap) / $layerCount)
        }
        foreach ($i in 0..($layerCount - 1)) {
            $layer = $space.Layers[$i]
            $offset = $i * 1
            $lx = $z.X + ($offset - (($layerCount - 1) * 1) / 2)
            $ly = $z.Y + ($offset - (($layerCount - 1) * 1) / 2)
            try {
                [WinAPI]::MoveWindow($layer.Handle, [int]$lx, [int]$ly, [int]$layerW, [int]$layerH)
                $applied++
            } catch {}
        }
        # Ultima layer fica em foreground
        if ($space.Layers.Count -gt 0) {
            $topLayer = $space.Layers[$space.Layers.Count - 1]
            [WinAPI]::SetForegroundWindow($topLayer.Handle)
        }
    }
    $lblStatus.Text      = "Layout aplicado: $applied janela(s) posicionada(s)."
    $lblStatus.ForeColor = $cGreen
})

# -- Salvar Layout Atual
$btnSaveCurrent.add_Click({
    if ($script:currentSpaces.Count -eq 0) {
        $lblStatus.Text      = "Nenhum layout para salvar."
        $lblStatus.ForeColor = $cOrange
        return
    }

    # Dialog para nome
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text            = "Salvar Layout"
    $dlg.Size            = New-Object System.Drawing.Size(350, 150)
    $dlg.StartPosition   = "CenterParent"
    $dlg.FormBorderStyle = "FixedToolWindow"
    $dlg.BackColor       = $cBg
    $dlg.ForeColor       = $cText

    $lblN = New-Object System.Windows.Forms.Label
    $lblN.Text     = "Nome do layout:"
    $lblN.Location = New-Object System.Drawing.Point(10, 15)
    $lblN.Size     = New-Object System.Drawing.Size(320, 18)
    $dlg.Controls.Add($lblN)

    $txtName = New-Object System.Windows.Forms.TextBox
    $txtName.Location  = New-Object System.Drawing.Point(10, 38)
    $txtName.Size      = New-Object System.Drawing.Size(315, 24)
    $txtName.BackColor = $cSurface
    $txtName.ForeColor = $cText
    $txtName.Text      = $script:currentName
    $dlg.Controls.Add($txtName)

    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text         = "Salvar"
    $btnOk.Location     = New-Object System.Drawing.Point(10, 72)
    $btnOk.Size         = New-Object System.Drawing.Size(100, 30)
    $btnOk.FlatStyle    = "Flat"
    $btnOk.BackColor    = $cAccent
    $btnOk.ForeColor    = [System.Drawing.Color]::White
    $btnOk.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $dlg.Controls.Add($btnOk)
    $dlg.AcceptButton = $btnOk

    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK -and $txtName.Text.Trim()) {
        $saveName = $txtName.Text.Trim()
        $spaces = @()
        foreach ($s in $script:currentSpaces) {
            $sp = @{
                Name   = $s.Name
                Zone   = @($s.Zone)
                Layers = [System.Collections.ArrayList]::new()
            }
            foreach ($l in $s.Layers) {
                [void]$sp.Layers.Add(@{ Handle = $l.Handle; Title = $l.Title })
            }
            $spaces += $sp
        }
        $existingShortcut = ""
        if ($script:savedLayouts.Contains($saveName)) {
            $existingShortcut = $script:savedLayouts[$saveName].Shortcut
        }
        $script:savedLayouts[$saveName] = @{ Spaces = $spaces; Shortcut = $existingShortcut }
        Save-AllLayouts
        Refresh-SavedList
        $lblStatus.Text      = "Layout '$saveName' salvo com sucesso."
        $lblStatus.ForeColor = $cGreen
    }
})

# -- Excluir Layout Salvo
$btnDeleteSaved.add_Click({
    $sel = $lstSaved.SelectedItem
    if (-not $sel) { return }
    $name = ($sel -replace '\s*\[.*\]\s*$', '')
    $confirm = [System.Windows.Forms.MessageBox]::Show(
        "Excluir o layout '$name'?", "Confirmar",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )
    if ($confirm -eq [System.Windows.Forms.DialogResult]::Yes) {
        $script:savedLayouts.Remove($name)
        Save-AllLayouts
        Refresh-SavedList
        $script:currentSpaces = @()
        Build-SpacePanel
        $pnlPreview.Invalidate()
        $lblStatus.Text      = "Layout '$name' excluido."
        $lblStatus.ForeColor = $cMuted
    }
})

# -- Definir Atalho
$btnSetShortcut.add_Click({
    $sel = $lstSaved.SelectedItem
    if (-not $sel) {
        $lblStatus.Text      = "Selecione um layout salvo primeiro."
        $lblStatus.ForeColor = $cOrange
        return
    }
    $name = ($sel -replace '\s*\[.*\]\s*$', '')

    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text            = "Definir Atalho - $name"
    $dlg.Size            = New-Object System.Drawing.Size(350, 160)
    $dlg.StartPosition   = "CenterParent"
    $dlg.FormBorderStyle = "FixedToolWindow"
    $dlg.BackColor       = $cBg
    $dlg.ForeColor       = $cText
    $dlg.KeyPreview      = $true

    $lblInfo = New-Object System.Windows.Forms.Label
    $lblInfo.Text     = "Pressione a combinacao de teclas desejada:"
    $lblInfo.Location = New-Object System.Drawing.Point(10, 15)
    $lblInfo.Size     = New-Object System.Drawing.Size(320, 18)
    $dlg.Controls.Add($lblInfo)

    $txtKey = New-Object System.Windows.Forms.TextBox
    $txtKey.Location  = New-Object System.Drawing.Point(10, 40)
    $txtKey.Size      = New-Object System.Drawing.Size(315, 24)
    $txtKey.BackColor = $cSurface
    $txtKey.ForeColor = $cAccent
    $txtKey.Font      = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
    $txtKey.ReadOnly  = $true
    $txtKey.Text      = $script:savedLayouts[$name].Shortcut
    $dlg.Controls.Add($txtKey)

    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text         = "Salvar"
    $btnOk.Location     = New-Object System.Drawing.Point(10, 78)
    $btnOk.Size         = New-Object System.Drawing.Size(80, 30)
    $btnOk.FlatStyle    = "Flat"
    $btnOk.BackColor    = $cAccent
    $btnOk.ForeColor    = [System.Drawing.Color]::White
    $btnOk.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $dlg.Controls.Add($btnOk)

    $btnClear = New-Object System.Windows.Forms.Button
    $btnClear.Text      = "Limpar"
    $btnClear.Location  = New-Object System.Drawing.Point(100, 78)
    $btnClear.Size      = New-Object System.Drawing.Size(80, 30)
    $btnClear.FlatStyle = "Flat"
    $btnClear.BackColor = $cSurface
    $btnClear.ForeColor = $cMuted
    $btnClear.FlatAppearance.BorderColor = $cBorder
    $btnClear.add_Click({ $txtKey.Text = "" })
    $dlg.Controls.Add($btnClear)

    $dlg.add_KeyDown({
        param($s2, $ke)
        $ke.SuppressKeyPress = $true
        $parts = @()
        if ($ke.Control) { $parts += "Ctrl" }
        if ($ke.Alt)     { $parts += "Alt" }
        if ($ke.Shift)   { $parts += "Shift" }
        $key = $ke.KeyCode.ToString()
        if ($key -notin @("ControlKey","ShiftKey","Menu","LMenu","RMenu")) {
            $parts += $key
        }
        if ($parts.Count -gt 0) {
            $txtKey.Text = ($parts -join "+")
        }
    })

    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $script:savedLayouts[$name].Shortcut = $txtKey.Text
        Save-AllLayouts
        Refresh-SavedList
        Register-Hotkeys
        $lblStatus.Text      = "Atalho '$($txtKey.Text)' definido para '$name'."
        $lblStatus.ForeColor = $cGreen
    }
})

# ============================================================
#  SNAPSHOT DO DESKTOP
# ============================================================

$snapshotIgnoredTitles = @(
    "ICA Seamless Host Agent",
    "Program Manager",
    "Settings",
    "Windows Input Experience"
)

$btnSnapshot.add_Click({
    $wins = Get-VisibleWindows | Where-Object {
        $_.Title -notlike "*SnapLayout*" -and
        $_.Title.Length -gt 1 -and
        $snapshotIgnoredTitles -notcontains $_.Title
    }
    if ($wins.Count -eq 0) {
        $lblStatus.Text      = "Nenhuma janela visivel encontrada."
        $lblStatus.ForeColor = $cOrange
        return
    }

    $script:currentSpaces = @()
    $spaceIdx = 1
    foreach ($w in $wins) {
        $rect = [WinAPI]::GetRect($w.Handle)
        # Converter para percentagem
        $xp = [Math]::Max(0, [Math]::Min(100, [Math]::Round(($rect.Left - $SX) / $SW * 100)))
        $yp = [Math]::Max(0, [Math]::Min(100, [Math]::Round(($rect.Top - $SY) / $SH * 100)))
        $wp = [Math]::Max(5, [Math]::Min(100, [Math]::Round(($rect.Right - $rect.Left) / $SW * 100)))
        $hp = [Math]::Max(5, [Math]::Min(100, [Math]::Round(($rect.Bottom - $rect.Top) / $SH * 100)))
        # Snap para grid de 5%
        $xp = [Math]::Round($xp / 5) * 5
        $yp = [Math]::Round($yp / 5) * 5
        $wp = [Math]::Round($wp / 5) * 5
        $hp = [Math]::Round($hp / 5) * 5
        if ($wp -lt 5) { $wp = 5 }
        if ($hp -lt 5) { $hp = 5 }

        $sp = @{
            Name     = "Space $spaceIdx"
            Zone     = @([int]$xp, [int]$yp, [int]$wp, [int]$hp)
            Layers   = [System.Collections.ArrayList]::new()
            Shortcut = ""
        }
        [void]$sp.Layers.Add(@{ Handle = $w.Handle; Title = $w.Title })
        $script:currentSpaces += $sp
        $spaceIdx++
    }

    $script:currentName   = "Snapshot"
    $script:currentSource = "snapshot"
    $lstSaved.ClearSelected()
    Build-SpacePanel
    $pnlPreview.Invalidate()
    $lblStatus.Text      = "Snapshot: $($script:currentSpaces.Count) janela(s) capturada(s). Salve para manter."
    $lblStatus.ForeColor = $cGreen
})

# ============================================================
#  HOTKEYS (polling via GetAsyncKeyState)
# ============================================================

# Mapa de nomes de tecla para VK codes
$script:vkMap = @{
    "F1"=0x70; "F2"=0x71; "F3"=0x72; "F4"=0x73; "F5"=0x74; "F6"=0x75
    "F7"=0x76; "F8"=0x77; "F9"=0x78; "F10"=0x79; "F11"=0x7A; "F12"=0x7B
    "D1"=0x31; "D2"=0x32; "D3"=0x33; "D4"=0x34; "D5"=0x35; "D6"=0x36
    "D7"=0x37; "D8"=0x38; "D9"=0x39; "D0"=0x30
    "A"=0x41; "B"=0x42; "C"=0x43; "D"=0x44; "E"=0x45; "F"=0x46
    "G"=0x47; "H"=0x48; "I"=0x49; "J"=0x4A; "K"=0x4B; "L"=0x4C
    "M"=0x4D; "N"=0x4E; "O"=0x4F; "P"=0x50; "Q"=0x51; "R"=0x52
    "S"=0x53; "T"=0x54; "U"=0x55; "V"=0x56; "W"=0x57; "X"=0x58
    "Y"=0x59; "Z"=0x5A
}

function Register-Hotkeys {
    $script:hotkeyBindings = @()
    foreach ($name in $script:savedLayouts.Keys) {
        $sc = $script:savedLayouts[$name].Shortcut
        if (-not $sc) { continue }
        $parts = $sc -split '\+'
        $mods = @()
        $key  = ""
        foreach ($p in $parts) {
            $p = $p.Trim()
            if ($p -eq "Ctrl" -or $p -eq "Control") { $mods += "Ctrl" }
            elseif ($p -eq "Alt")   { $mods += "Alt" }
            elseif ($p -eq "Shift") { $mods += "Shift" }
            else { $key = $p }
        }
        if (-not $key) { continue }
        $vk = $script:vkMap[$key]
        if (-not $vk) { continue }
        $script:hotkeyBindings += @{
            LayoutName = $name
            Mods       = $mods
            VKey       = $vk
            Cooldown   = $false
        }
    }
}

function Register-SpaceHotkeys {
    $script:spaceHotkeyBindings = @()
    foreach ($space in $script:currentSpaces) {
        $sc = $space.Shortcut
        if (-not $sc) { continue }
        $parts = $sc -split '\+'
        $mods = @()
        $key  = ""
        foreach ($p in $parts) {
            $p = $p.Trim()
            if ($p -eq "Ctrl" -or $p -eq "Control") { $mods += "Ctrl" }
            elseif ($p -eq "Alt")   { $mods += "Alt" }
            elseif ($p -eq "Shift") { $mods += "Shift" }
            else { $key = $p }
        }
        if (-not $key) { continue }
        $vk = $script:vkMap[$key]
        if (-not $vk) { continue }
        $script:spaceHotkeyBindings += @{
            Space = $space
            Mods  = $mods
            VKey  = $vk
        }
    }
}

function Apply-SpaceHotkey($space) {
    $foreWin = [WinAPI]::GetForegroundWindow()
    $title   = [WinAPI]::GetTitle($foreWin)
    if (-not $title -or $title -like "*SnapLayout*") { return }
    # Se o handle ja e layer neste space, apenas reposiciona sem duplicar
    $existing = $space.Layers | Where-Object { $_.Handle -eq $foreWin } | Select-Object -First 1
    if (-not $existing) {
        [void]$space.Layers.Add(@{ Handle = $foreWin; Title = $title })
    }
    $z = ConvertZone $space.Zone
    [WinAPI]::MoveWindow($foreWin, $z.X, $z.Y, $z.W, $z.H)
    [WinAPI]::SetForegroundWindow($foreWin)
    Build-SpacePanel
    $pnlPreview.Invalidate()
    $short = if ($title.Length -gt 40) { $title.Substring(0, 37) + "..." } else { $title }
    $lblStatus.Text      = "'$short' encaixado em '$($space.Name)'."
    $lblStatus.ForeColor = $cGreen
}

function Apply-LayoutByName($name) {
    $layout = $script:savedLayouts[$name]
    if (-not $layout) { return }
    foreach ($space in $layout.Spaces) {
        $z = ConvertZone $space.Zone
        $layerCount = $space.Layers.Count
        if ($layerCount -le 0) {
            continue
        } elseif ($layerCount -eq 1) {
            $layerW = $z.W
            $layerH = $z.H
        } else {
            $totalGap = ($layerCount - 1) * 1
            $layerW = [Math]::Max(100, ($z.W - $totalGap) / $layerCount)
            $layerH = [Math]::Max(50, ($z.H - $totalGap) / $layerCount)
        }
        foreach ($i in 0..($layerCount - 1)) {
            $layer = $space.Layers[$i]
            $offset = $i * 1
            $lx = $z.X + ($offset - (($layerCount - 1) * 1) / 2)
            $ly = $z.Y + ($offset - (($layerCount - 1) * 1) / 2)
            try {
                [WinAPI]::MoveWindow($layer.Handle, [int]$lx, [int]$ly, [int]$layerW, [int]$layerH)
            } catch {}
        }
        if ($space.Layers.Count -gt 0) {
            $topLayer = $space.Layers[$space.Layers.Count - 1]
            [WinAPI]::SetForegroundWindow($topLayer.Handle)
        }
    }
}

$hotkeyTimer = New-Object System.Windows.Forms.Timer
$hotkeyTimer.Interval = 200
$hotkeyTimer.add_Tick({
    # --- Pruning: remover layers de janelas fechadas ---
    $pruned = $false
    $visibleTitles = $null  # lazy: so carregado se necessario
    foreach ($space in $script:currentSpaces) {
        $toRemove = @()
        foreach ($layer in $space.Layers) {
            $gone = -not [WinAPI]::IsWindow($layer.Handle)
            if (-not $gone -and $layer.Title) {
                $curTitle = [WinAPI]::GetTitle($layer.Handle)
                # Se o titulo do HWND mudou, o documento original pode ter fechado
                # (ex: Word MDI onde varios docs compartilham o mesmo HWND).
                # Confirma buscando se alguma janela visivel ainda tem o titulo original.
                if ($curTitle -and $curTitle -ne $layer.Title) {
                    if ($null -eq $visibleTitles) {
                        $visibleTitles = @(Get-VisibleWindows | ForEach-Object { $_.Title })
                    }
                    if ($visibleTitles -notcontains $layer.Title) { $gone = $true }
                }
            }
            if ($gone) { $toRemove += $layer }
        }
        foreach ($r in $toRemove) {
            $space.Layers.Remove($r)
            $pruned = $true
        }
    }
    if ($pruned) {
        Build-SpacePanel
        $pnlPreview.Invalidate()
    }

    if ($script:hotkeyCooldown) { return }

    # --- Space hotkeys ---
    if (-not $script:spaceHotkeyCooldown) {
        foreach ($binding in $script:spaceHotkeyBindings) {
            $modsOk = $true
            if ($binding.Mods -contains "Ctrl")  { if (-not ([WinAPI]::GetAsyncKeyState(0x11) -band 0x8000)) { $modsOk = $false } }
            if ($binding.Mods -contains "Alt")   { if (-not ([WinAPI]::GetAsyncKeyState(0x12) -band 0x8000)) { $modsOk = $false } }
            if ($binding.Mods -contains "Shift") { if (-not ([WinAPI]::GetAsyncKeyState(0x10) -band 0x8000)) { $modsOk = $false } }
            if (-not $modsOk) { continue }
            $keyState = [WinAPI]::GetAsyncKeyState($binding.VKey)
            if ($keyState -band 0x8000) {
                Apply-SpaceHotkey $binding.Space
                $script:spaceHotkeyCooldown = $true
                $cd = New-Object System.Windows.Forms.Timer
                $cd.Interval = 600
                $cd.add_Tick({ $script:spaceHotkeyCooldown = $false; $this.Stop(); $this.Dispose() })
                $cd.Start()
                return
            }
        }
    }

    # --- Layout hotkeys ---
    foreach ($binding in $script:hotkeyBindings) {
        $modsOk = $true
        if ($binding.Mods -contains "Ctrl") {
            if (-not ([WinAPI]::GetAsyncKeyState(0x11) -band 0x8000)) { $modsOk = $false }
        }
        if ($binding.Mods -contains "Alt") {
            if (-not ([WinAPI]::GetAsyncKeyState(0x12) -band 0x8000)) { $modsOk = $false }
        }
        if ($binding.Mods -contains "Shift") {
            if (-not ([WinAPI]::GetAsyncKeyState(0x10) -band 0x8000)) { $modsOk = $false }
        }
        if (-not $modsOk) { continue }

        $keyState = [WinAPI]::GetAsyncKeyState($binding.VKey)
        if ($keyState -band 0x8000) {
            Apply-LayoutByName $binding.LayoutName
            $lblStatus.Text      = "Atalho: Layout '$($binding.LayoutName)' aplicado."
            $lblStatus.ForeColor = $cGreen
            $script:hotkeyCooldown = $true
            $cooldownTimer = New-Object System.Windows.Forms.Timer
            $cooldownTimer.Interval = 600
            $cooldownTimer.add_Tick({
                $script:hotkeyCooldown = $false
                $this.Stop()
                $this.Dispose()
            })
            $cooldownTimer.Start()
            return
        }
    }
})
$hotkeyTimer.Start()

# ============================================================
#  INICIALIZACAO
# ============================================================

Load-AllLayouts
Refresh-SavedList
Register-Hotkeys

[void]$form.ShowDialog()
$hotkeyTimer.Stop()
$hotkeyTimer.Dispose()
