# ============================================================
# SnapLayout v4 - Gestor de janelas com Spaces e Atalhos
# Sem dependencias externas. Snapshot + Atalhos por Space
# ============================================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

Add-Type @"
using System;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using System.Text;

public class WinAPI {
    [DllImport("user32.dll")] public static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern bool IsIconic(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int count);
    [DllImport("user32.dll")] public static extern int GetWindowTextLength(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern short GetAsyncKeyState(int vKey);
    [DllImport("user32.dll")] public static extern bool IsWindow(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

    [StructLayout(LayoutKind.Sequential)]
    public struct RECT { public int Left, Top, Right, Bottom; }

    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    public static List<IntPtr> GetVisibleWindows() {
        var list = new List<IntPtr>();
        EnumWindows((hWnd, _) => {
            if (!IsWindowVisible(hWnd) || IsIconic(hWnd)) return true;
            int len = GetWindowTextLength(hWnd);
            if (len == 0) return true;
            var sb = new StringBuilder(len + 1);
            GetWindowText(hWnd, sb, sb.Capacity);
            if (sb.ToString().Length > 1) list.Add(hWnd);
            return true;
        }, IntPtr.Zero);
        return list;
    }
    public const int SW_RESTORE = 9;
    public const uint SWP_SHOWWINDOW = 0x0040;
    public const uint SWP_NOZORDER = 0x0004;
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
    public static RECT GetRect(IntPtr hWnd) { RECT r; GetWindowRect(hWnd, out r); return r; }
}
"@

$screen = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
$SW = $screen.Width; $SH = $screen.Height; $SX = $screen.X; $SY = $screen.Y
$script:savePath = Join-Path $env:APPDATA "SnapLayout"
$script:saveFile = Join-Path $script:savePath "layouts.json"
$script:windowGap = 2
$script:savedLayouts = [ordered]@{}
$script:currentSpaces = @()
$script:currentName = ""
$script:highlightedSpaceIndex = -1
$script:activeLayoutName = ""
$script:hotkeyBindings = @()
$script:hotkeyCooldown = $false
$script:spaceHotkeyBindings = @()
$script:spaceHotkeyCooldown = $false
$script:dynamicLayers = @()

function ConvertZone($zone) {
    @{ X = $SX + [int]($SW * $zone[0] / 100); Y = $SY + [int]($SH * $zone[1] / 100); W = [int]($SW * $zone[2] / 100); H = [int]($SH * $zone[3] / 100) }
}

function Get-VisibleWindows {
    $handles = [WinAPI]::GetVisibleWindows()
    $wa = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
    $result = @()
    foreach ($h in $handles) {
        if ([WinAPI]::IsIconic($h) -or -not [WinAPI]::IsWindowVisible($h)) { continue }
        $title = [WinAPI]::GetTitle($h)
        if (-not $title -or $title.Length -le 1) { continue }
        $rect = [WinAPI]::GetRect($h)
        if (($rect.Right - $rect.Left) -lt 50 -or ($rect.Bottom - $rect.Top) -lt 50) { continue }
        if (-not ($rect.Right -gt $wa.X -and $rect.Left -lt ($wa.X + $wa.Width) -and $rect.Bottom -gt $wa.Y -and $rect.Top -lt ($wa.Y + $wa.Height))) { continue }
        $result += [PSCustomObject]@{ Handle = $h; Title = $title }
    }
    return $result | Sort-Object Title
}

function Save-AllLayouts {
    if (-not (Test-Path $script:savePath)) { New-Item -Path $script:savePath -ItemType Directory -Force | Out-Null }
    $data = @{ version = 2; layouts = @() }
    foreach ($name in $script:savedLayouts.Keys) {
        $layout = $script:savedLayouts[$name]
        $spacesData = @()
        foreach ($s in $layout.Spaces) {
            $titles = @(); foreach ($l in $s.Layers) { $titles += $l.Title }
            $spacesData += @{ name = $s.Name; zone = @($s.Zone); windowTitles = $titles; shortcut = $s.Shortcut }
        }
        $data.layouts += @{ name = $name; shortcut = $layout.Shortcut; spaces = $spacesData }
    }
    $data | ConvertTo-Json -Depth 6 | Set-Content $script:saveFile -Encoding UTF8
}

function Load-AllLayouts {
    if (-not (Test-Path $script:saveFile)) { return }
    try { $json = Get-Content $script:saveFile -Raw | ConvertFrom-Json } catch { return }
    foreach ($l in $json.layouts) {
        $spaces = @()
        foreach ($s in $l.spaces) {
            $sp = @{ Name = $s.name; Zone = @($s.zone); Layers = [System.Collections.ArrayList]::new(); Shortcut = if ($s.shortcut) { $s.shortcut } else { "" } }
            if ($s.windowTitles) {
                foreach ($wt in $s.windowTitles) {
                    if (-not $wt) { continue }
                    $match = Get-VisibleWindows | Where-Object { $_.Title -like "*$wt*" } | Select-Object -First 1
                    if ($match) { [void]$sp.Layers.Add(@{ Handle = $match.Handle; Title = $match.Title }) }
                }
            }
            $spaces += $sp
        }
        $script:savedLayouts[$l.name] = @{ Spaces = $spaces; Shortcut = if ($l.shortcut) { $l.shortcut } else { "" } }
    }
}

$cAccent = [System.Drawing.Color]::FromArgb(90, 130, 240)
$cSurface = [System.Drawing.Color]::FromArgb(42, 42, 52)
$cSurface2 = [System.Drawing.Color]::FromArgb(52, 52, 62)
$cBg = [System.Drawing.Color]::FromArgb(24, 24, 30)
$cBorder = [System.Drawing.Color]::FromArgb(68, 68, 78)
$cText = [System.Drawing.Color]::FromArgb(240, 240, 250)
$cMuted = [System.Drawing.Color]::FromArgb(160, 160, 180)
$cGreen = [System.Drawing.Color]::FromArgb(80, 220, 130)
$cOrange = [System.Drawing.Color]::FromArgb(255, 160, 70)
$cRed = [System.Drawing.Color]::FromArgb(230, 75, 75)

$script:spaceColors = @(
    @{ Fill=[System.Drawing.Color]::FromArgb(45, 60, 230, 100); Stroke=[System.Drawing.Color]::FromArgb(100, 140, 245) },
    @{ Fill=[System.Drawing.Color]::FromArgb(45, 90, 200, 245); Stroke=[System.Drawing.Color]::FromArgb(100, 150, 250) },
    @{ Fill=[System.Drawing.Color]::FromArgb(45, 100, 230, 120); Stroke=[System.Drawing.Color]::FromArgb(100, 200, 140) },
    @{ Fill=[System.Drawing.Color]::FromArgb(45, 220, 150, 50); Stroke=[System.Drawing.Color]::FromArgb(245, 180, 80) },
    @{ Fill=[System.Drawing.Color]::FromArgb(45, 140, 60, 220); Stroke=[System.Drawing.Color]::FromArgb(180, 100, 230) },
    @{ Fill=[System.Drawing.Color]::FromArgb(45, 180, 200, 220); Stroke=[System.Drawing.Color]::FromArgb(100, 160, 240) }
)

$form = New-Object System.Windows.Forms.Form
$form.Text = "SnapLayout v4"
$form.Size = New-Object System.Drawing.Size(1100, 650)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedSingle"
$form.MaximizeBox = $false
$form.BackColor = $cBg
$form.ForeColor = $cText
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9)

$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Text = "SnapLayout v4 - Spaces & Atalhos"
$lblTitle.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
$lblTitle.ForeColor = $cAccent
$lblTitle.Location = New-Object System.Drawing.Point(16, 16)
$lblTitle.Size = New-Object System.Drawing.Size(300, 32)
$form.Controls.Add($lblTitle)

$sep = New-Object System.Windows.Forms.Panel
$sep.Location = New-Object System.Drawing.Point(16, 52)
$sep.Size = New-Object System.Drawing.Size(1060, 1)
$sep.BackColor = $cBorder
$form.Controls.Add($sep)

$lblSaved = New-Object System.Windows.Forms.Label
$lblSaved.Text = "LAYOUTS SALVOS"
$lblSaved.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
$lblSaved.ForeColor = $cMuted
$lblSaved.Location = New-Object System.Drawing.Point(16, 68)
$lblSaved.Size = New-Object System.Drawing.Size(150, 18)
$form.Controls.Add($lblSaved)

$lstSaved = New-Object System.Windows.Forms.ListBox
$lstSaved.Location = New-Object System.Drawing.Point(16, 88)
$lstSaved.Size = New-Object System.Drawing.Size(240, 358)
$lstSaved.BackColor = $cSurface
$lstSaved.ForeColor = $cText
$lstSaved.BorderStyle = "None"
$lstSaved.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
$form.Controls.Add($lstSaved)

$btnRenameLayout = New-Object System.Windows.Forms.Button
$btnRenameLayout.Text = "Renomear"
$btnRenameLayout.Location = New-Object System.Drawing.Point(16, 454)
$btnRenameLayout.Size = New-Object System.Drawing.Size(100, 26)
$btnRenameLayout.FlatStyle = "Flat"
$btnRenameLayout.BackColor = $cSurface
$btnRenameLayout.ForeColor = $cAccent
$btnRenameLayout.FlatAppearance.BorderColor = $cAccent
$btnRenameLayout.Font = New-Object System.Drawing.Font("Segoe UI", 8)
$btnRenameLayout.Cursor = [System.Windows.Forms.Cursors]::Hand
$form.Controls.Add($btnRenameLayout)

$btnSetShortcut = New-Object System.Windows.Forms.Button
$btnSetShortcut.Text = "Atalho"
$btnSetShortcut.Location = New-Object System.Drawing.Point(126, 454)
$btnSetShortcut.Size = New-Object System.Drawing.Size(100, 26)
$btnSetShortcut.FlatStyle = "Flat"
$btnSetShortcut.BackColor = $cSurface
$btnSetShortcut.ForeColor = $cMuted
$btnSetShortcut.FlatAppearance.BorderColor = $cBorder
$btnSetShortcut.Font = New-Object System.Drawing.Font("Segoe UI", 8)
$btnSetShortcut.Cursor = [System.Windows.Forms.Cursors]::Hand
$form.Controls.Add($btnSetShortcut)

$btnDeleteSaved = New-Object System.Windows.Forms.Button
$btnDeleteSaved.Text = "Excluir"
$btnDeleteSaved.Location = New-Object System.Drawing.Point(236, 454)
$btnDeleteSaved.Size = New-Object System.Drawing.Size(100, 26)
$btnDeleteSaved.FlatStyle = "Flat"
$btnDeleteSaved.BackColor = $cSurface
$btnDeleteSaved.ForeColor = $cRed
$btnDeleteSaved.FlatAppearance.BorderColor = $cRed
$btnDeleteSaved.Font = New-Object System.Drawing.Font("Segoe UI", 8)
$btnDeleteSaved.Cursor = [System.Windows.Forms.Cursors]::Hand
$form.Controls.Add($btnDeleteSaved)

$lblTools = New-Object System.Windows.Forms.Label
$lblTools.Text = "FERRAMENTAS"
$lblTools.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
$lblTools.ForeColor = $cMuted
$lblTools.Location = New-Object System.Drawing.Point(270, 68)
$lblTools.Size = New-Object System.Drawing.Size(150, 18)
$form.Controls.Add($lblTools)

$btnSnapshot = New-Object System.Windows.Forms.Button
$btnSnapshot.Text = "CAPTURAR LAYOUT"
$btnSnapshot.Location = New-Object System.Drawing.Point(270, 88)
$btnSnapshot.Size = New-Object System.Drawing.Size(280, 44)
$btnSnapshot.FlatStyle = "Flat"
$btnSnapshot.BackColor = $cSurface
$btnSnapshot.ForeColor = $cText
$btnSnapshot.FlatAppearance.BorderColor = $cBorder
$btnSnapshot.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$btnSnapshot.Cursor = [System.Windows.Forms.Cursors]::Hand
$form.Controls.Add($btnSnapshot)

$lblPreview = New-Object System.Windows.Forms.Label
$lblPreview.Text = "👁️ PRÉ-VISUALIZAÇÃO"
$lblPreview.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
$lblPreview.ForeColor = $cMuted
$lblPreview.Location = New-Object System.Drawing.Point(270, 142)
$lblPreview.Size = New-Object System.Drawing.Size(150, 18)
$form.Controls.Add($lblPreview)

$pnlPreview = New-Object System.Windows.Forms.Panel
$pnlPreview.Location = New-Object System.Drawing.Point(270, 162)
$pnlPreview.Size = New-Object System.Drawing.Size(520, 330)
$pnlPreview.BackColor = $cSurface
$pnlPreview.BorderStyle = "None"
$form.Controls.Add($pnlPreview)

$lblSpaces = New-Object System.Windows.Forms.Label
$lblSpaces.Text = "SPACES & LAYERS"
$lblSpaces.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
$lblSpaces.ForeColor = $cMuted
$lblSpaces.Location = New-Object System.Drawing.Point(800, 68)
$lblSpaces.Size = New-Object System.Drawing.Size(150, 18)
$form.Controls.Add($lblSpaces)

$pnlSpaces = New-Object System.Windows.Forms.Panel
$pnlSpaces.Location = New-Object System.Drawing.Point(800, 88)
$pnlSpaces.Size = New-Object System.Drawing.Size(280, 470)
$pnlSpaces.BackColor = $cSurface
$pnlSpaces.AutoScroll = $true
$form.Controls.Add($pnlSpaces)

$btnApply = New-Object System.Windows.Forms.Button
$btnApply.Text = "APLICAR LAYOUT"
$btnApply.Location = New-Object System.Drawing.Point(270, 506)
$btnApply.Size = New-Object System.Drawing.Size(520, 48)
$btnApply.FlatStyle = "Flat"
$btnApply.BackColor = $cAccent
$btnApply.ForeColor = [System.Drawing.Color]::White
$btnApply.FlatAppearance.BorderSize = 0
$btnApply.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
$btnApply.Cursor = [System.Windows.Forms.Cursors]::Hand
$form.Controls.Add($btnApply)

$btnSaveCurrent = New-Object System.Windows.Forms.Button
$btnSaveCurrent.Text = "Salvar Como..."
$btnSaveCurrent.Location = New-Object System.Drawing.Point(270, 560)
$btnSaveCurrent.Size = New-Object System.Drawing.Size(180, 32)
$btnSaveCurrent.FlatStyle = "Flat"
$btnSaveCurrent.BackColor = $cSurface
$btnSaveCurrent.ForeColor = $cAccent
$btnSaveCurrent.FlatAppearance.BorderColor = $cAccent
$btnSaveCurrent.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$btnSaveCurrent.Cursor = [System.Windows.Forms.Cursors]::Hand
$form.Controls.Add($btnSaveCurrent)

$lblRes = New-Object System.Windows.Forms.Label
$lblRes.Text = "$SW x $SH px"
$lblRes.ForeColor = $cMuted
$lblRes.Font = New-Object System.Drawing.Font("Segoe UI", 8)
$lblRes.Location = New-Object System.Drawing.Point(272, 494)
$lblRes.Size = New-Object System.Drawing.Size(140, 14)
$form.Controls.Add($lblRes)

$lblActive = New-Object System.Windows.Forms.Label
$lblActive.Text = ""
$lblActive.ForeColor = $cGreen
$lblActive.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
$lblActive.Location = New-Object System.Drawing.Point(272, 506)
$lblActive.Size = New-Object System.Drawing.Size(240, 18)
$form.Controls.Add($lblActive)

$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Text = "Pronto. Selecione um layout ou use 'Capturar Layout Atual'."
$lblStatus.ForeColor = $cMuted
$lblStatus.Font = New-Object System.Drawing.Font("Segoe UI", 8)
$lblStatus.Location = New-Object System.Drawing.Point(16, 580)
$lblStatus.Size = New-Object System.Drawing.Size(1040, 20)
$form.Controls.Add($lblStatus)

$script:currentName = ""
$script:highlightedSpaceIndex = -1
$script:activeLayoutName = ""

function Refresh-SavedList {
    $lstSaved.Items.Clear()
    foreach ($name in $script:savedLayouts.Keys) {
        $sc = $script:savedLayouts[$name].Shortcut
        $active = if ($name -eq $script:activeLayoutName) { " (*)" } else { "" }
        [void]$lstSaved.Items.Add((if ($sc) { "$name [$sc]$active" } else { "$name$active" }))
    }
}

function Update-ActiveLabel {
    $lblActive.Text = if ($script:activeLayoutName) { " [*] Ativo: $($script:activeLayoutName)" } else { "" }
}

$pnlPreview.add_Paint({
    param($s, $e)
    $g = $e.Graphics
    $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $g.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::ClearTypeGridFit
    $g.Clear($cSurface)
    if ($script:currentSpaces.Count -eq 0) {
        $f = New-Object System.Drawing.Font("Segoe UI", 12)
        $sf = New-Object System.Drawing.StringFormat
        $sf.Alignment = [System.Drawing.StringAlignment]::Center
        $sf.LineAlignment = [System.Drawing.StringAlignment]::Center
        $rect = [System.Drawing.RectangleF]::new(0, 0, $pnlPreview.Width, $pnlPreview.Height)
        $g.DrawString("Selecione um layout ou use 'Capturar Layout Atual'" , $f, [System.Drawing.Brushes]::Gray, $rect, $sf)
        return
    }
    $pw = $pnlPreview.Width - 8; $ph = $pnlPreview.Height - 8; $i = 0
    foreach ($space in $script:currentSpaces) {
        $zp = $space.Zone
        $rx = 4 + [int]($pw * $zp[0] / 100)
        $ry = 4 + [int]($ph * $zp[1] / 100)
        $rw = [int]($pw * $zp[2] / 100) - 4
        $rh = [int]($ph * $zp[3] / 100) - 4
        $ci = $i % $script:spaceColors.Count
        $fillColor = $script:spaceColors[$ci].Fill
        $strokeColor = $script:spaceColors[$ci].Stroke
        if ($i -eq $script:highlightedSpaceIndex) {
            $fillColor = [System.Drawing.Color]::FromArgb(100, $strokeColor.R, $strokeColor.G, $strokeColor.B)
        }
        $brush = New-Object System.Drawing.SolidBrush($fillColor)
        $penW = if ($i -eq $script:highlightedSpaceIndex) { 3 } else { 2 }
        $pen = New-Object System.Drawing.Pen($strokeColor, $penW)
        $radius = 8
        $path = New-Object System.Drawing.Drawing2D.GraphicsPath
        $path.AddArc($rx, $ry, $radius, $radius, 180, 90)
        $path.AddArc($rx + $rw - $radius, $ry, $radius, $radius, 270, 90)
        $path.AddArc($rx + $rw - $radius, $ry + $rh - $radius, $radius, $radius, 0, 90)
        $path.AddArc($rx, $ry + $rh - $radius, $radius, $radius, 90, 90)
        $path.CloseFigure()
        $g.FillPath($brush, $path)
        $g.DrawPath($pen, $path)
        if ($i -eq $script:highlightedSpaceIndex) {
            $glowPen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(150, 255, 255, 150), 2)
            $g.DrawPath($glowPen, $path)
        }
        $fontBold = New-Object System.Drawing.Font("Segoe UI Semibold", 10)
        $fontSmall = New-Object System.Drawing.Font("Segoe UI", 8)
        $textBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::White)
        $mutedBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(220, 220, 235))
        $g.DrawString(" " + $space.Name, $fontBold, $textBrush, $rx + 10, $ry + 8)
        if ($space.Shortcut) {
            $badgeFont = New-Object System.Drawing.Font("Segoe UI Semibold", 7.5)
            $badgeText = $space.Shortcut
            $badgeSize = $g.MeasureString($badgeText, $badgeFont)
            $bg = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(90, 30, 30, 40))
            $g.FillRectangle($bg, $rx + $rw - $badgeSize.Width - 10, $ry + 6, $badgeSize.Width + 14, $badgeSize.Height + 4)
            $g.DrawString($badgeText, $badgeFont, [System.Drawing.Brushes]::WhiteSmoke, $rx + $rw - $badgeSize.Width - 8, $ry + 7)
        }
        $ly = $ry + 30
        foreach ($layer in $space.Layers) {
            if ($ly + 16 -gt $ry + $rh - 8) { $g.DrawString("...", $fontSmall, $mutedBrush, $rx + 12, $ly); break }
            $lt = if ($layer.Title.Length -gt 28) { $layer.Title.Substring(0, 25) + "..." } else { $layer.Title }
            $g.DrawString("[L] $lt", $fontSmall, $mutedBrush, $rx + 12, $ly)
            $ly += 16
        }
        $i++
    }
    $borderPen = New-Object System.Drawing.Pen($cBorder, 1)
    $g.DrawRectangle($borderPen, 0, 0, $pnlPreview.Width - 1, $pnlPreview.Height - 1)
})

$pnlPreview.add_MouseClick({
    param($s, $e)
    if ($script:currentSpaces.Count -eq 0) { return }
    $pw = $pnlPreview.Width - 8; $ph = $pnlPreview.Height - 8
    for ($i = 0; $i -lt $script:currentSpaces.Count; $i++) {
        $zp = $script:currentSpaces[$i].Zone
        $rx = 4 + [int]($pw * $zp[0] / 100)
        $ry = 4 + [int]($ph * $zp[1] / 100)
        $rw = [int]($pw * $zp[2] / 100) - 4
        $rh = [int]($ph * $zp[3] / 100) - 4
        if ($e.X -ge $rx -and $e.X -le $rx + $rw -and $e.Y -ge $ry -and $e.Y -le $ry + $rh) {
            $script:highlightedSpaceIndex = if ($script:highlightedSpaceIndex -eq $i) { -1 } else { $i }
            $pnlPreview.Invalidate()
            return
        }
    }
    $script:highlightedSpaceIndex = -1
    $pnlPreview.Invalidate()
})

function Build-SpacePanel {
    $pnlSpaces.Controls.Clear()
    $y = 6; $i = 0
    foreach ($space in $script:currentSpaces) {
        $ci = $i % $script:spaceColors.Count
        $strokeColor = $script:spaceColors[$ci].Stroke
        $pnlHeader = New-Object System.Windows.Forms.Panel
        $pnlHeader.Location = New-Object System.Drawing.Point(4, $y)
        $pnlHeader.Size = New-Object System.Drawing.Size(270, 34)
        $pnlHeader.BackColor = $cSurface2
        $pnlHeader.Cursor = [System.Windows.Forms.Cursors]::Hand
        $pnlHeader.Tag = $i
        $pnlHeader.add_Click({ param($s, $e)
            if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Left) {
                $script:highlightedSpaceIndex = $s.Tag
                $pnlPreview.Invalidate()
            }
        })
        $pnlSpaces.Controls.Add($pnlHeader)
        $colorBar = New-Object System.Windows.Forms.Panel
        $colorBar.Location = New-Object System.Drawing.Point(0, 0)
        $colorBar.Size = New-Object System.Drawing.Size(4, 34)
        $colorBar.BackColor = $strokeColor
        $pnlHeader.Controls.Add($colorBar)
        $lblName = New-Object System.Windows.Forms.Label
        $lblName.Text = $space.Name
        $lblName.ForeColor = $cText
        $lblName.Font = New-Object System.Drawing.Font("Segoe UI Semibold", 9)
        $lblName.Location = New-Object System.Drawing.Point(10, 8)
        $lblName.Size = New-Object System.Drawing.Size(95, 20)
        $pnlHeader.Controls.Add($lblName)
        $btnSpaceSC = New-Object System.Windows.Forms.Button
        $scText = if ($space.Shortcut) { $space.Shortcut } else { "⌨️" }
        $btnSpaceSC.Text = $scText
        $btnSpaceSC.Location = New-Object System.Drawing.Point(74, 6)
        $btnSpaceSC.Size = New-Object System.Drawing.Size(34, 22)
        $btnSpaceSC.FlatStyle = "Flat"
        $btnSpaceSC.BackColor = $cSurface2
        $btnSpaceSC.ForeColor = if ($space.Shortcut) { $cAccent } else { $cMuted }
        $btnSpaceSC.FlatAppearance.BorderColor = if ($space.Shortcut) { $cAccent } else { $cBorder }
        $btnSpaceSC.Font = New-Object System.Drawing.Font("Segoe UI", 6.5)
        $btnSpaceSC.Tag = $i
        $btnSpaceSC.Cursor = [System.Windows.Forms.Cursors]::Hand
        $btnSpaceSC.add_Click({ param($s, $ea)
            $spIdx = $s.Tag; $space = $script:currentSpaces[$spIdx]
            $dlg = New-Object System.Windows.Forms.Form
            $dlg.Text = "Atalho - " + $space.Name
            $dlg.Size = New-Object System.Drawing.Size(360, 170)
            $dlg.StartPosition = "CenterParent"
            $dlg.FormBorderStyle = "FixedToolWindow"
            $dlg.BackColor = $cBg; $dlg.ForeColor = $cText
            $dlg.KeyPreview = $true
            $lblInfo = New-Object System.Windows.Forms.Label
            $lblInfo.Text = "Pressione a combinacao de teclas para aplicar este space na janela ativa:"
            $lblInfo.Location = New-Object System.Drawing.Point(12, 12)
            $lblInfo.Size = New-Object System.Drawing.Size(340, 18)
            $dlg.Controls.Add($lblInfo)
            $lblHint = New-Object System.Windows.Forms.Label
            $lblHint.Text = "O atalho encaixa a janela ativa neste space"
            $lblHint.ForeColor = $cMuted
            $lblHint.Font = New-Object System.Drawing.Font("Segoe UI", 7.5)
            $lblHint.Location = New-Object System.Drawing.Point(12, 32)
            $lblHint.Size = New-Object System.Drawing.Size(320, 16)
            $dlg.Controls.Add($lblHint)
            $txtKey = New-Object System.Windows.Forms.TextBox
$txtKey.Location = New-Object System.Drawing.Point(12, 54)
$txtKey.Size = New-Object System.Drawing.Size(320, 28)
$txtKey.BackColor = $cSurface
$txtKey.ForeColor = $cAccent
$txtKey.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
$txtKey.ReadOnly = $true
$txtKey.Text = $space.Shortcut
$dlg.Controls.Add($txtKey)
            $btnOk = New-Object System.Windows.Forms.Button
$btnOk.Text = "Salvar"
$btnOk.Location = New-Object System.Drawing.Point(12, 92)
$btnOk.Size = New-Object System.Drawing.Size(90, 32)
$btnOk.FlatStyle = "Flat"
$btnOk.BackColor = $cAccent
$btnOk.ForeColor = [System.Drawing.Color]::White
$btnOk.DialogResult = [System.Windows.Forms.DialogResult]::OK
$dlg.Controls.Add($btnOk)
            $btnClear = New-Object System.Windows.Forms.Button
$btnClear.Text = "Limpar"
$btnClear.Location = New-Object System.Drawing.Point(110, 92)
$btnClear.Size = New-Object System.Drawing.Size(90, 32)
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
    if ($parts.Count -gt 0) { $txtKey.Text = ($parts -join "+") }
})
if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
    $space.Shortcut = $txtKey.Text
    Register-SpaceHotkeys
    Build-SpacePanel
    $pnlPreview.Invalidate()
    if ($txtKey.Text) {
        $lblStatus.Text = "Atalho '$($txtKey.Text)' definido para '$($space.Name)'."
    } else {
        $lblStatus.Text = "Atalho removido de '$($space.Name)'."
    }
    $lblStatus.ForeColor = $cGreen
    }
})
$pnlHeader.Controls.Add($btnSpaceSC)

$y += 38

$layerIdx = 0
foreach ($layer in $space.Layers) {
    $lt = if ($layer.Title.Length -gt 26) { $layer.Title.Substring(0, 23) + "..." } else { $layer.Title }
    $isDynamic = $script:dynamicLayers | Where-Object { $_.Handle -eq $layer.Handle }
    $layerColor = if ($isDynamic) { $cOrange } else { $cMuted }

    $lblLayer = New-Object System.Windows.Forms.Label
    $lblLayer.Text = "  $($layerIdx + 1). $lt"
    $lblLayer.ForeColor = $layerColor
    $lblLayer.Font = New-Object System.Drawing.Font("Segoe UI", 7.5)
    $lblLayer.Location = New-Object System.Drawing.Point(10, $y)
    $lblLayer.Size = New-Object System.Drawing.Size(220, 16)
    $pnlSpaces.Controls.Add($lblLayer)

    $btnRemove = New-Object System.Windows.Forms.Button
    $btnRemove.Text = "×"
    $btnRemove.Location = New-Object System.Drawing.Point(242, ($y - 2))
    $btnRemove.Size = New-Object System.Drawing.Size(22, 18)
    $btnRemove.FlatStyle = "Flat"
    $btnRemove.BackColor = $cSurface
    $btnRemove.ForeColor = $cRed
    $btnRemove.FlatAppearance.BorderSize = 0
    $btnRemove.Font = New-Object System.Drawing.Font("Segoe UI", 7)
    $btnRemove.Tag = @{ SpaceIdx = $i; LayerIdx = $layerIdx }
    $btnRemove.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btnRemove.add_Click({
        param($sender, $ea)
        $tag = $sender.Tag
        $sp = $script:currentSpaces[$tag.SpaceIdx]
        if ($tag.LayerIdx -lt $sp.Layers.Count) {
            $removedLayer = $sp.Layers[$tag.LayerIdx]
            $sp.Layers.RemoveAt($tag.LayerIdx)
            $script:dynamicLayers = @($script:dynamicLayers | Where-Object { $_.Handle -ne $removedLayer.Handle })
        }
        Build-SpacePanel
        $pnlPreview.Invalidate()
    })
    $pnlSpaces.Controls.Add($btnRemove)
    $y += 18
    $layerIdx++
}

$btnRenameSpace = New-Object System.Windows.Forms.Button
$btnRenameSpace.Text = "Renomear"
$btnRenameSpace.Location = New-Object System.Drawing.Point(10, $y)
$btnRenameSpace.Size = New-Object System.Drawing.Size(80, 22)
$btnRenameSpace.FlatStyle = "Flat"
$btnRenameSpace.BackColor = $cSurface
$btnRenameSpace.ForeColor = $cAccent
$btnRenameSpace.FlatAppearance.BorderColor = $cAccent
$btnRenameSpace.Font = New-Object System.Drawing.Font("Segoe UI", 8)
$btnRenameSpace.Tag = $i
$btnRenameSpace.Cursor = [System.Windows.Forms.Cursors]::Hand
$btnRenameSpace.add_Click({ param($s, $ea)
    $spIdx = $s.Tag; $space = $script:currentSpaces[$spIdx]
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = "Renomear Space"
    $dlg.Size = New-Object System.Drawing.Size(320, 140)
    $dlg.StartPosition = "CenterParent"
    $dlg.FormBorderStyle = "FixedToolWindow"
    $dlg.BackColor = $cBg; $dlg.ForeColor = $cText
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = "Novo nome:"; $lbl.Location = New-Object System.Drawing.Point(12, 16); $lbl.Size = New-Object System.Drawing.Size(280, 20)
    $dlg.Controls.Add($lbl)
    $txt = New-Object System.Windows.Forms.TextBox
    $txt.Location = New-Object System.Drawing.Point(12, 42); $txt.Size = New-Object System.Drawing.Size(280, 28)
    $txt.BackColor = $cSurface; $txt.ForeColor = $cText; $txt.Text = $space.Name
    $dlg.Controls.Add($txt)
    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text = "OK"; $btnOk.Location = New-Object System.Drawing.Point(12, 78)
    $btnOk.Size = New-Object System.Drawing.Size(80, 32)
    $btnOk.FlatStyle = "Flat"; $btnOk.BackColor = $cAccent; $btnOk.ForeColor = [System.Drawing.Color]::White
    $btnOk.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $dlg.Controls.Add($btnOk)
    $dlg.AcceptButton = $btnOk
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK -and $txt.Text.Trim()) {
        $space.Name = $txt.Text.Trim()
        $lblStatus.Text = "Space renomeado para '" + $space.Name + "'."
        $lblStatus.ForeColor = $cGreen
        Build-SpacePanel; $pnlPreview.Invalidate()
    }
})
$pnlSpaces.Controls.Add($btnRenameSpace)
$y += 28

$btnAdd = New-Object System.Windows.Forms.Button
$btnAdd.Text = "Add Layer"
$btnAdd.Location = New-Object System.Drawing.Point(10, $y)
$btnAdd.Size = New-Object System.Drawing.Size(80, 22)
$btnAdd.FlatStyle = "Flat"
$btnAdd.BackColor = $cSurface
$btnAdd.ForeColor = $cGreen
$btnAdd.FlatAppearance.BorderColor = $cGreen
$btnAdd.Font = New-Object System.Drawing.Font("Segoe UI", 8)
$btnAdd.Tag = $i
$btnAdd.Cursor = [System.Windows.Forms.Cursors]::Hand
$btnAdd.add_Click({
    param($sender, $ea)
    $spIdx = $sender.Tag
    $popup = New-Object System.Windows.Forms.Form
    $popup.Text = "Selecionar Janela"
    $popup.Size = New-Object System.Drawing.Size(450, 400)
    $popup.StartPosition = "CenterParent"
    $popup.FormBorderStyle = "FixedToolWindow"
    $popup.BackColor = $cBg
    $popup.ForeColor = $cText

    $lstPopup = New-Object System.Windows.Forms.ListBox
    $lstPopup.Location = New-Object System.Drawing.Point(10, 12)
    $lstPopup.Size = New-Object System.Drawing.Size(420, 300)
    $lstPopup.BackColor = $cSurface
    $lstPopup.ForeColor = $cText
    $lstPopup.BorderStyle = "None"
    $lstPopup.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $popup.Controls.Add($lstPopup)

    $popupHandles = @()
    $wins = Get-VisibleWindows | Where-Object { $_.Title -notlike "*SnapLayout*" -and $_.Title.Length -gt 1 }
    foreach ($w in $wins) {
        $short = if ($w.Title.Length -gt 58) { $w.Title.Substring(0, 55) + "..." } else { $w.Title }
        [void]$lstPopup.Items.Add($short)
        $popupHandles += $w
    }

                $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text = "Adicionar"
    $btnOk.Location = New-Object System.Drawing.Point(10, 320)
    $btnOk.Size = New-Object System.Drawing.Size(120, 32)
    $btnOk.FlatStyle = "Flat"
    $btnOk.BackColor = $cAccent
    $btnOk.ForeColor = [System.Drawing.Color]::White
    $btnOk.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $popup.Controls.Add($btnOk)
    $popup.AcceptButton = $btnOk

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Cancelar"
    $btnCancel.Location = New-Object System.Drawing.Point(140, 320)
    $btnCancel.Size = New-Object System.Drawing.Size(100, 32)
    $btnCancel.FlatStyle = "Flat"
    $btnCancel.BackColor = $cSurface
    $btnCancel.ForeColor = $cMuted
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $popup.Controls.Add($btnCancel)
    $popup.CancelButton = $btnCancel

    if ($popup.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK -and $lstPopup.SelectedIndex -ge 0) {
        $selWin = $popupHandles[$lstPopup.SelectedIndex]
        $space = $script:currentSpaces[$spIdx]
        [void]$space.Layers.Add(@{ Handle = $selWin.Handle; Title = $selWin.Title })
        Build-SpacePanel
        $pnlPreview.Invalidate()
    }
})
$pnlSpaces.Controls.Add($btnAdd)

$btnReset = New-Object System.Windows.Forms.Button
$btnReset.Text = "❌ Reset"
$btnReset.Location = New-Object System.Drawing.Point(106, $y)
$btnReset.Size = New-Object System.Drawing.Size(75, 24)
$btnReset.FlatStyle = "Flat"
$btnReset.BackColor = $cSurface
$btnReset.ForeColor = $cRed
$btnReset.FlatAppearance.BorderColor = $cRed
$btnReset.Font = New-Object System.Drawing.Font("Segoe UI", 8)
$btnReset.Tag = $i
$btnReset.Cursor = [System.Windows.Forms.Cursors]::Hand
$btnReset.add_Click({
    param($sender, $ea)
    $spIdx = $sender.Tag
    $space = $script:currentSpaces[$spIdx]
    foreach ($l in $space.Layers) {
        $script:dynamicLayers = @($script:dynamicLayers | Where-Object { $_.Handle -ne $l.Handle })
    }
    $space.Layers.Clear()
    Build-SpacePanel
    $pnlPreview.Invalidate()
    $lblStatus.Text = "Layers de '$($space.Name)' limpas."
    $lblStatus.ForeColor = $cMuted
})
$pnlSpaces.Controls.Add($btnReset)

$y += 32

$sepSpace = New-Object System.Windows.Forms.Panel
$sepSpace.Location = New-Object System.Drawing.Point(4, $y)
$sepSpace.Size = New-Object System.Drawing.Size(268, 1)
$sepSpace.BackColor = $cBorder
$pnlSpaces.Controls.Add($sepSpace)
$y += 6
$i++
    }
}

function Select-Layout($name, $source) {
    $script:currentName = $name
    $script:highlightedSpaceIndex = -1
    $saved = $script:savedLayouts[$name]
    $script:currentSpaces = @()
    foreach ($s in $saved.Spaces) {
        $sp = @{ Name = $s.Name; Zone = @($s.Zone); Layers = [System.Collections.ArrayList]::new(); Shortcut = if ($s.Shortcut) { $s.Shortcut } else { "" } }
        foreach ($l in $s.Layers) {
            [void]$sp.Layers.Add(@{ Handle = $l.Handle; Title = $l.Title })
        }
        $script:currentSpaces += $sp
    }
    Build-SpacePanel
    $pnlPreview.Invalidate()
    $lblStatus.Text = "Layout: $name"
    $lblStatus.ForeColor = $cMuted
}

$lstSaved.add_SelectedIndexChanged({
    $sel = $lstSaved.SelectedItem
    if ($sel) {
        $name = ($sel -replace '\s*\[.*$', '').Trim()
        $name = ($name -replace '\s*$', '').Trim()
        if ($script:savedLayouts.Contains($name)) {
            Select-Layout $name "saved"
        }
    }
})

$btnApply.add_Click({
    if ($script:currentSpaces.Count -eq 0) {
        $lblStatus.Text = "Nenhum layout selecionado."
        $lblStatus.ForeColor = $cOrange
        return
    }
    $applied = 0
    $gap = $script:windowGap
    foreach ($space in $script:currentSpaces) {
        $z = ConvertZone $space.Zone
        $layerCount = $space.Layers.Count
        if ($layerCount -le 0) { continue }
        $sx = $z.X + $gap; $sy = $z.Y + $gap; $sw2 = $z.W - ($gap * 2); $sh2 = $z.H - ($gap * 2)
        foreach ($i in 0..($layerCount - 1)) {
            $layer = $space.Layers[$i]
            try { [WinAPI]::MoveWindow($layer.Handle, [int]$sx, [int]$sy, [int]$sw2, [int]$sh2); $applied++ } catch {}
        }
        if ($space.Layers.Count -gt 0) {
            $topLayer = $space.Layers[$space.Layers.Count - 1]
            [WinAPI]::SetForegroundWindow($topLayer.Handle)
        }
    }
    $script:activeLayoutName = $script:currentName
    Update-ActiveLabel
    Register-SpaceHotkeys
    Refresh-SavedList
    $lblStatus.Text = " Layout aplicado: $applied janela(s). Atalhos de space ativos."
    $lblStatus.ForeColor = $cGreen
})

$btnSaveCurrent.add_Click({
    if ($script:currentSpaces.Count -eq 0) {
        $lblStatus.Text = "Nenhum layout para salvar."
        $lblStatus.ForeColor = $cOrange
        return
    }
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = "Salvar Layout"
    $dlg.Size = New-Object System.Drawing.Size(350, 150)
    $dlg.StartPosition = "CenterParent"
    $dlg.FormBorderStyle = "FixedToolWindow"
    $dlg.BackColor = $cBg
    $dlg.ForeColor = $cText
    $lblN = New-Object System.Windows.Forms.Label
    $lblN.Text = "Nome do layout:"
    $lblN.Location = New-Object System.Drawing.Point(10, 15)
    $lblN.Size = New-Object System.Drawing.Size(320, 18)
    $dlg.Controls.Add($lblN)
    $txtName = New-Object System.Windows.Forms.TextBox
    $txtName.Location = New-Object System.Drawing.Point(10, 38)
    $txtName.Size = New-Object System.Drawing.Size(315, 24)
    $txtName.BackColor = $cSurface
    $txtName.ForeColor = $cText
    $txtName.Text = $script:currentName
    $dlg.Controls.Add($txtName)
                $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text = "Salvar"
    $btnOk.Location = New-Object System.Drawing.Point(10, 72)
    $btnOk.Size = New-Object System.Drawing.Size(100, 30)
    $btnOk.FlatStyle = "Flat"
    $btnOk.BackColor = $cAccent
    $btnOk.ForeColor = [System.Drawing.Color]::White
    $btnOk.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $dlg.Controls.Add($btnOk)
    $dlg.AcceptButton = $btnOk
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK -and $txtName.Text.Trim()) {
        $saveName = $txtName.Text.Trim()
        $spaces = @()
        foreach ($s in $script:currentSpaces) {
            $sp = @{ Name = $s.Name; Zone = @($s.Zone); Layers = [System.Collections.ArrayList]::new(); Shortcut = if ($s.Shortcut) { $s.Shortcut } else { "" } }
            foreach ($l in $s.Layers) {
                [void]$sp.Layers.Add(@{ Handle = $l.Handle; Title = $l.Title })
            }
            $spaces += $sp
        }
        $existingShortcut = ""
        if ($script:savedLayouts.Contains($saveName)) { $existingShortcut = $script:savedLayouts[$saveName].Shortcut }
        $script:savedLayouts[$saveName] = @{ Spaces = $spaces; Shortcut = $existingShortcut }
        Save-AllLayouts
        Refresh-SavedList
        $lblStatus.Text = " Layout '$saveName' salvo."
        $lblStatus.ForeColor = $cGreen
    }
})

$btnDeleteSaved.add_Click({
    $sel = $lstSaved.SelectedItem
    if (-not $sel) { return }
    $name = ($sel -replace '\s*\[.*$', '').Trim()
    $name = ($name -replace '\s*$', '').Trim()
    $confirm = [System.Windows.Forms.MessageBox]::Show("Excluir o layout '$name'?", "Confirmar", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
    if ($confirm -eq [System.Windows.Forms.DialogResult]::Yes) {
        $script:savedLayouts.Remove($name)
        if ($script:activeLayoutName -eq $name) {
            $script:activeLayoutName = ""
            Update-ActiveLabel
        }
        Save-AllLayouts
        Refresh-SavedList
        if ($script:currentName -eq $name) {
            $script:currentSpaces = @()
            Build-SpacePanel
            $pnlPreview.Invalidate()
        }
        $lblStatus.Text = "Layout '$name' excluido."
        $lblStatus.ForeColor = $cMuted
    }
})

$btnRenameLayout.add_Click({
    $sel = $lstSaved.SelectedItem
    if (-not $sel) {
        $lblStatus.Text = "Selecione um layout salvo primeiro."
        $lblStatus.ForeColor = $cOrange
        return
    }
    $name = ($sel -replace '\s*\[.*$', '').Trim()
    $name = ($name -replace '\s*$', '').Trim()
    $savedLayout = $script:savedLayouts[$name]
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = "Renomear Layout"
    $dlg.Size = New-Object System.Drawing.Size(350, 150)
    $dlg.StartPosition = "CenterParent"
    $dlg.FormBorderStyle = "FixedToolWindow"
    $dlg.BackColor = $cBg
    $dlg.ForeColor = $cText
    $lblN = New-Object System.Windows.Forms.Label
    $lblN.Text = "Novo nome:"
    $lblN.Location = New-Object System.Drawing.Point(10, 15)
    $lblN.Size = New-Object System.Drawing.Size(320, 18)
    $dlg.Controls.Add($lblN)
    $txtName = New-Object System.Windows.Forms.TextBox
    $txtName.Location = New-Object System.Drawing.Point(10, 38)
    $txtName.Size = New-Object System.Drawing.Size(315, 24)
    $txtName.BackColor = $cSurface
    $txtName.ForeColor = $cText
    $txtName.Text = $name
    $dlg.Controls.Add($txtName)
                $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text = "OK"
    $btnOk.Location = New-Object System.Drawing.Point(10, 72)
    $btnOk.Size = New-Object System.Drawing.Size(100, 32)
    $btnOk.FlatStyle = "Flat"
    $btnOk.BackColor = $cAccent
    $btnOk.ForeColor = [System.Drawing.Color]::White
    $btnOk.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $dlg.Controls.Add($btnOk)
    $dlg.AcceptButton = $btnOk
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK -and $txtName.Text.Trim() -and $txtName.Text.Trim() -ne $name) {
        $newName = $txtName.Text.Trim()
        $spaces = $savedLayout.Spaces
        $shortcut = $savedLayout.Shortcut
        $script:savedLayouts.Remove($name)
        $script:savedLayouts[$newName] = @{ Spaces = $spaces; Shortcut = $shortcut }
        Save-AllLayouts
        Refresh-SavedList
        if ($script:currentName -eq $name) { $script:currentName = $newName }
        if ($script:activeLayoutName -eq $name) {
            $script:activeLayoutName = $newName
            Update-ActiveLabel
        }
        $lblStatus.Text = "Layout renomeado para '$newName'."
        $lblStatus.ForeColor = $cGreen
    }
})

$btnSetShortcut.add_Click({
    $sel = $lstSaved.SelectedItem
    if (-not $sel) {
        $lblStatus.Text = "Selecione um layout salvo primeiro."
        $lblStatus.ForeColor = $cOrange
        return
    }
    $name = ($sel -replace '\s*\[.*$', '').Trim()
    $name = ($name -replace '\s*$', '').Trim()
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = "Atalho - " + $name
    $dlg.Size = New-Object System.Drawing.Size(360, 170)
    $dlg.StartPosition = "CenterParent"
    $dlg.FormBorderStyle = "FixedToolWindow"
    $dlg.BackColor = $cBg
    $dlg.ForeColor = $cText
    $dlg.KeyPreview = $true
    $lblInfo = New-Object System.Windows.Forms.Label
    $lblInfo.Text = "Pressione a combinacao de teclas desejada:"
    $lblInfo.Location = New-Object System.Drawing.Point(10, 15)
    $lblInfo.Size = New-Object System.Drawing.Size(330, 18)
    $dlg.Controls.Add($lblInfo)
    $txtKey = New-Object System.Windows.Forms.TextBox
    $txtKey.Location = New-Object System.Drawing.Point(10, 40)
    $txtKey.Size = New-Object System.Drawing.Size(330, 24)
    $txtKey.BackColor = $cSurface
    $txtKey.ForeColor = $cAccent
    $txtKey.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
    $txtKey.ReadOnly = $true
    $txtKey.Text = $script:savedLayouts[$name].Shortcut
    $dlg.Controls.Add($txtKey)
                $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text = "Salvar"
    $btnOk.Location = New-Object System.Drawing.Point(10, 78)
    $btnOk.Size = New-Object System.Drawing.Size(80, 30)
    $btnOk.FlatStyle = "Flat"
    $btnOk.BackColor = $cAccent
    $btnOk.ForeColor = [System.Drawing.Color]::White
    $btnOk.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $dlg.Controls.Add($btnOk)
                $btnClear = New-Object System.Windows.Forms.Button
    $btnClear.Text = "Limpar"
    $btnClear.Location = New-Object System.Drawing.Point(100, 78)
    $btnClear.Size = New-Object System.Drawing.Size(80, 30)
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
        if ($parts.Count -gt 0) { $txtKey.Text = ($parts -join "+") }
    })
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $script:savedLayouts[$name].Shortcut = $txtKey.Text
        Save-AllLayouts
        Refresh-SavedList
        Register-Hotkeys
        $lblStatus.Text = "Atalho '$($txtKey.Text)' definido para '$name'."
        $lblStatus.ForeColor = $cGreen
    }
})

$btnSnapshot.add_Click({
    $form.WindowState = [System.Windows.Forms.FormWindowState]::Minimized
    Start-Sleep -Milliseconds 400
    $wins = Get-VisibleWindows | Where-Object { $_.Title -notlike "*SnapLayout*" -and $_.Title.Length -gt 1 }
    $form.WindowState = [System.Windows.Forms.FormWindowState]::Normal
    if ($wins.Count -eq 0) {
        $lblStatus.Text = "Nenhuma janela visivel encontrada."
        $lblStatus.ForeColor = $cOrange
        return
    }
    $script:currentSpaces = @()
    $spaceIdx = 1
    foreach ($w in $wins) {
        $rect = [WinAPI]::GetRect($w.Handle)
        $xp = [Math]::Max(0, [Math]::Min(100, [Math]::Round(($rect.Left - $SX) / $SW * 100)))
        $yp = [Math]::Max(0, [Math]::Min(100, [Math]::Round(($rect.Top - $SY) / $SH * 100)))
        $wp = [Math]::Max(5, [Math]::Min(100, [Math]::Round(($rect.Right - $rect.Left) / $SW * 100)))
        $hp = [Math]::Max(5, [Math]::Min(100, [Math]::Round(($rect.Bottom - $rect.Top) / $SH * 100)))
        $xp = [Math]::Round($xp / 5) * 5; $yp = [Math]::Round($yp / 5) * 5
        $wp = [Math]::Round($wp / 5) * 5; $hp = [Math]::Round($hp / 5) * 5
        if ($wp -lt 5) { $wp = 5 }; if ($hp -lt 5) { $hp = 5 }
        $sp = @{ Name = "Space $spaceIdx"; Zone = @([int]$xp, [int]$yp, [int]$wp, [int]$hp); Layers = [System.Collections.ArrayList]::new(); Shortcut = "" }
        [void]$sp.Layers.Add(@{ Handle = $w.Handle; Title = $w.Title })
        $script:currentSpaces += $sp
        $spaceIdx++
    }
    $script:currentName = "Snapshot"
    $lstSaved.ClearSelected()
    Build-SpacePanel
    $pnlPreview.Invalidate()
    $lblStatus.Text = " Snapshot: $($script:currentSpaces.Count) janela(s) capturada(s)."
    $lblStatus.ForeColor = $cGreen
})

$vkMap = @{
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

function Parse-Shortcut($shortcutStr) {
    if (-not $shortcutStr) { return $null }
    $parts = $shortcutStr -split '\+'
    $mods = @()
    $key = ""
    foreach ($p in $parts) {
        $p = $p.Trim()
        if ($p -eq "Ctrl" -or $p -eq "Control") { $mods += "Ctrl" }
        elseif ($p -eq "Alt") { $mods += "Alt" }
        elseif ($p -eq "Shift") { $mods += "Shift" }
        else { $key = $p }
    }
    if (-not $key) { return $null }
    $vk = $vkMap[$key]
    if (-not $vk) { return $null }
    return @{ Mods = $mods; VKey = $vk }
}

function Check-Modifiers($mods) {
    if ($mods -contains "Ctrl") {
        if (-not ([WinAPI]::GetAsyncKeyState(0x11) -band 0x8000)) { return $false }
    }
    if ($mods -contains "Alt") {
        if (-not ([WinAPI]::GetAsyncKeyState(0x12) -band 0x8000)) { return $false }
    }
    if ($mods -contains "Shift") {
        if (-not ([WinAPI]::GetAsyncKeyState(0x10) -band 0x8000)) { return $false }
    }
    return $true
}

function Register-Hotkeys {
    $script:hotkeyBindings = @()
    foreach ($name in $script:savedLayouts.Keys) {
        $sc = $script:savedLayouts[$name].Shortcut
        $parsed = Parse-Shortcut $sc
        if (-not $parsed) { continue }
        $script:hotkeyBindings += @{ LayoutName = $name; Mods = $parsed.Mods; VKey = $parsed.VKey }
    }
}

function Apply-LayoutByName($name) {
    $layout = $script:savedLayouts[$name]
    if (-not $layout) { return }
    $gap = $script:windowGap
    foreach ($space in $layout.Spaces) {
        $z = ConvertZone $space.Zone
        $layerCount = $space.Layers.Count
        if ($layerCount -le 0) { continue }
        $sx = $z.X + $gap; $sy = $z.Y + $gap; $sw2 = $z.W - ($gap * 2); $sh2 = $z.H - ($gap * 2)
        foreach ($i in 0..($layerCount - 1)) {
            $layer = $space.Layers[$i]
            try { [WinAPI]::MoveWindow($layer.Handle, [int]$sx, [int]$sy, [int]$sw2, [int]$sh2) } catch {}
        }
        if ($space.Layers.Count -gt 0) {
            $topLayer = $space.Layers[$space.Layers.Count - 1]
            [WinAPI]::SetForegroundWindow($topLayer.Handle)
        }
    }
    $script:activeLayoutName = $name
    Update-ActiveLabel
    Register-SpaceHotkeys
}

function Register-SpaceHotkeys {
    $script:spaceHotkeyBindings = @()
    if (-not $script:activeLayoutName) { return }
    $layout = $script:savedLayouts[$script:activeLayoutName]
    if (-not $layout) { return }
    for ($i = 0; $i -lt $layout.Spaces.Count; $i++) {
        $space = $layout.Spaces[$i]
        if (-not $space.Shortcut) { continue }
        $parsed = Parse-Shortcut $space.Shortcut
        if (-not $parsed) { continue }
        $script:spaceHotkeyBindings += @{ SpaceIndex = $i; SpaceName = $space.Name; Mods = $parsed.Mods; VKey = $parsed.VKey; Zone = $space.Zone }
    }
}

function Snap-WindowToSpace($spaceBinding) {
    $foreWin = [WinAPI]::GetForegroundWindow()
    $title = [WinAPI]::GetTitle($foreWin)
    if (-not $title -or $title -like "*SnapLayout*") { return }
    $layout = $script:savedLayouts[$script:activeLayoutName]
    if (-not $layout) { return }
    $space = $layout.Spaces[$spaceBinding.SpaceIndex]
    if (-not $space) { return }
    [void]$space.Layers.Add(@{ Handle = $foreWin; Title = $title })
    $script:dynamicLayers += @{ Handle = $foreWin; SpaceIndex = $spaceBinding.SpaceIndex; LayoutName = $script:activeLayoutName }
    $z = ConvertZone $space.Zone
    $gap = $script:windowGap
    $sx = $z.X + $gap; $sy = $z.Y + $gap; $sw2 = $z.W - ($gap * 2); $sh2 = $z.H - ($gap * 2)
    try { [WinAPI]::MoveWindow($foreWin, [int]$sx, [int]$sy, [int]$sw2, [int]$sh2) } catch {}
    if ($script:currentName -eq $script:activeLayoutName) {
        $script:currentSpaces = @()
        foreach ($s in $layout.Spaces) {
            $sp = @{ Name = $s.Name; Zone = @($s.Zone); Layers = [System.Collections.ArrayList]::new(); Shortcut = if ($s.Shortcut) { $s.Shortcut } else { "" } }
            foreach ($l in $s.Layers) {
                [void]$sp.Layers.Add(@{ Handle = $l.Handle; Title = $l.Title })
            }
            $script:currentSpaces += $sp
        }
        Build-SpacePanel
        $pnlPreview.Invalidate()
    }
    $short = if ($title.Length -gt 35) { $title.Substring(0, 32) + "..." } else { $title }
    $lblStatus.Text = " '$short' encaixado em '$($spaceBinding.SpaceName)'."
    $lblStatus.ForeColor = $cGreen
}

$hotkeyTimer = New-Object System.Windows.Forms.Timer
$hotkeyTimer.Interval = 150
$hotkeyTimer.add_Tick({
    if (-not $script:hotkeyCooldown) {
        foreach ($binding in $script:hotkeyBindings) {
            if (-not (Check-Modifiers $binding.Mods)) { continue }
            $keyState = [WinAPI]::GetAsyncKeyState($binding.VKey)
            if ($keyState -band 0x8000) {
                Apply-LayoutByName $binding.LayoutName
                $lblStatus.Text = " Layout '$($binding.LayoutName)' aplicado via atalho."
                $lblStatus.ForeColor = $cGreen
                $script:hotkeyCooldown = $true
                $cdTimer = New-Object System.Windows.Forms.Timer
                $cdTimer.Interval = 500
                $cdTimer.add_Tick({ $script:hotkeyCooldown = $false; $this.Stop(); $this.Dispose() })
                $cdTimer.Start()
                return
            }
        }
    }
    if (-not $script:spaceHotkeyCooldown -and $script:activeLayoutName) {
        foreach ($binding in $script:spaceHotkeyBindings) {
            if (-not (Check-Modifiers $binding.Mods)) { continue }
            $keyState = [WinAPI]::GetAsyncKeyState($binding.VKey)
            if ($keyState -band 0x8000) {
                Snap-WindowToSpace $binding
                $script:spaceHotkeyCooldown = $true
                $cdTimer2 = New-Object System.Windows.Forms.Timer
                $cdTimer2.Interval = 500
                $cdTimer2.add_Tick({ $script:spaceHotkeyCooldown = $false; $this.Stop(); $this.Dispose() })
                $cdTimer2.Start()
                return
            }
        }
    }
    if ($script:dynamicLayers.Count -gt 0 -and $script:activeLayoutName) {
        $layout = $script:savedLayouts[$script:activeLayoutName]
        if ($layout) {
            $needRefresh = $false
            $toRemove = @()
            foreach ($dl in $script:dynamicLayers) {
                if (-not [WinAPI]::IsWindow($dl.Handle)) {
                    if ($dl.LayoutName -eq $script:activeLayoutName) {
                        $space = $layout.Spaces[$dl.SpaceIndex]
                        if ($space) {
                            $layerToRemove = $space.Layers | Where-Object { $_.Handle -eq $dl.Handle }
                            if ($layerToRemove) {
                                $space.Layers.Remove($layerToRemove)
                                $needRefresh = $true
                            }
                        }
                    }
                    $toRemove += $dl
                }
            }
            if ($toRemove.Count -gt 0) {
                $script:dynamicLayers = @($script:dynamicLayers | Where-Object {
                    $keep = $true
                    foreach ($r in $toRemove) {
                        if ($_.Handle -eq $r.Handle) { $keep = $false; break }
                    }
                    $keep
                })
            }
            if ($needRefresh -and $script:currentName -eq $script:activeLayoutName) {
                $script:currentSpaces = @()
                foreach ($s in $layout.Spaces) {
                    $sp = @{ Name = $s.Name; Zone = @($s.Zone); Layers = [System.Collections.ArrayList]::new(); Shortcut = if ($s.Shortcut) { $s.Shortcut } else { "" } }
                    foreach ($l in $s.Layers) {
                        [void]$sp.Layers.Add(@{ Handle = $l.Handle; Title = $l.Title })
                    }
                    $script:currentSpaces += $sp
                }
                Build-SpacePanel
                $pnlPreview.Invalidate()
            }
        }
    }
})

Load-AllLayouts
Refresh-SavedList
Register-Hotkeys

[void]$form.ShowDialog()
$hotkeyTimer.Stop()
$hotkeyTimer.Dispose()
