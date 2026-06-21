# ============================================================
#  SnapLayouts - Gestor de janelas com Spaces, Layers e Canvas
#  Sem dependencias externas. Requer PowerShell 5+ no Windows 11.
#  Execucao: powershell -ExecutionPolicy Bypass -File layouts.ps1
# ============================================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

if (-not ([System.Management.Automation.PSTypeName]'SnapAPI').Type) { Add-Type @"
using System;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using System.Text;

public class SnapAPI {
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
    [DllImport("user32.dll")] public static extern bool ReleaseCapture();
    [DllImport("user32.dll")] public static extern IntPtr SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);
    [DllImport("dwmapi.dll")] public static extern int DwmGetWindowAttribute(
        IntPtr hwnd, int dwAttribute, out RECT pvAttribute, int cbAttribute);

    [StructLayout(LayoutKind.Sequential)]
    public struct RECT { public int Left, Top, Right, Bottom; }

    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    public const int SW_RESTORE        = 9;
    public const int DWMWA_EXTENDED_FRAME_BOUNDS = 9;
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

    public static List<IntPtr> GetAllApplicationWindows() {
        var list = new List<IntPtr>();
        EnumWindows((hWnd, _) => {
            if (!IsWindowVisible(hWnd)) return true;  // skip hidden, but include minimized
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
        // Compensar borda invisivel DWM (sombra/frame transparente do Windows 10/11)
        RECT wr, vr;
        if (GetWindowRect(hWnd, out wr) &&
            DwmGetWindowAttribute(hWnd, DWMWA_EXTENDED_FRAME_BOUNDS, out vr,
                System.Runtime.InteropServices.Marshal.SizeOf(typeof(RECT))) == 0) {
            int bl = vr.Left  - wr.Left;
            int bt = vr.Top   - wr.Top;
            int br = wr.Right  - vr.Right;
            int bb = wr.Bottom - vr.Bottom;
            SetWindowPos(hWnd, IntPtr.Zero, x - bl, y - bt, w + bl + br, h + bt + bb, SWP_SHOWWINDOW | SWP_NOZORDER);
        } else {
            SetWindowPos(hWnd, IntPtr.Zero, x, y, w, h, SWP_SHOWWINDOW | SWP_NOZORDER);
        }
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

$screen   = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
$SW       = $screen.Width
$SH       = $screen.Height
$SX       = $screen.X
$SY       = $screen.Y
$scrFull  = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds
# Tamanho inicial do form: preview com mesmo aspect ratio da tela
# preview.Width = formCW - 566 ; preview.Height = formCH - 189
$_previewH = 400
$_previewW = [int]($_previewH * $scrFull.Width / $scrFull.Height)
$_formCW   = $_previewW + 566
$_formCH   = $_previewH + 189
# Garante que nao ultrapasse 90% da area util
if ($_formCW -gt [int]($SW * 0.90)) {
    $_formCW   = [int]($SW * 0.90)
    $_previewW = $_formCW - 566
    $_previewH = [int]($_previewW * $scrFull.Height / $scrFull.Width)
    $_formCH   = $_previewH + 189
}
if ($_formCH -gt [int]($SH * 0.90)) { $_formCH = [int]($SH * 0.90) }
# Dimensoes fixas do canvas (nao mudam com resize)
$script:previewW = $_previewW
$script:previewH = $_previewH

$script:monitors      = @([System.Windows.Forms.Screen]::AllScreens | Sort-Object { $_.Bounds.X })
$script:previewMonitor = 0   # pagina do canvas atualmente visivel

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
$script:dragSpaceIdx      = -1
$script:dragEdge          = ""
$script:dragStartZone     = @()
$script:dragStartMousePct = @(0, 0)

# ============================================================
#  FUNCOES UTILITARIAS
# ============================================================

function ConvertZone($zone, $monIdx = 0) {
    $mon = if ($monIdx -ge 0 -and $monIdx -lt $script:monitors.Count) {
        $script:monitors[$monIdx].WorkingArea
    } else { $script:monitors[0].WorkingArea }
    return @{
        X = $mon.X + [int]($mon.Width  * $zone[0] / 100)
        Y = $mon.Y + [int]($mon.Height * $zone[1] / 100)
        W = [int]($mon.Width  * $zone[2] / 100)
        H = [int]($mon.Height * $zone[3] / 100)
    }
}

function Get-VisibleWindows {
    $handles = [SnapAPI]::GetVisibleWindows()
    $result  = @()
    foreach ($h in $handles) {
        # Verificar se janela está minimizada ou oculta
        if ([SnapAPI]::IsIconic($h) -or -not [SnapAPI]::IsWindowVisible($h)) {
            continue
        }
        $title = [SnapAPI]::GetTitle($h)
        if ($title -and $title.Length -gt 1) {
            $result += [PSCustomObject]@{ Handle = $h; Title = $title; Process = (Get-WindowProcess $h) }
        }
    }
    return $result | Sort-Object Title
}

function Get-AllWindowsForApply {
    $handles = [SnapAPI]::GetAllApplicationWindows()
    $result  = @()
    foreach ($h in $handles) {
        $title = [SnapAPI]::GetTitle($h)
        if ($title -and $title.Length -gt 1) {
            $result += [PSCustomObject]@{ Handle = $h; Title = $title; Process = (Get-WindowProcess $h) }
        }
    }
    return $result
}

function Get-WindowProcess($hwnd) {
    try {
        $pid2 = [uint32]0
        [void][SnapAPI]::GetWindowThreadProcessId($hwnd, [ref]$pid2)
        return (Get-Process -Id ([int]$pid2) -ErrorAction Stop).ProcessName
    } catch { return "" }
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
            $layersData = @()
            foreach ($l in $s.Layers) {
                $layersData += @{
                    title   = if ($l.Title)   { $l.Title }   else { "" }
                    process = if ($l.Process) { $l.Process } else { "" }
                    locked  = if ($l.Locked -eq $true) { $true } else { $false }
                }
            }
            $spacesData += @{
                name     = $s.Name
                zone     = @($s.Zone)
                monitor  = if ($null -ne $s.Monitor) { [int]$s.Monitor } else { 0 }
                shortcut = if ($s.Shortcut) { $s.Shortcut } else { "" }
                layers   = $layersData
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
                Monitor  = if ($null -ne $s.monitor) { [int]$s.monitor } else { 0 }
                Layers   = [System.Collections.ArrayList]::new()
                Shortcut = if ($s.shortcut) { $s.shortcut } else { "" }
            }
            $allWins = Get-AllWindowsForApply  # inclui minimizadas
            if ($s.layers) {
                # Novo formato: array de { title, process, locked }
                foreach ($ld in $s.layers) {
                    $proc   = if ($ld.process) { $ld.process } else { "" }
                    $locked = ($ld.locked -eq $true)
                    if ($locked) {
                        $match = $allWins | Where-Object { $_.Title -like "*$($ld.title)*" } | Select-Object -First 1
                    } else {
                        # Preferir match exato por titulo dentro do processo
                        if ($proc -and $ld.title) {
                            $match = $allWins | Where-Object { $_.Process -eq $proc -and $_.Title -like "*$($ld.title)*" } | Select-Object -First 1
                            if (-not $match) {
                                $match = $allWins | Where-Object { $_.Process -eq $proc } | Select-Object -First 1
                            }
                        } elseif ($proc) {
                            $match = $allWins | Where-Object { $_.Process -eq $proc } | Select-Object -First 1
                        } else { $match = $null }
                    }
                    if ($match) {
                        [void]$sp.Layers.Add(@{ Handle = $match.Handle; Title = $match.Title; Process = $match.Process; Locked = $locked })
                    }
                }
            } elseif ($s.windowTitles) {
                # Formato antigo: backward compat
                foreach ($wt in $s.windowTitles) {
                    if (-not $wt) { continue }
                    $match = $allWins | Where-Object { $_.Title -like "*$wt*" } | Select-Object -First 1
                    if ($match) {
                        [void]$sp.Layers.Add(@{ Handle = $match.Handle; Title = $match.Title; Process = $match.Process; Locked = $false })
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

$cAccent  = [System.Drawing.Color]::FromArgb(98, 224, 239)      # neon blue
$cSurface = [System.Drawing.Color]::FromArgb(10, 10, 28)
$cBg      = [System.Drawing.Color]::FromArgb(24, 5, 37)
$cBorder  = [System.Drawing.Color]::FromArgb(70, 170, 185)
$cText    = [System.Drawing.Color]::FromArgb(220, 220, 255)
$cMuted   = [System.Drawing.Color]::FromArgb(70, 170, 185)
$cGreen   = [System.Drawing.Color]::FromArgb(0, 255, 180)
$cOrange  = [System.Drawing.Color]::FromArgb(255, 140, 0)
$cRed     = [System.Drawing.Color]::FromArgb(255, 0, 120)
$cPink    = [System.Drawing.Color]::FromArgb(255, 0, 220)

$script:spaceColors = @(
    @{ Fill=[System.Drawing.Color]::FromArgb(35, 255, 0, 120);   Stroke=[System.Drawing.Color]::FromArgb(255, 0, 120)   },  # pink
    @{ Fill=[System.Drawing.Color]::FromArgb(35, 120, 0, 255);   Stroke=[System.Drawing.Color]::FromArgb(140, 0, 255)   },  # purple
    @{ Fill=[System.Drawing.Color]::FromArgb(35, 255, 230, 0);   Stroke=[System.Drawing.Color]::FromArgb(255, 230, 0)   },  # yellow
    @{ Fill=[System.Drawing.Color]::FromArgb(35, 255, 140, 0);   Stroke=[System.Drawing.Color]::FromArgb(255, 140, 0)   },  # orange
    @{ Fill=[System.Drawing.Color]::FromArgb(35, 255, 0, 220);   Stroke=[System.Drawing.Color]::FromArgb(255, 0, 220)   },  # magenta
    @{ Fill=[System.Drawing.Color]::FromArgb(35, 0, 255, 255);   Stroke=[System.Drawing.Color]::FromArgb(0, 255, 255)   }   # cyan
)

# ============================================================
#  FORMULARIO PRINCIPAL
# ============================================================

$form = New-Object System.Windows.Forms.Form
$form.Text            = "SnapLayout"
$form.Size            = New-Object System.Drawing.Size($_formCW, $_formCH)
$form.MinimumSize     = New-Object System.Drawing.Size($_formCW, $_formCH)
$form.StartPosition   = "CenterScreen"
$form.FormBorderStyle = "None"
$form.MaximizeBox     = $false
$form.BackColor       = $cBg
$form.ForeColor       = $cText
$form.Font            = New-Object System.Drawing.Font("Consolas", 8, [System.Drawing.FontStyle]::Bold)
$form.KeyPreview      = $true
# DoubleBuffered e ResizeRedraw sao propriedade/metodo protected — acesso via reflection
$setProp = [System.Windows.Forms.Control].GetProperty('DoubleBuffered', [System.Reflection.BindingFlags]'NonPublic,Instance')
$setProp.SetValue($form, $true, $null)
$setStyle = [System.Windows.Forms.Control].GetMethod('SetStyle', [System.Reflection.BindingFlags]'NonPublic,Instance')
$setStyle.Invoke($form, @([System.Windows.Forms.ControlStyles]::ResizeRedraw, $true))

# Borda externa cyan (2px) via Paint no form
$form.add_Paint({
    param($s, $e)
    $pen = New-Object System.Drawing.Pen($cAccent, 3)
    $e.Graphics.DrawRectangle($pen, 1, 1, ($form.Width - 3), ($form.Height - 3))
    $pen.Dispose()
})

# Drag via SnapAPI
$dragHandler = { param($s,$e)
    if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Left) {
        [SnapAPI]::ReleaseCapture()
        [SnapAPI]::SendMessage($form.Handle, 0xA1, 0x2, 0) | Out-Null
    }
}

# Barra de titulo cyan fina (so o titulo + botao fechar)
$pnlTitleBar = New-Object System.Windows.Forms.Panel
$pnlTitleBar.Location  = New-Object System.Drawing.Point(3, 3)
$pnlTitleBar.Size      = New-Object System.Drawing.Size(1074, 38)
$pnlTitleBar.BackColor = $cAccent
$pnlTitleBar.add_MouseDown($dragHandler)
$form.Controls.Add($pnlTitleBar)

# Botao minimizar
$btnMinimize = New-Object System.Windows.Forms.Button
$btnMinimize.Text      = "0"
$btnMinimize.Location  = New-Object System.Drawing.Point(1022, 5)
$btnMinimize.Size      = New-Object System.Drawing.Size(22, 22)
$btnMinimize.FlatStyle = "Flat"
$btnMinimize.BackColor = $cAccent
$btnMinimize.ForeColor = $cBg
$btnMinimize.FlatAppearance.BorderSize = 0
$btnMinimize.Font      = New-Object System.Drawing.Font("Webdings", 11, [System.Drawing.FontStyle]::Bold)
$btnMinimize.add_Click({ $form.WindowState = [System.Windows.Forms.FormWindowState]::Minimized })
$pnlTitleBar.Controls.Add($btnMinimize)

# Botao fechar
$btnClose = New-Object System.Windows.Forms.Button
$btnClose.Text      = "Ò"
$btnClose.Location  = New-Object System.Drawing.Point(1046, 7)
$btnClose.Size      = New-Object System.Drawing.Size(22, 22)
$btnClose.FlatStyle = "Flat"
$btnClose.BackColor = $cAccent
$btnClose.ForeColor = $cRed
$btnClose.FlatAppearance.BorderSize = 0
$btnClose.Font      = New-Object System.Drawing.Font("Wingdings 2", 10, [System.Drawing.FontStyle]::Bold)
$btnClose.add_Click({ $form.Close() })
$pnlTitleBar.Controls.Add($btnClose)

# Titulo centralizado na faixa cyan
$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Text      = "SNAPLAYOUTS"
$lblTitle.Font      = New-Object System.Drawing.Font("Consolas", 15, [System.Drawing.FontStyle]::Bold)
$lblTitle.ForeColor = $cBg
$lblTitle.Location  = New-Object System.Drawing.Point(0, 4)
$lblTitle.Size      = New-Object System.Drawing.Size(1074, 30)
$lblTitle.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$lblTitle.add_MouseDown($dragHandler)
$pnlTitleBar.Controls.Add($lblTitle)

# Faixa escura abaixo com o subtitulo (tambem arrastavel)
$pnlSubBar = New-Object System.Windows.Forms.Panel
$pnlSubBar.Location  = New-Object System.Drawing.Point(3, 41)
$pnlSubBar.Size      = New-Object System.Drawing.Size(1074, 22)
$pnlSubBar.BackColor = $cBg
$pnlSubBar.add_MouseDown($dragHandler)
$form.Controls.Add($pnlSubBar)

$lblSub = New-Object System.Windows.Forms.Label
$lblSub.Text      = "CRIADO POR NICKOLAS OLIVEIRA"
$lblSub.ForeColor = $cAccent
$lblSub.Font      = New-Object System.Drawing.Font("Consolas", 9, [System.Drawing.FontStyle]::Bold)
$lblSub.Location  = New-Object System.Drawing.Point(6, 3)
$lblSub.Size      = New-Object System.Drawing.Size(400, 16)
$lblSub.add_MouseDown($dragHandler)
$pnlSubBar.Controls.Add($lblSub)

# Linha separadora cyan abaixo do subtitulo
$sep = New-Object System.Windows.Forms.Panel
$sep.Location  = New-Object System.Drawing.Point(3, 63)
$sep.Size      = New-Object System.Drawing.Size(1074, 2)
$sep.BackColor = $cAccent
$form.Controls.Add($sep)

# Separadores verticais entre colunas
$sepV1 = New-Object System.Windows.Forms.Panel
$sepV1.Location  = New-Object System.Drawing.Point(210, 65)
$sepV1.Size      = New-Object System.Drawing.Size(1, 415)
$sepV1.BackColor = $cBorder
$form.Controls.Add($sepV1)

$sepV2 = New-Object System.Windows.Forms.Panel
$sepV2.Location  = New-Object System.Drawing.Point(782, 65)
$sepV2.Size      = New-Object System.Drawing.Size(1, 415)
$sepV2.BackColor = $cBorder
$form.Controls.Add($sepV2)

# ============================================================
#  COLUNA ESQUERDA
# ============================================================

# -- Coluna esquerda: w=210 (x=0..210)
$btnSnapshot = New-Object System.Windows.Forms.Button
$btnSnapshot.Text      = "SNAPSHOT DO DESKTOP"
$btnSnapshot.Location  = New-Object System.Drawing.Point(8, 78)
$btnSnapshot.Size      = New-Object System.Drawing.Size(194, 32)
$btnSnapshot.FlatStyle = "Flat"
$btnSnapshot.BackColor = $cBg
$btnSnapshot.ForeColor = $cAccent
$btnSnapshot.FlatAppearance.BorderColor = $cAccent
$btnSnapshot.FlatAppearance.BorderSize  = 1
$btnSnapshot.Font      = New-Object System.Drawing.Font("Consolas", 8, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($btnSnapshot)

$sepSnap = New-Object System.Windows.Forms.Panel
$sepSnap.Location  = New-Object System.Drawing.Point(8, 118)
$sepSnap.Size      = New-Object System.Drawing.Size(194, 1)
$sepSnap.BackColor = $cBorder
$form.Controls.Add($sepSnap)

$lblSaved = New-Object System.Windows.Forms.Label
$lblSaved.Text      = "LAYOUTS SALVOS"
$lblSaved.Font      = New-Object System.Drawing.Font("Consolas", 8, [System.Drawing.FontStyle]::Bold)
$lblSaved.ForeColor = $cAccent
$lblSaved.Location  = New-Object System.Drawing.Point(8, 126)
$lblSaved.Size      = New-Object System.Drawing.Size(194, 13)
$form.Controls.Add($lblSaved)

$lstSaved = New-Object System.Windows.Forms.ListBox
$lstSaved.Location    = New-Object System.Drawing.Point(8, 142)
$lstSaved.Size        = New-Object System.Drawing.Size(194, 290)
$lstSaved.BackColor   = $cBg
$lstSaved.ForeColor   = $cText
$lstSaved.BorderStyle = "None"
$lstSaved.Font        = New-Object System.Drawing.Font("Consolas", 8, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($lstSaved)

$btnSavedUp = New-Object System.Windows.Forms.Button
$btnSavedUp.Text      = "p"   # Wingdings 3: seta cima
$btnSavedUp.Font      = New-Object System.Drawing.Font("Wingdings 3", 7, [System.Drawing.FontStyle]::Bold)
$btnSavedUp.Location  = New-Object System.Drawing.Point(8, 436)
$btnSavedUp.Size      = New-Object System.Drawing.Size(28, 26)
$btnSavedUp.FlatStyle = "Flat"
$btnSavedUp.BackColor = $cBg
$btnSavedUp.ForeColor = $cAccent
$btnSavedUp.FlatAppearance.BorderColor = $cBorder
$btnSavedUp.FlatAppearance.BorderSize  = 1
$form.Controls.Add($btnSavedUp)

$btnSavedDown = New-Object System.Windows.Forms.Button
$btnSavedDown.Text      = "q"   # Wingdings 3: seta baixo
$btnSavedDown.Font      = New-Object System.Drawing.Font("Wingdings 3", 7, [System.Drawing.FontStyle]::Bold)
$btnSavedDown.Location  = New-Object System.Drawing.Point(40, 436)
$btnSavedDown.Size      = New-Object System.Drawing.Size(28, 26)
$btnSavedDown.FlatStyle = "Flat"
$btnSavedDown.BackColor = $cBg
$btnSavedDown.ForeColor = $cAccent
$btnSavedDown.FlatAppearance.BorderColor = $cBorder
$btnSavedDown.FlatAppearance.BorderSize  = 1
$form.Controls.Add($btnSavedDown)

$btnSavedRename = New-Object System.Windows.Forms.Button
$btnSavedRename.Text      = "RENOMEAR"
$btnSavedRename.Location  = New-Object System.Drawing.Point(72, 436)
$btnSavedRename.Size      = New-Object System.Drawing.Size(130, 26)
$btnSavedRename.FlatStyle = "Flat"
$btnSavedRename.BackColor = $cBg
$btnSavedRename.ForeColor = $cAccent
$btnSavedRename.FlatAppearance.BorderColor = $cBorder
$btnSavedRename.FlatAppearance.BorderSize  = 1
$btnSavedRename.Font      = New-Object System.Drawing.Font("Consolas", 8, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($btnSavedRename)

$btnDeleteSaved = New-Object System.Windows.Forms.Button
$btnDeleteSaved.Text      = "EXCLUIR"
$btnDeleteSaved.Location  = New-Object System.Drawing.Point(8, 483)
$btnDeleteSaved.Size      = New-Object System.Drawing.Size(95, 34)
$btnDeleteSaved.FlatStyle = "Flat"
$btnDeleteSaved.BackColor = $cBg
$btnDeleteSaved.ForeColor = $cRed
$btnDeleteSaved.FlatAppearance.BorderColor = $cRed
$btnDeleteSaved.FlatAppearance.BorderSize  = 1
$btnDeleteSaved.Font      = New-Object System.Drawing.Font("Consolas", 8, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($btnDeleteSaved)

$btnSetShortcut = New-Object System.Windows.Forms.Button
$btnSetShortcut.Text      = "ATALHO"
$btnSetShortcut.Location  = New-Object System.Drawing.Point(107, 483)
$btnSetShortcut.Size      = New-Object System.Drawing.Size(95, 34)
$btnSetShortcut.FlatStyle = "Flat"
$btnSetShortcut.BackColor = $cBg
$btnSetShortcut.ForeColor = $cAccent
$btnSetShortcut.FlatAppearance.BorderColor = $cAccent
$btnSetShortcut.FlatAppearance.BorderSize  = 1
$btnSetShortcut.Font      = New-Object System.Drawing.Font("Consolas", 8, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($btnSetShortcut)

# ============================================================
#  PRE-VISUALIZACAO - centro: x=212..782 (w=570)
# ============================================================

$lblPreview = New-Object System.Windows.Forms.Label
$lblPreview.Text      = "PREVIEW"
$lblPreview.Font      = New-Object System.Drawing.Font("Consolas", 8, [System.Drawing.FontStyle]::Bold)
$lblPreview.ForeColor = $cAccent
$lblPreview.Location  = New-Object System.Drawing.Point(220, 70)
$lblPreview.Size      = New-Object System.Drawing.Size(80, 14)
$form.Controls.Add($lblPreview)

# Botoes de pagina de monitor (so aparecem com 2+ monitores)
$btnMonPrev = New-Object System.Windows.Forms.Button
$btnMonPrev.Text      = "t"   # Wingdings 3: triangulo esquerdo
$btnMonPrev.Font      = New-Object System.Drawing.Font("Wingdings 3", 7, [System.Drawing.FontStyle]::Bold)
$btnMonPrev.Location  = New-Object System.Drawing.Point(302, 68)
$btnMonPrev.Size      = New-Object System.Drawing.Size(20, 16)
$btnMonPrev.FlatStyle = "Flat"
$btnMonPrev.BackColor = $cBg
$btnMonPrev.ForeColor = $cAccent
$btnMonPrev.FlatAppearance.BorderSize = 0
$btnMonPrev.Visible   = ($script:monitors.Count -gt 1)
$form.Controls.Add($btnMonPrev)

$lblMonIcon = New-Object System.Windows.Forms.Label
$lblMonIcon.Text      = "¿"   # Wingdings 2: icone de monitor
$lblMonIcon.Font      = New-Object System.Drawing.Font("Webdings", 9, [System.Drawing.FontStyle]::Bold)
$lblMonIcon.ForeColor = $cAccent
$lblMonIcon.Location  = New-Object System.Drawing.Point(323, 68)
$lblMonIcon.Size      = New-Object System.Drawing.Size(16, 14)
$lblMonIcon.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$lblMonIcon.Visible   = ($script:monitors.Count -gt 1)
$form.Controls.Add($lblMonIcon)

$lblMonPage = New-Object System.Windows.Forms.Label
$lblMonPage.Text      = "1"
$lblMonPage.Font      = New-Object System.Drawing.Font("Consolas", 8, [System.Drawing.FontStyle]::Bold)
$lblMonPage.ForeColor = $cAccent
$lblMonPage.Location  = New-Object System.Drawing.Point(339, 70)
$lblMonPage.Size      = New-Object System.Drawing.Size(25, 13)
$lblMonPage.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$lblMonPage.Visible   = ($script:monitors.Count -gt 1)
$form.Controls.Add($lblMonPage)

$btnMonNext = New-Object System.Windows.Forms.Button
$btnMonNext.Text      = "u"   # Wingdings 3: triangulo direito
$btnMonNext.Font      = New-Object System.Drawing.Font("Wingdings 3", 7, [System.Drawing.FontStyle]::Bold)
$btnMonNext.Location  = New-Object System.Drawing.Point(367, 68)
$btnMonNext.Size      = New-Object System.Drawing.Size(20, 16)
$btnMonNext.FlatStyle = "Flat"
$btnMonNext.BackColor = $cBg
$btnMonNext.ForeColor = $cAccent
$btnMonNext.FlatAppearance.BorderSize = 0
$btnMonNext.Visible   = ($script:monitors.Count -gt 1)
$form.Controls.Add($btnMonNext)

$lblRes = New-Object System.Windows.Forms.Label
$lblRes.Text      = "${SW} X ${SH}"
$lblRes.ForeColor = $cAccent
$lblRes.Font      = New-Object System.Drawing.Font("Consolas", 8, [System.Drawing.FontStyle]::Bold)
$lblRes.Location  = New-Object System.Drawing.Point(640, 70)
$lblRes.Size      = New-Object System.Drawing.Size(134, 13)
$lblRes.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$form.Controls.Add($lblRes)

# Funcao de atualizar label e canvas ao trocar de pagina
function Update-MonitorPage {
    $mon = $script:monitors[$script:previewMonitor]
    $lblMonPage.Text = "$($script:previewMonitor + 1)"
    $lblRes.Text     = "$($mon.Bounds.Width) X $($mon.Bounds.Height)"
    $pnlPreview.Invalidate()
}

$btnMonPrev.add_Click({
    if ($script:previewMonitor -gt 0) {
        $script:previewMonitor--
        Update-MonitorPage
    }
})
$btnMonNext.add_Click({
    if ($script:previewMonitor -lt ($script:monitors.Count - 1)) {
        $script:previewMonitor++
        Update-MonitorPage
    }
})

$pnlPreview = New-Object System.Windows.Forms.Panel
$pnlPreview.Location    = New-Object System.Drawing.Point(220, 84)
$pnlPreview.Size        = New-Object System.Drawing.Size($_previewW, $_previewH)
$pnlPreview.BackColor   = $cBg
$pnlPreview.BorderStyle = "None"
$form.Controls.Add($pnlPreview)

# ============================================================
#  SPACES / LAYERS - direita: x=784..1072 (w=288)
# ============================================================

$lblSpaces = New-Object System.Windows.Forms.Label
$lblSpaces.Text      = "SPACES LAYERS"
$lblSpaces.Font      = New-Object System.Drawing.Font("Consolas", 8, [System.Drawing.FontStyle]::Bold)
$lblSpaces.ForeColor = $cAccent
$lblSpaces.Location  = New-Object System.Drawing.Point(752, 70)
$lblSpaces.Size      = New-Object System.Drawing.Size(310, 14)
$form.Controls.Add($lblSpaces)

$pnlSpaces = New-Object System.Windows.Forms.Panel
$pnlSpaces.Location            = New-Object System.Drawing.Point(752, 84)
$pnlSpaces.Size                = New-Object System.Drawing.Size(310, 358)
$pnlSpaces.BackColor           = $cBg
$pnlSpaces.AutoScroll          = $true
$pnlSpaces.AutoScrollMinSize   = New-Object System.Drawing.Size(1, 1)
$form.Controls.Add($pnlSpaces)

$btnAddSpace = New-Object System.Windows.Forms.Button
$btnAddSpace.Text      = "+ ADD SPACE"
$btnAddSpace.Location  = New-Object System.Drawing.Point(752, 446)
$btnAddSpace.Size      = New-Object System.Drawing.Size(310, 28)
$btnAddSpace.FlatStyle = "Flat"
$btnAddSpace.BackColor = $cBg
$btnAddSpace.ForeColor = $cGreen
$btnAddSpace.FlatAppearance.BorderColor = $cGreen
$btnAddSpace.FlatAppearance.BorderSize  = 1
$btnAddSpace.Font      = New-Object System.Drawing.Font("Consolas", 8, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($btnAddSpace)

# ============================================================
#  BARRA DE ACOES (baixo)
# ============================================================

$sepBottom = New-Object System.Windows.Forms.Panel
$sepBottom.Location  = New-Object System.Drawing.Point(3, 480)
$sepBottom.Size      = New-Object System.Drawing.Size(1074, 1)
$sepBottom.BackColor = $cBorder
$form.Controls.Add($sepBottom)

$btnApply = New-Object System.Windows.Forms.Button
$btnApply.Text      = "APLICAR LAYOUT"
$btnApply.Location  = New-Object System.Drawing.Point(280, 483)
$btnApply.Size      = New-Object System.Drawing.Size(190, 34)
$btnApply.FlatStyle = "Flat"
$btnApply.BackColor = $cPink
$btnApply.ForeColor = [System.Drawing.Color]::White
$btnApply.FlatAppearance.BorderSize = 0
$btnApply.Font      = New-Object System.Drawing.Font("Consolas", 9, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($btnApply)

$btnSaveCurrent = New-Object System.Windows.Forms.Button
$btnSaveCurrent.Text      = "SALVAR NOVO LAYOUT"
$btnSaveCurrent.Location  = New-Object System.Drawing.Point(478, 483)
$btnSaveCurrent.Size      = New-Object System.Drawing.Size(190, 34)
$btnSaveCurrent.FlatStyle = "Flat"
$btnSaveCurrent.BackColor = $cBg
$btnSaveCurrent.ForeColor = $cAccent
$btnSaveCurrent.FlatAppearance.BorderColor = $cAccent
$btnSaveCurrent.FlatAppearance.BorderSize  = 1
$btnSaveCurrent.Font      = New-Object System.Drawing.Font("Consolas", 8, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($btnSaveCurrent)

# -- Botao salvar sobrescrevendo (disquete Wingdings "<")
$btnOverwrite = New-Object System.Windows.Forms.Button
$btnOverwrite.Text      = "<"
$btnOverwrite.Location  = New-Object System.Drawing.Point(674, 483)
$btnOverwrite.Size      = New-Object System.Drawing.Size(34, 34)
$btnOverwrite.FlatStyle = "Flat"
$btnOverwrite.BackColor = $cBg
$btnOverwrite.ForeColor = $cAccent
$btnOverwrite.FlatAppearance.BorderColor = $cAccent
$btnOverwrite.FlatAppearance.BorderSize  = 1
$btnOverwrite.Font      = New-Object System.Drawing.Font("Wingdings", 16)
$form.Controls.Add($btnOverwrite)

# -- Status
$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Text      = ""
$lblStatus.ForeColor = $cAccent
$lblStatus.Font      = New-Object System.Drawing.Font("Consolas", 8, [System.Drawing.FontStyle]::Bold)
$lblStatus.Location  = New-Object System.Drawing.Point(8, 520)
$lblStatus.Size      = New-Object System.Drawing.Size(1056, 14)
$form.Controls.Add($lblStatus)

# Grip de resize — usa WM_NCLBUTTONDOWN (mesmo mecanismo do drag da titlebar)
$grip = New-Object System.Windows.Forms.Panel
$grip.Size      = New-Object System.Drawing.Size(12, 12)
$grip.Location  = New-Object System.Drawing.Point(1065, 560)
$grip.BackColor = $cBg
$grip.Cursor    = [System.Windows.Forms.Cursors]::SizeNWSE
$form.Controls.Add($grip)
$grip.add_MouseDown({
    param($s, $e)
    if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Left) {
        [SnapAPI]::ReleaseCapture()
        [SnapAPI]::SendMessage($form.Handle, 0xA1, 0x11, 0) | Out-Null  # WM_NCLBUTTONDOWN, HTBOTTOMRIGHT
    }
})

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
    $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::None
    $g.Clear($cBg)

    if ($script:currentSpaces.Count -eq 0) {
        $f = New-Object System.Drawing.Font("Consolas", 9, [System.Drawing.FontStyle]::Bold)
        $br = New-Object System.Drawing.SolidBrush($cAccent)
        $g.DrawString("SELECIONE UM LAYOUT", $f, $br, 180, 170)
        return
    }

    $pw = $pnlPreview.Width  - 8
    $ph = $pnlPreview.Height - 8
    $i  = 0
    $globalIdx = 0

    foreach ($space in $script:currentSpaces) {
        if ($space.Monitor -ne $script:previewMonitor) { $globalIdx++; continue }
        $zp = $space.Zone
        $rx = 4 + [int]($pw * $zp[0] / 100)
        $ry = 4 + [int]($ph * $zp[1] / 100)
        $rw = [int]($pw * $zp[2] / 100) - 4
        $rh = [int]($ph * $zp[3] / 100) - 4

        $ci = $i % $script:spaceColors.Count
        $strokeCol = $script:spaceColors[$ci].Stroke
        $brush = New-Object System.Drawing.SolidBrush($script:spaceColors[$ci].Fill)
        $pen   = New-Object System.Drawing.Pen($strokeCol, 2)
        $g.FillRectangle($brush, $rx, $ry, $rw, $rh)
        $g.DrawRectangle($pen, $rx, $ry, $rw, $rh)

        # Highlight do space selecionado
        if ($globalIdx -eq $script:highlightedSpaceIndex) {
            $highlightPen = New-Object System.Drawing.Pen([System.Drawing.Color]::White, 3)
            $g.DrawRectangle($highlightPen, ($rx - 2), ($ry - 2), ($rw + 4), ($rh + 4))
        }

        $fontBold  = New-Object System.Drawing.Font("Consolas", 8, [System.Drawing.FontStyle]::Bold)
        $fontSmall = New-Object System.Drawing.Font("Consolas", 7, [System.Drawing.FontStyle]::Bold)
        $nameBrush = New-Object System.Drawing.SolidBrush($strokeCol)
        $textBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(200, 220, 220, 255))

        # StringFormat com word wrap e clip
        $sf = New-Object System.Drawing.StringFormat
        $sf.Trimming  = [System.Drawing.StringTrimming]::None
        $sf.FormatFlags = [System.Drawing.StringFormatFlags]::NoClip

        # Clip ao retangulo do space para nao vazar
        $clipRect = New-Object System.Drawing.RectangleF(($rx + 1), ($ry + 1), ($rw - 2), ($rh - 2))
        $g.SetClip($clipRect)

        # Nome do space com word wrap
        $nameRect = New-Object System.Drawing.RectangleF(($rx + 4), ($ry + 3), ($rw - 8), 30)
        $g.DrawString($space.Name.ToUpper(), $fontBold, $nameBrush, $nameRect, $sf)

        # Layers com word wrap
        $ly = [float]($ry + 20)
        $lineH = 13.0
        $layerIdx = 1
        foreach ($layer in $space.Layers) {
            if (($ly + $lineH) -gt ($ry + $rh - 14)) { break }
            $txt = "L${layerIdx}: $($layer.Title.ToUpper())"
            $layerRect = New-Object System.Drawing.RectangleF(($rx + 4), $ly, ($rw - 8), ($ry + $rh - 14 - $ly))
            $g.DrawString($txt, $fontSmall, $textBrush, $layerRect, $sf)
            # Medir quantas linhas o texto ocupa para avancar corretamente
            $measured = $g.MeasureString($txt, $fontSmall, ([int]($rw - 8)), $sf)
            $ly += [Math]::Max($lineH, $measured.Height)
            $layerIdx++
        }

        $g.ResetClip()

        # Atalho (canto inferior direito)
        if ($space.Shortcut) {
            $fontShortcut = New-Object System.Drawing.Font("Consolas", 7, [System.Drawing.FontStyle]::Bold)
            $scBrush = New-Object System.Drawing.SolidBrush($strokeCol)
            $scText  = $space.Shortcut
            $scSize  = $g.MeasureString($scText, $fontShortcut)
            $g.DrawString($scText, $fontShortcut, $scBrush, ($rx + $rw - $scSize.Width - 4), ($ry + $rh - $scSize.Height - 2))
        }

        $i++
        $globalIdx++
    }

    # Borda do preview
    $borderPen = New-Object System.Drawing.Pen($cBorder, 1)
    $g.DrawRectangle($borderPen, 0, 0, ($pnlPreview.Width - 1), ($pnlPreview.Height - 1))
})

# -- Double-click no preview para resetar highlight
$pnlPreview.add_MouseDoubleClick({
    param($s, $e)
    $script:highlightedSpaceIndex = -1
    Build-SpacePanel
    $pnlPreview.Invalidate()
})

$form.add_KeyDown({
    param($s, $e)
    if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Escape -and $script:highlightedSpaceIndex -ge 0) {
        $script:highlightedSpaceIndex = -1
        Build-SpacePanel
        $pnlPreview.Invalidate()
        $e.Handled = $true
    }
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
        if ($script:dragEdge -eq "move") {
            $dx = $curX - $script:dragStartMousePct[0]
            $dy = $curY - $script:dragStartMousePct[1]
            $space.Zone[0] = [Math]::Max(0, [Math]::Min(100 - $oz[2], $oz[0] + $dx))
            $space.Zone[1] = [Math]::Max(0, [Math]::Min(100 - $oz[3], $oz[1] + $dy))
        } else {
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
        }
        $pnlPreview.Invalidate()
        return
    }

    # Hover: com space destacado, checar só ele; sem highlight, checar todos
    $spacesToHover = if ($script:highlightedSpaceIndex -ge 0) {
        $sp = $script:currentSpaces[$script:highlightedSpaceIndex]
        if ($sp.Monitor -eq $script:previewMonitor) { @($sp) } else { @() }
    } else {
        @($script:currentSpaces | Where-Object { $_.Monitor -eq $script:previewMonitor })
    }

    $found = $false
    foreach ($space in $spacesToHover) {
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
        $inside = $e.X -gt ($rx + $hit) -and $e.X -lt ($ex - $hit) -and
                  $e.Y -gt ($ry + $hit) -and $e.Y -lt ($ey - $hit)
        $cursor = if     ($nr -and $nb) { [System.Windows.Forms.Cursors]::SizeNWSE }
                  elseif ($nl -and $nt) { [System.Windows.Forms.Cursors]::SizeNWSE }
                  elseif ($nr -or $nl)  { [System.Windows.Forms.Cursors]::SizeWE }
                  elseif ($nb -or $nt)  { [System.Windows.Forms.Cursors]::SizeNS }
                  elseif ($script:highlightedSpaceIndex -ge 0 -and $inside) { [System.Windows.Forms.Cursors]::SizeAll }
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
    $candidates = if ($script:highlightedSpaceIndex -ge 0) {
        $sp = $script:currentSpaces[$script:highlightedSpaceIndex]
        if ($sp.Monitor -eq $script:previewMonitor) {
            @([PSCustomObject]@{ Space = $sp; Idx = $script:highlightedSpaceIndex })
        } else { @() }
    } else {
        $k = 0
        $script:currentSpaces | ForEach-Object {
            if ($_.Monitor -eq $script:previewMonitor) {
                [PSCustomObject]@{ Space = $_; Idx = $k }
            }
            $k++
        }
    }
    foreach ($entry in $candidates) {
        $space = $entry.Space
        $idx   = $entry.Idx
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
        $inside = $e.X -gt ($rx + $hit) -and $e.X -lt ($ex - $hit) -and
                  $e.Y -gt ($ry + $hit) -and $e.Y -lt ($ey - $hit)
        $edge = ""
        if ($nr) { $edge += "right" }
        if ($nb) { $edge += "bottom" }
        if ($nl) { $edge += "left" }
        if ($nt) { $edge += "top" }
        if ($edge) {
            $script:dragSpaceIdx      = $idx
            $script:dragStartZone     = @($zp[0], $zp[1], $zp[2], $zp[3])
            $script:dragEdge          = $edge
            $pnlPreview.Capture       = $true
            break
        } elseif ($script:highlightedSpaceIndex -ge 0 -and $inside) {
            $script:dragSpaceIdx      = $idx
            $script:dragStartZone     = @($zp[0], $zp[1], $zp[2], $zp[3])
            $script:dragStartMousePct = @([Math]::Round($e.X / $pw * 100), [Math]::Round($e.Y / $ph * 100))
            $script:dragEdge          = "move"
            $pnlPreview.Capture       = $true
            break
        }
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
    $scrollY = [Math]::Abs($pnlSpaces.AutoScrollPosition.Y)
    $pnlSpaces.Controls.Clear()
    $pw = $pnlSpaces.Width - 20   # largura util: desconta margem + scrollbar
    $y = 6
    $i = 0
    $lastMonitor = -1
    $monColorIdx = @{}   # contador de cor por monitor, igual ao canvas

    foreach ($space in $script:currentSpaces) {
        if (-not $monColorIdx.ContainsKey($space.Monitor)) { $monColorIdx[$space.Monitor] = 0 }
        $ci = $monColorIdx[$space.Monitor] % $script:spaceColors.Count
        $monColorIdx[$space.Monitor]++

        if ($script:monitors.Count -gt 1 -and $space.Monitor -ne $lastMonitor) {
            $lastMonitor = $space.Monitor
            $lblMon = New-Object System.Windows.Forms.Label
            $lblMon.Text      = "── TELA $($space.Monitor + 1) ──"
            $lblMon.Location  = New-Object System.Drawing.Point(2, $y)
            $lblMon.Size      = New-Object System.Drawing.Size($pw, 16)
            $lblMon.Font      = New-Object System.Drawing.Font("Consolas", 7, [System.Drawing.FontStyle]::Bold)
            $lblMon.ForeColor = $cMuted
            $lblMon.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
            $pnlSpaces.Controls.Add($lblMon)
            $y += 18
        }

        $strokeColor = $script:spaceColors[$ci].Stroke

        # Indicador de cor + Nome + botoes
        $isHighlighted = ($i -eq $script:highlightedSpaceIndex)
        $pnlHeader = New-Object System.Windows.Forms.Panel
        $pnlHeader.Location  = New-Object System.Drawing.Point(2, $y)
        $pnlHeader.Size      = New-Object System.Drawing.Size($pw, 26)
        $pnlHeader.BackColor = if ($isHighlighted) { [System.Drawing.Color]::FromArgb(10, 60, 75) } else { $cBg }
        $pnlHeader.Cursor    = [System.Windows.Forms.Cursors]::Hand
        $pnlHeader.Tag       = $i
        # Click no header destaca o space no canvas
        $pnlHeader.add_Click({
            param($s, $e)
            if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Left) {
                $script:highlightedSpaceIndex = $s.Tag
                Build-SpacePanel
                $pnlPreview.Invalidate()
            }
        })
        $pnlSpaces.Controls.Add($pnlHeader)

        $colorBar = New-Object System.Windows.Forms.Panel
        $colorBar.Location  = New-Object System.Drawing.Point(0, 0)
        $colorBar.Size      = New-Object System.Drawing.Size(5, 26)
        $colorBar.BackColor = $strokeColor
        $pnlHeader.Controls.Add($colorBar)

        if ($isHighlighted) {
            $selBar = New-Object System.Windows.Forms.Panel
            $selBar.Location  = New-Object System.Drawing.Point(5, 0)
            $selBar.Size      = New-Object System.Drawing.Size(($pw - 5), 2)
            $selBar.BackColor = $cAccent
            $pnlHeader.Controls.Add($selBar)
        }

        # Label clicavel para renomear inline
        $lblName = New-Object System.Windows.Forms.Label
        $lblName.Text      = $space.Name.ToUpper()
        $lblName.ForeColor = if ($isHighlighted) { $cAccent } else { $cText }
        $lblName.Font      = New-Object System.Drawing.Font("Consolas", 10, [System.Drawing.FontStyle]::Bold)
        $lblName.Location  = New-Object System.Drawing.Point(10, 4)
        $lblName.Size      = New-Object System.Drawing.Size(($pw - 120), 18)
        $lblName.Cursor    = [System.Windows.Forms.Cursors]::IBeam

        # TextBox inline (oculto por padrao)
        $txtRename = New-Object System.Windows.Forms.TextBox
        $txtRename.Text        = $space.Name
        $txtRename.Font        = New-Object System.Drawing.Font("Consolas", 10, [System.Drawing.FontStyle]::Bold)
        $txtRename.Location    = New-Object System.Drawing.Point(9, 3)
        $txtRename.Size        = New-Object System.Drawing.Size(($pw - 120), 20)
        $txtRename.BackColor   = $cBg
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
        $btnUp.Text      = "p"
        $btnUp.Location  = New-Object System.Drawing.Point(($pw - 102), 2)
        $btnUp.Size      = New-Object System.Drawing.Size(18, 22)
        $btnUp.FlatStyle = "Flat"
        $btnUp.BackColor = $cBg
        $btnUp.ForeColor = $cAccent
        $btnUp.FlatAppearance.BorderSize = 0
        $btnUp.Font      = New-Object System.Drawing.Font("Wingdings 3", 8, [System.Drawing.FontStyle]::Bold)
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
        $btnDown.Text      = "q"
        $btnDown.Location  = New-Object System.Drawing.Point(($pw - 82), 2)
        $btnDown.Size      = New-Object System.Drawing.Size(18, 22)
        $btnDown.FlatStyle = "Flat"
        $btnDown.BackColor = $cBg
        $btnDown.ForeColor = $cAccent
        $btnDown.FlatAppearance.BorderSize = 0
        $btnDown.Font      = New-Object System.Drawing.Font("Wingdings 3", 8, [System.Drawing.FontStyle]::Bold)
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
        $btnDelSpace.Text      = "Ò"
        $btnDelSpace.Location  = New-Object System.Drawing.Point(($pw - 24), 2)
        $btnDelSpace.Size      = New-Object System.Drawing.Size(22, 22)
        $btnDelSpace.FlatStyle = "Flat"
        $btnDelSpace.BackColor = $cBg
        $btnDelSpace.ForeColor = $cRed
        $btnDelSpace.FlatAppearance.BorderSize = 0
        $btnDelSpace.Font      = New-Object System.Drawing.Font("Wingdings 2", 8, [System.Drawing.FontStyle]::Bold)
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
            $isLocked = ($layer.Locked -eq $true)

            # Botao lock/unlock (●=locked laranja, ○=process muted)
            $btnLock = New-Object System.Windows.Forms.Button
            $btnLock.Text      = if ($isLocked) { "Ï" } else { "Ð" }
            $btnLock.Location  = New-Object System.Drawing.Point(10, ($y - 6))
            $btnLock.Size      = New-Object System.Drawing.Size(20, 18)
            $btnLock.FlatStyle = "Flat"
            $btnLock.BackColor = $cBg
            $btnLock.ForeColor = if ($isLocked) { $cOrange } else { $cAccent }
            $btnLock.FlatAppearance.BorderSize = 0
            $btnLock.Font      = New-Object System.Drawing.Font("Webdings", 6)
            $btnLock.Tag       = @{ SpaceIdx = $i; LayerIdx = $layerIdx }
            $btnLock.add_Click({
                param($sender, $ea)
                $tag = $sender.Tag
                $sp  = $script:currentSpaces[$tag.SpaceIdx]
                if ($tag.LayerIdx -lt $sp.Layers.Count) {
                    $lyr = $sp.Layers[$tag.LayerIdx]
                    $lyr.Locked = -not ($lyr.Locked -eq $true)
                    # Se desbloqueando e o handle mudou, atualiza com janela corrente do processo
                    if (-not $lyr.Locked -and $lyr.Process) {
                        $cur = Get-VisibleWindows | Where-Object { $_.Process -eq $lyr.Process } | Select-Object -First 1
                        if ($cur) { $lyr.Handle = $cur.Handle; $lyr.Title = $cur.Title }
                    }
                }
                Build-SpacePanel
                $pnlPreview.Invalidate()
            })
            $pnlSpaces.Controls.Add($btnLock)

            # Texto: process para unlocked, titulo para locked
            $displayText = if ($isLocked -or -not $layer.Process) {
                $t = if ($layer.Title) { $layer.Title } else { "?" }
                if ($t.Length -gt 22) { $t.Substring(0, 19) + "..." } else { $t }
            } else {
                "[$($layer.Process)]"
            }

            $lblLayer = New-Object System.Windows.Forms.Label
            $lblLayer.Text      = " L$($layerIdx + 1): $displayText"
            $lblLayer.ForeColor = if ($isLocked) { $cOrange } else { $cText }
            $lblLayer.Font      = New-Object System.Drawing.Font("Consolas", 9, [System.Drawing.FontStyle]::Bold)
            $lblLayer.Location  = New-Object System.Drawing.Point(28, $y)
            $lblLayer.Size      = New-Object System.Drawing.Size(($pw - 62), 18)
            $pnlSpaces.Controls.Add($lblLayer)

            # Botao [X] remover layer
            $btnRemove = New-Object System.Windows.Forms.Button
            $btnRemove.Text      = "Ò"
            $btnRemove.Location  = New-Object System.Drawing.Point(($pw - 32), ($y - 1))
            $btnRemove.Size      = New-Object System.Drawing.Size(20, 18)
            $btnRemove.FlatStyle = "Flat"
            $btnRemove.BackColor = $cBg
            $btnRemove.ForeColor = $cRed
            $btnRemove.FlatAppearance.BorderSize = 0
            $btnRemove.Font      = New-Object System.Drawing.Font("Wingdings 2", 9, [System.Drawing.FontStyle]::Bold)
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
        $btnAdd.Text      = "+ ADD LAYER"
        $btnAdd.Location  = New-Object System.Drawing.Point(4, $y)
        $btnAdd.Size      = New-Object System.Drawing.Size(([int](($pw - 10) / 2)), 22)
        $btnAdd.FlatStyle = "Flat"
        $btnAdd.BackColor = $cBg
        $btnAdd.ForeColor = $cGreen
        $btnAdd.FlatAppearance.BorderColor = $cGreen
        $btnAdd.FlatAppearance.BorderSize  = 1
        $btnAdd.Font      = New-Object System.Drawing.Font("Consolas", 9, [System.Drawing.FontStyle]::Bold)
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
            $lstPopup.Font        = New-Object System.Drawing.Font("Consolas", 8, [System.Drawing.FontStyle]::Bold)
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
                [void]$space.Layers.Add(@{ Handle = $selWin.Handle; Title = $selWin.Title; Process = $selWin.Process; Locked = $false })
                Build-SpacePanel
                $pnlPreview.Invalidate()
            }
        })
        $pnlSpaces.Controls.Add($btnAdd)

        # Botao "Atalho..." por space
        $btnSpaceShortcut = New-Object System.Windows.Forms.Button
        $btnSpaceShortcut.Text      = if ($space.Shortcut) { $space.Shortcut.ToUpper() } else { "ATALHO" }
        $btnSpaceShortcut.Location  = New-Object System.Drawing.Point(([int](($pw - 10) / 2) + 6), $y)
        $btnSpaceShortcut.Size      = New-Object System.Drawing.Size(([int](($pw - 10) / 2)), 22)
        $btnSpaceShortcut.FlatStyle = "Flat"
        $btnSpaceShortcut.BackColor = $cBg
        $btnSpaceShortcut.ForeColor = $cAccent
        $btnSpaceShortcut.FlatAppearance.BorderColor = $cAccent
        $btnSpaceShortcut.FlatAppearance.BorderSize  = 1
        $btnSpaceShortcut.Font      = New-Object System.Drawing.Font("Consolas", 9, [System.Drawing.FontStyle]::Bold)
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
            $txtKey.Font      = New-Object System.Drawing.Font("Consolas", 11, [System.Drawing.FontStyle]::Bold)
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
    $pnlSpaces.AutoScrollPosition = New-Object System.Drawing.Point(0, $scrollY)
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
            Name     = $s.Name
            Zone     = @($s.Zone)
            Monitor  = if ($null -ne $s.Monitor) { [int]$s.Monitor } else { 0 }
            Layers   = [System.Collections.ArrayList]::new()
            Shortcut = if ($s.Shortcut) { $s.Shortcut } else { "" }
        }
        foreach ($l in $s.Layers) {
            [void]$sp.Layers.Add(@{ Handle = $l.Handle; Title = $l.Title; Process = $l.Process; Locked = $l.Locked })
        }
        $script:currentSpaces += $sp
    }

    Build-SpacePanel
    $pnlPreview.Invalidate()
    $lblStatus.Text      = "Layout: $name"
    $lblStatus.ForeColor = $cMuted
}

$btnAddSpace.add_Click({
    $targetMonitor = $script:previewMonitor
    if ($script:monitors.Count -gt 1) {
        $choices = @()
        for ($m = 0; $m -lt $script:monitors.Count; $m++) { $choices += "Tela $($m + 1)" }
        $dlg = New-Object System.Windows.Forms.Form
        $dlg.Text            = "Adicionar Space"
        $dlg.Size            = New-Object System.Drawing.Size(260, (130 + $choices.Count * 32))
        $dlg.StartPosition   = "CenterParent"
        $dlg.FormBorderStyle = "FixedToolWindow"
        $dlg.BackColor       = $cBg
        $dlg.ForeColor       = $cText
        $lbl = New-Object System.Windows.Forms.Label
        $lbl.Text     = "Em qual tela?"
        $lbl.Location = New-Object System.Drawing.Point(10, 12)
        $lbl.Size     = New-Object System.Drawing.Size(230, 18)
        $lbl.Font     = New-Object System.Drawing.Font("Consolas", 9, [System.Drawing.FontStyle]::Bold)
        $dlg.Controls.Add($lbl)
        $by = 36
        foreach ($label in $choices) {
            $btn = New-Object System.Windows.Forms.Button
            $btn.Text         = $label
            $btn.Location     = New-Object System.Drawing.Point(10, $by)
            $btn.Size         = New-Object System.Drawing.Size(225, 28)
            $btn.FlatStyle    = "Flat"
            $btn.BackColor    = $cSurface
            $btn.ForeColor    = $cAccent
            $btn.FlatAppearance.BorderColor = $cAccent
            $btn.FlatAppearance.BorderSize  = 1
            $btn.Font         = New-Object System.Drawing.Font("Consolas", 9, [System.Drawing.FontStyle]::Bold)
            $btn.Tag          = $choices.IndexOf($label)
            $btn.add_Click({
                param($s, $e)
                $script:_addSpaceMonitor = $s.Tag
                $dlg.DialogResult = [System.Windows.Forms.DialogResult]::OK
                $dlg.Close()
            })
            $dlg.Controls.Add($btn)
            $by += 32
        }
        if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }
        $targetMonitor = $script:_addSpaceMonitor
    }
    $n = $script:currentSpaces.Count + 1
    $newSpace = @{
        Name     = "Space $n"
        Zone     = @(0, 0, 100, 100)
        Monitor  = $targetMonitor
        Layers   = [System.Collections.ArrayList]::new()
        Shortcut = ""
    }
    $insertIdx = -1
    for ($k = 0; $k -lt $script:currentSpaces.Count; $k++) {
        if ($script:currentSpaces[$k].Monitor -eq $targetMonitor) { $insertIdx = $k }
    }
    if ($insertIdx -ge 0 -and $insertIdx -lt $script:currentSpaces.Count - 1) {
        $before = @($script:currentSpaces[0..$insertIdx])
        $after  = @($script:currentSpaces[($insertIdx + 1)..($script:currentSpaces.Count - 1)])
        $script:currentSpaces = $before + $newSpace + $after
    } else {
        $script:currentSpaces += $newSpace
    }
    Build-SpacePanel
    $pnlPreview.Invalidate()
    $pnlSpaces.ScrollControlIntoView($pnlSpaces.Controls[$pnlSpaces.Controls.Count - 1])
})

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
    $applied = Apply-Spaces $script:currentSpaces
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
                Name     = $s.Name
                Zone     = @($s.Zone)
                Monitor  = if ($null -ne $s.Monitor) { [int]$s.Monitor } else { 0 }
                Layers   = [System.Collections.ArrayList]::new()
                Shortcut = if ($s.Shortcut) { $s.Shortcut } else { "" }
            }
            foreach ($l in $s.Layers) {
                [void]$sp.Layers.Add(@{ Handle = $l.Handle; Title = $l.Title; Process = $l.Process; Locked = $l.Locked })
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

# -- Sobrescrever layout atual (disquete)
$btnOverwrite.add_Click({
    if ($script:currentSpaces.Count -eq 0) {
        $lblStatus.Text      = "Nenhum layout selecionado."
        $lblStatus.ForeColor = $cOrange
        return
    }
    if (-not $script:currentName) {
        $lblStatus.Text      = "Selecione um layout salvo primeiro."
        $lblStatus.ForeColor = $cOrange
        return
    }
    $spaces = @()
    foreach ($s in $script:currentSpaces) {
        $sp = @{
            Name     = $s.Name
            Zone     = @($s.Zone)
            Monitor  = if ($null -ne $s.Monitor) { [int]$s.Monitor } else { 0 }
            Layers   = [System.Collections.ArrayList]::new()
            Shortcut = if ($s.Shortcut) { $s.Shortcut } else { "" }
        }
        foreach ($l in $s.Layers) {
            [void]$sp.Layers.Add(@{ Handle = $l.Handle; Title = $l.Title; Process = $l.Process; Locked = $l.Locked })
        }
        $spaces += $sp
    }
    $existingShortcut = if ($script:savedLayouts.Contains($script:currentName)) { $script:savedLayouts[$script:currentName].Shortcut } else { "" }
    $script:savedLayouts[$script:currentName] = @{ Spaces = $spaces; Shortcut = $existingShortcut }
    Save-AllLayouts
    Refresh-SavedList
    $lblStatus.Text      = "Layout '$($script:currentName)' sobrescrito."
    $lblStatus.ForeColor = $cGreen
})

function Move-SavedLayout($name, $direction) {
    $keys = [System.Collections.ArrayList]@($script:savedLayouts.Keys)
    $idx  = $keys.IndexOf($name)
    if ($idx -lt 0) { return }
    if ($direction -eq "up"   -and $idx -eq 0)                   { return }
    if ($direction -eq "down" -and $idx -eq ($keys.Count - 1))   { return }
    $swap = if ($direction -eq "up") { $idx - 1 } else { $idx + 1 }
    $tmp = $keys[$idx]; $keys[$idx] = $keys[$swap]; $keys[$swap] = $tmp
    $newDict = [ordered]@{}
    foreach ($k in $keys) { $newDict[$k] = $script:savedLayouts[$k] }
    $script:savedLayouts = $newDict
}

function Get-SelectedLayoutName {
    $sel = $lstSaved.SelectedItem
    if (-not $sel) { return $null }
    return ($sel -replace '\s*\[.*\]\s*$', '')
}

$btnSavedUp.add_Click({
    $name = Get-SelectedLayoutName
    if (-not $name) { return }
    Move-SavedLayout $name "up"
    Save-AllLayouts
    Refresh-SavedList
    $match = $lstSaved.Items | Where-Object { ($_ -replace '\s*\[.*\]\s*$', '') -eq $name } | Select-Object -First 1
    if ($match) { $lstSaved.SelectedItem = $match }
})

$btnSavedDown.add_Click({
    $name = Get-SelectedLayoutName
    if (-not $name) { return }
    Move-SavedLayout $name "down"
    Save-AllLayouts
    Refresh-SavedList
    $match = $lstSaved.Items | Where-Object { ($_ -replace '\s*\[.*\]\s*$', '') -eq $name } | Select-Object -First 1
    if ($match) { $lstSaved.SelectedItem = $match }
})

$btnSavedRename.add_Click({
    $oldName = Get-SelectedLayoutName
    if (-not $oldName) {
        $lblStatus.Text      = "Selecione um layout salvo primeiro."
        $lblStatus.ForeColor = $cOrange
        return
    }
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text            = "Renomear Layout"
    $dlg.Size            = New-Object System.Drawing.Size(350, 150)
    $dlg.StartPosition   = "CenterParent"
    $dlg.FormBorderStyle = "FixedToolWindow"
    $dlg.BackColor       = $cBg
    $dlg.ForeColor       = $cText
    $lblN = New-Object System.Windows.Forms.Label
    $lblN.Text     = "Novo nome:"
    $lblN.Location = New-Object System.Drawing.Point(10, 15)
    $lblN.Size     = New-Object System.Drawing.Size(320, 18)
    $dlg.Controls.Add($lblN)
    $txtN = New-Object System.Windows.Forms.TextBox
    $txtN.Location  = New-Object System.Drawing.Point(10, 38)
    $txtN.Size      = New-Object System.Drawing.Size(315, 24)
    $txtN.BackColor = $cSurface
    $txtN.ForeColor = $cText
    $txtN.Text      = $oldName
    $txtN.SelectAll()
    $dlg.Controls.Add($txtN)
    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text         = "Renomear"
    $btnOk.Location     = New-Object System.Drawing.Point(10, 72)
    $btnOk.Size         = New-Object System.Drawing.Size(100, 30)
    $btnOk.FlatStyle    = "Flat"
    $btnOk.BackColor    = $cAccent
    $btnOk.ForeColor    = [System.Drawing.Color]::White
    $btnOk.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $dlg.Controls.Add($btnOk)
    $dlg.AcceptButton = $btnOk
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $newName = $txtN.Text.Trim()
        if ($newName -and $newName -ne $oldName -and -not $script:savedLayouts.Contains($newName)) {
            $newDict = [ordered]@{}
            foreach ($k in $script:savedLayouts.Keys) {
                $key = if ($k -eq $oldName) { $newName } else { $k }
                $newDict[$key] = $script:savedLayouts[$k]
            }
            $script:savedLayouts = $newDict
            if ($script:currentName -eq $oldName) { $script:currentName = $newName }
            Save-AllLayouts
            Refresh-SavedList
            $match = $lstSaved.Items | Where-Object { ($_ -replace '\s*\[.*\]\s*$', '') -eq $newName } | Select-Object -First 1
            if ($match) { $lstSaved.SelectedItem = $match }
            $lblStatus.Text      = "Layout renomeado para '$newName'."
            $lblStatus.ForeColor = $cGreen
        } elseif ($script:savedLayouts.Contains($newName) -and $newName -ne $oldName) {
            $lblStatus.Text      = "Ja existe um layout com esse nome."
            $lblStatus.ForeColor = $cOrange
        }
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
    $txtKey.Font      = New-Object System.Drawing.Font("Consolas", 11, [System.Drawing.FontStyle]::Bold)
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
        $rect = [SnapAPI]::GetRect($w.Handle)
        # Detectar em qual monitor a janela esta pelo centro da janela
        $cx = [int](($rect.Left + $rect.Right) / 2)
        $cy = [int](($rect.Top + $rect.Bottom) / 2)
        $monIdx = 0
        for ($mi = 0; $mi -lt $script:monitors.Count; $mi++) {
            $mb = $script:monitors[$mi].Bounds
            if ($cx -ge $mb.X -and $cx -lt ($mb.X + $mb.Width) -and
                $cy -ge $mb.Y -and $cy -lt ($mb.Y + $mb.Height)) {
                $monIdx = $mi
                break
            }
        }
        $mon = $script:monitors[$monIdx].WorkingArea
        # Converter para percentagem relativa ao monitor detectado
        $xp = [Math]::Max(0, [Math]::Min(100, [Math]::Round(($rect.Left - $mon.X) / $mon.Width  * 100)))
        $yp = [Math]::Max(0, [Math]::Min(100, [Math]::Round(($rect.Top  - $mon.Y) / $mon.Height * 100)))
        $wp = [Math]::Max(5, [Math]::Min(100, [Math]::Round(($rect.Right  - $rect.Left) / $mon.Width  * 100)))
        $hp = [Math]::Max(5, [Math]::Min(100, [Math]::Round(($rect.Bottom - $rect.Top)  / $mon.Height * 100)))
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
            Monitor  = $monIdx
            Layers   = [System.Collections.ArrayList]::new()
            Shortcut = ""
        }
        [void]$sp.Layers.Add(@{ Handle = $w.Handle; Title = $w.Title; Process = $w.Process; Locked = $false })
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
    $foreWin = [SnapAPI]::GetForegroundWindow()
    $title   = [SnapAPI]::GetTitle($foreWin)
    if (-not $title -or $title -like "*SnapLayout*") { return }
    # Verificar se esta janela esta locked em outro space
    foreach ($sp in $script:currentSpaces) {
        if ($sp -eq $space) { continue }
        $lockedThere = $sp.Layers | Where-Object { $_.Locked -eq $true -and $_.Handle -eq $foreWin }
        if ($lockedThere) {
            $short = if ($title.Length -gt 30) { $title.Substring(0, 27) + "..." } else { $title }
            $lblStatus.Text      = "'$short' esta locked em '$($sp.Name)'."
            $lblStatus.ForeColor = $cOrange
            return
        }
    }
    $proc = Get-WindowProcess $foreWin
    # Se ja existe layer com este handle, apenas reposiciona sem duplicar
    $existing = $space.Layers | Where-Object { $_.Handle -eq $foreWin } | Select-Object -First 1
    if (-not $existing) {
        [void]$space.Layers.Add(@{ Handle = $foreWin; Title = $title; Process = $proc; Locked = $false })
    }
    $z = ConvertZone $space.Zone $(if ($null -ne $space.Monitor) { $space.Monitor } else { 0 })
    [SnapAPI]::MoveWindow($foreWin, $z.X, $z.Y, $z.W, $z.H)
    [SnapAPI]::SetForegroundWindow($foreWin)
    Build-SpacePanel
    $pnlPreview.Invalidate()
    $short = if ($title.Length -gt 40) { $title.Substring(0, 37) + "..." } else { $title }
    $lblStatus.Text      = "'$short' encaixado em '$($space.Name)'."
    $lblStatus.ForeColor = $cGreen
}

function Apply-Spaces($spaces) {
    $applied    = 0
    $allWindows = Get-AllWindowsForApply  # inclui janelas minimizadas
    # Coletar handles locked para nao move-los em slots de processo de outros spaces
    $lockedHandles = [System.Collections.Generic.HashSet[IntPtr]]::new()
    foreach ($sp in $spaces) {
        foreach ($l in $sp.Layers) {
            if ($l.Locked -eq $true -and [SnapAPI]::IsWindow($l.Handle)) {
                [void]$lockedHandles.Add($l.Handle)
            }
        }
    }
    # Pre-reservar handles especificos de cada space para nao serem roubados como extras
    $reservedHandles = [System.Collections.Generic.HashSet[IntPtr]]::new()
    foreach ($sp in $spaces) {
        foreach ($l in $sp.Layers) {
            if ($l.Process -and -not ($l.Locked -eq $true) `
                    -and $l.Handle -ne [IntPtr]::Zero -and [SnapAPI]::IsWindow($l.Handle)) {
                [void]$reservedHandles.Add($l.Handle)
            }
        }
    }
    # Deduplicacao cross-space: handles ja alocados a um space anterior
    $usedHandles = [System.Collections.Generic.HashSet[IntPtr]]::new()
    foreach ($space in $spaces) {
        $z = ConvertZone $space.Zone $(if ($null -ne $space.Monitor) { $space.Monitor } else { 0 })
        $windows = [System.Collections.ArrayList]::new()
        foreach ($layer in $space.Layers) {
            if ($layer.Locked -eq $true) {
                if ([SnapAPI]::IsWindow($layer.Handle) -and -not $windows.Contains($layer.Handle)) {
                    [void]$windows.Add($layer.Handle)
                }
            } elseif ($layer.Process) {
                # 1. Adicionar handle especifico salvo (funciona mesmo minimizado)
                if ($layer.Handle -ne [IntPtr]::Zero -and [SnapAPI]::IsWindow($layer.Handle) `
                        -and -not $lockedHandles.Contains($layer.Handle) `
                        -and -not $usedHandles.Contains($layer.Handle) `
                        -and -not $windows.Contains($layer.Handle)) {
                    [void]$windows.Add($layer.Handle)
                }
                # 2. Empilhar extras: janelas do mesmo processo sem space reservado
                $extras = $allWindows | Where-Object {
                    $_.Process -eq $layer.Process `
                    -and -not $lockedHandles.Contains($_.Handle) `
                    -and -not $usedHandles.Contains($_.Handle) `
                    -and -not $reservedHandles.Contains($_.Handle) `
                    -and -not $windows.Contains($_.Handle)
                }
                foreach ($m in $extras) { [void]$windows.Add($m.Handle) }
            } elseif ([SnapAPI]::IsWindow($layer.Handle) `
                    -and -not $usedHandles.Contains($layer.Handle) `
                    -and -not $windows.Contains($layer.Handle)) {
                [void]$windows.Add($layer.Handle)
            }
        }
        foreach ($hwnd in $windows) { [void]$usedHandles.Add($hwnd) }
        $count = $windows.Count
        if ($count -le 0) { continue }
        # Todas as janelas ocupam o space inteiro (empilhamento, nao tiling)
        $layerW = $z.W; $layerH = $z.H
        foreach ($wi in 0..($count - 1)) {
            $hwnd = $windows[$wi]
            try { [void][SnapAPI]::MoveWindow($hwnd, [int]$z.X, [int]$z.Y, [int]$layerW, [int]$layerH); $applied++ } catch {}
        }
        [void][SnapAPI]::SetForegroundWindow($windows[$count - 1])
    }
    return $applied
}

function Apply-LayoutByName($name) {
    $layout = $script:savedLayouts[$name]
    if (-not $layout) { return }
    Apply-Spaces $layout.Spaces
}

$hotkeyTimer = New-Object System.Windows.Forms.Timer
$hotkeyTimer.Interval = 200
$hotkeyTimer.add_Tick({
    # --- Pruning: remover layers de janelas/processos encerrados ---
    $pruned = $false
    $runningProcs = $null  # lazy: so carregado para layers de processo
    foreach ($space in $script:currentSpaces) {
        $toRemove = @()
        foreach ($layer in $space.Layers) {
            if ($layer.Locked -eq $true) {
                # Locked: prune apenas se o HWND especifico foi destruido
                $gone = -not [SnapAPI]::IsWindow($layer.Handle)
            } elseif ($layer.Process) {
                # Baseado em processo: prune se o processo nao esta mais rodando
                if ($null -eq $runningProcs) {
                    $runningProcs = @([System.Diagnostics.Process]::GetProcesses() | ForEach-Object { $_.ProcessName })
                }
                $gone = $runningProcs -notcontains $layer.Process
            } else {
                # Legado (sem process): verificar HWND
                $gone = -not [SnapAPI]::IsWindow($layer.Handle)
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
            if ($binding.Mods -contains "Ctrl")  { if (-not ([SnapAPI]::GetAsyncKeyState(0x11) -band 0x8000)) { $modsOk = $false } }
            if ($binding.Mods -contains "Alt")   { if (-not ([SnapAPI]::GetAsyncKeyState(0x12) -band 0x8000)) { $modsOk = $false } }
            if ($binding.Mods -contains "Shift") { if (-not ([SnapAPI]::GetAsyncKeyState(0x10) -band 0x8000)) { $modsOk = $false } }
            if (-not $modsOk) { continue }
            $keyState = [SnapAPI]::GetAsyncKeyState($binding.VKey)
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
            if (-not ([SnapAPI]::GetAsyncKeyState(0x11) -band 0x8000)) { $modsOk = $false }
        }
        if ($binding.Mods -contains "Alt") {
            if (-not ([SnapAPI]::GetAsyncKeyState(0x12) -band 0x8000)) { $modsOk = $false }
        }
        if ($binding.Mods -contains "Shift") {
            if (-not ([SnapAPI]::GetAsyncKeyState(0x10) -band 0x8000)) { $modsOk = $false }
        }
        if (-not $modsOk) { continue }

        $keyState = [SnapAPI]::GetAsyncKeyState($binding.VKey)
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

function Sync-Layout {
    param([int]$cw, [int]$ch)
    if ($cw -lt 100 -or $ch -lt 100) { return }

    # Titlebar, subbar e separador horizontal do topo
    $pnlTitleBar.Width = $cw - 6
    $btnMinimize.Left  = $pnlTitleBar.Width - 52
    $btnClose.Left     = $pnlTitleBar.Width - 28
    $pnlSubBar.Width   = $cw - 6
    $sep.Width         = $cw - 6

    # Painel direito: mantém margem de 18px à direita
    $rightX = $cw - 328
    $pnlSpaces.Left   = $rightX
    $pnlSpaces.Height = $ch - 217
    $lblSpaces.Left   = $rightX
    $btnAddSpace.Left = $rightX
    $btnAddSpace.Top  = $ch - 129

    # Canvas fixo: centralizado entre sepV1 (x=210) e pnlSpaces
    $midLeft  = 211
    $midRight = $cw - 328
    $prevX = [Math]::Max($midLeft + 4, [int](($midLeft + $midRight - $script:previewW) / 2))
    $prevY = [Math]::Max(84, [int](84 + ((($ch - 95) - 84 - $script:previewH) / 2)))
    $pnlPreview.Left = $prevX
    $pnlPreview.Top  = $prevY
    $sepV2.Left   = $prevX + $script:previewW + 8
    $sepV2.Height = $ch - 160
    $sepV1.Height = $ch - 160
    $lblRes.Left  = $sepV2.Left - 142

    # Barra inferior
    $sepBottom.Top      = $ch - 95
    $sepBottom.Width    = $cw - 6
    $btnApply.Top       = $ch - 92
    $btnSaveCurrent.Top = $ch - 92
    $btnOverwrite.Top   = $ch - 92
    $btnDeleteSaved.Top = $ch - 92
    $btnSetShortcut.Top = $ch - 92
    $lblStatus.Top      = $ch - 55
    $lblStatus.Width    = $cw - 24

    # Painel esquerdo
    $lstSaved.Height    = $ch - 285
    $btnSavedUp.Top     = $ch - 139
    $btnSavedDown.Top   = $ch - 139
    $btnSavedRename.Top = $ch - 139

    # Grip de resize
    $grip.Left = $cw - 15
    $grip.Top  = $ch - 15

    $pnlPreview.Invalidate()
}

# Shown: dispara uma vez, ja com o tamanho final correto
$form.add_Shown({ Sync-Layout $form.ClientSize.Width $form.ClientSize.Height })

# Resize: reposiciona controles e forca repaint completo para evitar rastro
$form.add_Resize({
    Sync-Layout $form.ClientSize.Width $form.ClientSize.Height
    $form.Invalidate($true)
    $form.Update()
})

[void]$form.ShowDialog()
$hotkeyTimer.Stop()
$hotkeyTimer.Dispose()
