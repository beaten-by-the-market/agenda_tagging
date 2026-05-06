param(
    [string]$Folder = $PSScriptRoot,
    [switch]$Recurse,
    [switch]$Overwrite,
    [switch]$Visible,
    [switch]$NoAutoApprove
)

$ErrorActionPreference = "Stop"

$resolvedFolder = (Resolve-Path -LiteralPath $Folder).Path

Add-Type -TypeDefinition @"
using System;
using System.Text;
using System.Threading;
using System.Runtime.InteropServices;

public static class HwpAccessPromptApprover
{
    private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern bool EnumChildWindows(IntPtr hWndParent, EnumWindowsProc lpEnumFunc, IntPtr lParam);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

    [DllImport("user32.dll")]
    private static extern bool IsWindowVisible(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern IntPtr SendMessage(IntPtr hWnd, int Msg, IntPtr wParam, IntPtr lParam);

    private const int BM_CLICK = 0x00F5;
    private static Thread worker;
    private static volatile bool stop;

    public static void Start()
    {
        if (worker != null && worker.IsAlive) return;
        stop = false;
        worker = new Thread(Loop);
        worker.IsBackground = true;
        worker.Start();
    }

    public static void Stop()
    {
        stop = true;
        if (worker != null && worker.IsAlive) worker.Join(1000);
    }

    private static void Loop()
    {
        while (!stop)
        {
            try { ScanOnce(); } catch {}
            Thread.Sleep(250);
        }
    }

    private static void ScanOnce()
    {
        EnumWindows((hWnd, lParam) =>
        {
            if (!IsWindowVisible(hWnd)) return true;
            EnumChildWindows(hWnd, (child, childParam) =>
            {
                string cls = GetText(child, true);
                if (!cls.Equals("Button", StringComparison.OrdinalIgnoreCase)) return true;

                string text = Normalize(GetText(child, false));
                if (text == "\uBAA8\uB450\uC811\uADFC\uD5C8\uC6A9" ||
                    text == "\uBAA8\uB450\uD5C8\uC6A9" ||
                    text == "\uC804\uCCB4\uD5C8\uC6A9")
                {
                    SendMessage(child, BM_CLICK, IntPtr.Zero, IntPtr.Zero);
                }
                return true;
            }, IntPtr.Zero);
            return true;
        }, IntPtr.Zero);
    }

    private static string GetText(IntPtr hWnd, bool className)
    {
        StringBuilder sb = new StringBuilder(512);
        if (className) GetClassName(hWnd, sb, sb.Capacity);
        else GetWindowText(hWnd, sb, sb.Capacity);
        return sb.ToString();
    }

    private static string Normalize(string value)
    {
        if (value == null) return "";
        return value.Replace("&", "").Replace(" ", "").Replace("\t", "").Trim();
    }
}
"@

function Convert-HwpFile {
    param(
        [Parameter(Mandatory)]
        [__ComObject]$Hwp,

        [Parameter(Mandatory)]
        [string]$InputPath
    )

    $outputPath = [System.IO.Path]::ChangeExtension($InputPath, ".hwpx")
    if ((Test-Path -LiteralPath $outputPath) -and -not $Overwrite) {
        Write-Host "[SKIP] Exists: $outputPath"
        return
    }

    Write-Host "[OPEN] $InputPath"
    $opened = $Hwp.Open($InputPath, "HWP", "forceopen:true")
    if (-not $opened) {
        throw "Hwp.Open failed: $InputPath"
    }

    Write-Host "[SAVE] $outputPath"
    $saved = $Hwp.SaveAs($outputPath, "HWPX", "")
    if (-not $saved) {
        throw "Hwp.SaveAs HWPX failed: $outputPath"
    }
}

$files = Get-ChildItem -LiteralPath $resolvedFolder -Filter "*.hwp" -File -Recurse:$Recurse |
    Where-Object { $_.Extension -ieq ".hwp" }

if (-not $files) {
    Write-Host "No .hwp files found: $resolvedFolder"
    exit 0
}

if (-not $NoAutoApprove) {
    [HwpAccessPromptApprover]::Start()
}

$hwp = $null
try {
    $hwp = New-Object -ComObject HWPFrame.HwpObject

    try {
        $hwp.XHwpWindows.Item(0).Visible = [bool]$Visible
    } catch {}

    try {
        $registered = $hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
        Write-Host "[INFO] FilePathCheckerModule registered: $registered"
    } catch {
        Write-Host "[INFO] FilePathCheckerModule registration failed: $($_.Exception.Message)"
    }

    foreach ($file in $files) {
        Convert-HwpFile -Hwp $hwp -InputPath $file.FullName
    }

    Write-Host "[DONE] Converted: $($files.Count)"
}
finally {
    if ($hwp -ne $null) {
        try { $hwp.Quit() } catch {}
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($hwp) | Out-Null
    }

    if (-not $NoAutoApprove) {
        [HwpAccessPromptApprover]::Stop()
    }
}
