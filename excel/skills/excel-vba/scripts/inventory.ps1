# Read-only workbook identity inventory. Never starts Excel or changes its state.
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
public static class ExcelIdentityWindows {
    private delegate bool EnumProc(IntPtr hwnd, IntPtr data);
    [DllImport("user32.dll")] private static extern bool EnumWindows(EnumProc callback, IntPtr data);
    [DllImport("user32.dll")] private static extern bool EnumChildWindows(IntPtr parent, EnumProc callback, IntPtr data);
    [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr hwnd, out int pid);
    [DllImport("oleacc.dll")] private static extern int AccessibleObjectFromWindow(IntPtr hwnd, int id, ref Guid iid, [MarshalAs(UnmanagedType.Interface)] out object value);
    public static IntPtr[] Windows(int pid) {
        var result = new List<IntPtr>();
        EnumWindows((hwnd, data) => {
            int owner; GetWindowThreadProcessId(hwnd, out owner);
            if (owner == pid) {
                result.Add(hwnd);
                EnumChildWindows(hwnd, (child, unused) => { result.Add(child); return true; }, IntPtr.Zero);
            }
            return true;
        }, IntPtr.Zero);
        return result.ToArray();
    }
    public static object NativeObject(IntPtr hwnd) {
        var iid = new Guid("00020400-0000-0000-C000-000000000046");
        object value;
        return AccessibleObjectFromWindow(hwnd, unchecked((int)0xFFFFFFF0), ref iid, out value) == 0 ? value : null;
    }
    public static int Owner(long hwnd) {
        int pid; GetWindowThreadProcessId(new IntPtr(hwnd), out pid); return pid;
    }
}
'@

function Release-IdentityObject($value) {
    if ($null -ne $value -and [Runtime.InteropServices.Marshal]::IsComObject($value)) {
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($value)
    }
}

function Read-WorkbookIdentity($book) {
    return @{
        path = [string]$book.FullName
        directory = [string]$book.Path
        saved = [bool]$book.Saved
        format = [int]$book.FileFormat
    }
}

$before = @([Diagnostics.Process]::GetProcessesByName('EXCEL') | Sort-Object Id)
$identities = @($before | ForEach-Object { "$($_.Id):$($_.StartTime.ToUniversalTime().Ticks)" })
$rows = @()
foreach ($process in $before) {
    $app = $null
    $books = $null
    $protected = $null
    $addins = $null
    $row = @{ pid = $process.Id; complete = $false; workbooks = @() }
    try {
        foreach ($handle in [ExcelIdentityWindows]::Windows($process.Id)) {
            $native = $null
            $candidate = $null
            try {
                $native = [ExcelIdentityWindows]::NativeObject($handle)
                if ($null -eq $native) { continue }
                $candidate = $native.Application
                if ([ExcelIdentityWindows]::Owner([long]$candidate.Hwnd) -ne $process.Id) { continue }
                $app = $candidate
                $candidate = $null
                break
            } catch {
                # Non-Excel child windows do not expose the Excel native object model.
            } finally {
                Release-IdentityObject $candidate
                Release-IdentityObject $native
            }
        }
        if ($null -eq $app) { throw 'Excel native object model unavailable' }
        $protected = $app.ProtectedViewWindows
        if ([int]$protected.Count -gt 0) { throw 'Protected View workbooks cannot establish writable uniqueness' }
        $books = $app.Workbooks
        $count = [int]$books.Count
        for ($index = 1; $index -le $count; $index++) {
            $book = $null
            try {
                $book = $books.Item($index)
                $row.workbooks += Read-WorkbookIdentity $book
            } finally {
                Release-IdentityObject $book
            }
        }
        # Loaded VBA add-ins can be omitted from Workbooks enumeration.
        $addins = $app.AddIns
        for ($index = 1; $index -le [int]$addins.Count; $index++) {
            $addin = $null
            $book = $null
            try {
                $addin = $addins.Item($index)
                if (-not $addin.Installed -or [IO.Path]::GetExtension([string]$addin.FullName) -notin @('.xlam', '.xla')) { continue }
                $path = [string]$addin.FullName
                if (@($row.workbooks | Where-Object { $_.path -eq $path }).Count) { continue }
                $book = $books.Item([string]$addin.Name)
                if ([string]$book.FullName -ne $path) { throw 'Loaded add-in identity unavailable' }
                $row.workbooks += Read-WorkbookIdentity $book
            } finally {
                Release-IdentityObject $book
                Release-IdentityObject $addin
            }
        }
        if ([int]$books.Count -ne $count) { throw 'Workbook inventory changed' }
        $row.complete = $true
    } catch {
        $row.error = 'Excel workbook enumeration unavailable or changed; dismiss blocking dialogs and retry.'
    } finally {
        Release-IdentityObject $books
        Release-IdentityObject $protected
        Release-IdentityObject $addins
        Release-IdentityObject $app
        $process.Dispose()
    }
    $rows += $row
}
$after = @([Diagnostics.Process]::GetProcessesByName('EXCEL') | Sort-Object Id)
try {
    $afterIdentities = @($after | ForEach-Object { "$($_.Id):$($_.StartTime.ToUniversalTime().Ticks)" })
    $stable = ($identities -join ',') -eq ($afterIdentities -join ',')
    @{ complete = ($stable -and @($rows | Where-Object { -not $_.complete }).Count -eq 0); processes = $rows } | ConvertTo-Json -Depth 6 -Compress
} finally {
    $after | ForEach-Object { $_.Dispose() }
}
