<#
PowerShell script: build_xlsm.ps1
Purpose: Build a macro-enabled Excel workbook (.xlsm) from CSV data and .bas modules in this repository.

Usage (example):
1. Open PowerShell as Administrator (recommended).
2. Navigate to the repository root where this script is located, e.g.:
   cd C:\path\to\school-inventory
3. Run (may need to set execution policy once):
   Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
   .\release\scripts\build_xlsm.ps1

Requirements:
- Windows with Microsoft Excel installed (desktop).
- In Excel: File -> Options -> Trust Center -> Trust Center Settings -> Macro Settings -> check "Trust access to the VBA project object model".
- PowerShell allowed to run scripts (use Bypass or RemoteSigned for this session).

What the script does:
- Creates a new Excel workbook
- Imports every CSV from release/data/ into a worksheet (sheet name = filename without extension)
- Converts each range into an Excel Table and names it (UsersTbl, ItemsTbl, TxTbl, etc.)
- Imports every .bas module from release/vba/ into the workbook VBA project
- Adds a simple fallback login/show routine (AutoShowLogin) so you can test login even if the UserForm hasn't been created manually
- Saves the result as an .xlsm file under release/files/School_Inventory.xlsm

Notes:
- If the repository path or folders differ, run the script from the repo root or pass a custom --root parameter.
- If Excel blocks programmatic VBProject access, enable the Trust setting (see requirements above).
#>
param(
    [string]$Root = (Get-Location).Path,
    [string]$OutFile = "$PSScriptRoot\..\..\files\School_Inventory.xlsm"
)

function Write-Log { param($m) Write-Host "[build_xlsm] $m" }

# Normalize paths
$Root = (Resolve-Path $Root).Path
$DataDir = Join-Path $Root 'release\data'
$VbaDir = Join-Path $Root 'release\vba'
$OutDir = Split-Path (Resolve-Path $OutFile -ErrorAction SilentlyContinue) -Parent
if (-not $OutDir) { $OutDir = Join-Path $PSScriptRoot '\..\..\files' }
if (-not (Test-Path $OutDir)) { New-Item -ItemType Directory -Path $OutDir | Out-Null }
$OutFile = Join-Path $OutDir (Split-Path $OutFile -Leaf)

# Table name mapping
$tableMap = @{
    'Users' = 'UsersTbl'
    'Categories' = 'CategoriesTbl'
    'Warehouses' = 'WarehousesTbl'
    'Suppliers' = 'SuppliersTbl'
    'Items' = 'ItemsTbl'
    'PurchaseOrders' = 'POsTbl'
    'PO_Lines' = 'POLinesTbl'
    'StockTransactions' = 'TxTbl'
    'Transfers' = 'TransfersTbl'
    'AuditLog' = 'AuditTbl'
}

# Check requirements
Write-Log "Checking prerequisites..."
try {
    $excel = New-Object -ComObject Excel.Application -ErrorAction Stop
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    Write-Log "Excel COM available."
} catch {
    Write-Log "ERROR: Excel COM object cannot be created. Ensure Excel is installed and you are on Windows.";
    throw
}

# Warn about Trust setting
Write-Log "Make sure Excel option 'Trust access to the VBA project object model' is enabled (Excel -> Options -> Trust Center -> Trust Center Settings -> Macro Settings)."

# Create workbook and import CSVs
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$wb = $excel.Workbooks.Add()
# remove default extra sheets
while ($wb.Worksheets.Count -gt 0) { $wb.Worksheets.Item(1).Delete(); if ($wb.Worksheets.Count -eq 0) { break } }

# For each CSV in data directory
if (-not (Test-Path $DataDir)) { Write-Log "No data directory found at $DataDir"; $wb.Close($false); $excel.Quit(); exit 1 }
$csvFiles = Get-ChildItem -Path $DataDir -Filter *.csv | Sort-Object Name
foreach ($csv in $csvFiles) {
    $sheetName = [System.IO.Path]::GetFileNameWithoutExtension($csv.Name)
    Write-Log "Importing CSV: $($csv.Name) -> Sheet '$sheetName'"
    $ws = $wb.Worksheets.Add()
    $ws.Name = $sheetName
    # Import CSV using Import-Csv for correct handling of commas
    $rows = Import-Csv -Path $csv.FullName -Delimiter ',' -Encoding UTF8
    if ($rows.Count -eq 0) { Write-Log "  (empty CSV)"; continue }
    # Write headers
    $headers = $rows[0].PSObject.Properties.Name
    for ($c=0; $c -lt $headers.Count; $c++) {
        $ws.Cells.Item(1, $c+1).Value2 = $headers[$c]
    }
    $r = 2
    foreach ($row in $rows) {
        for ($c=0; $c -lt $headers.Count; $c++) {
            $val = $row.$($headers[$c])
            $ws.Cells.Item($r, $c+1).Value2 = $val
        }
        $r++
    }
    # Convert to Table
    $lastRow = $ws.UsedRange.Rows.Count
    $lastCol = $ws.UsedRange.Columns.Count
    $rangeAddress = $ws.Range($ws.Cells.Item(1,1), $ws.Cells.Item($lastRow, $lastCol))
    try {
        $lst = $ws.ListObjects.Add(1, $rangeAddress, $null, 1)
        $tblName = $tableMap[$sheetName]
        if ($tblName) { try { $lst.Name = $tblName } catch { Write-Log "  Could not rename table to $tblName. Ensure name is valid." } }
    } catch {
        Write-Log "  Warning: could not create table on sheet $sheetName: $_"
    }
}

# Import VBA modules
if (Test-Path $VbaDir) {
    Write-Log "Importing VBA modules from $VbaDir"
    foreach ($bas in Get-ChildItem -Path $VbaDir -Filter *.bas) {
        Write-Log "  Importing $($bas.Name)"
        try {
            $wb.VBProject.VBComponents.Import($bas.FullName) | Out-Null
        } catch {
            Write-Log "  ERROR importing $($bas.Name): $_"
            Write-Log "  If this fails, ensure Excel Trust settings allow programmatic access to VBProject and run PowerShell as Administrator."
        }
    }
    # Also import UserForm code (we've stored it as a .txt). We'll add it as a standard module fallback if actual .frm import not possible.
    foreach ($uf in Get-ChildItem -Path $VbaDir -Filter *.txt) {
        $code = Get-Content -Path $uf.FullName -Raw
        # Create a standard module to hold the text of the form (as guidance)
        try {
            $mod = $wb.VBProject.VBComponents.Add(1) # vbext_ct_StdModule
            $mod.Name = ([System.IO.Path]::GetFileNameWithoutExtension($uf.Name)) + "_TXT"
            $mod.CodeModule.AddFromString($code)
            Write-Log "  Added fallback module for $($uf.Name)"
        } catch {
            Write-Log "  Could not add fallback module for $($uf.Name): $_"
        }
    }
} else {
    Write-Log "No VBA directory found at $VbaDir"
}

# Add fallback AutoShowLogin module (uses UserForm if available, otherwise InputBox)
$autoLoginCode = @'
Option Explicit

Sub AutoShowLogin()
    On Error Resume Next
    Dim uf As Object
    Err.Clear
    Set uf = VBA.UserForms.Add("frmLogin")
    If Err.Number = 0 Then
        Err.Clear
        uf.Show
        Exit Sub
    End If
    Err.Clear
    ' Fallback: InputBoxes
    Dim u As String, p As String
    u = InputBox("اسم المستخدم:", "تسجيل الدخول")
    If u = "" Then Exit Sub
    p = InputBox("كلمة المرور:", "تسجيل الدخول")
    If p = "" Then Exit Sub
    If ValidateLogin(u, p) Then
        AppendAuditLog(CurrentUserID, "Login", "User logged in: " & u)
        MsgBox "تم تسجيل الدخول باسم: " & u, vbInformation
    Else
        MsgBox "اسم المستخدم أو كلمة المرور غير صحيحة.", vbExclamation
    End If
End Sub
'@
try {
    $m = $wb.VBProject.VBComponents.Add(1) # vbext_ct_StdModule
    $m.Name = "AutoLoginUI"
    $m.CodeModule.AddFromString($autoLoginCode)
    Write-Log "Added AutoLoginUI module."
} catch {
    Write-Log "Could not add AutoLoginUI module: $_"
}

# Save workbook
try {
    Write-Log "Saving workbook to $OutFile"
    $wb.SaveAs($OutFile, 52) # 52 = xlOpenXMLWorkbookMacroEnabled (.xlsm)
    Write-Log "Saved successfully."
} catch {
    Write-Log "ERROR saving workbook: $_"
} finally {
    $wb.Close($false)
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}

Write-Log "Done. Open the file in Excel and run the AutoShowLogin sub (or create the frmLogin form manually)."
Write-Log "Output file: $OutFile"
