param(
    [Parameter(Mandatory = $true)]
    [string]$PayloadPath
)

$ErrorActionPreference = 'Stop'
$excel = $null
$workbook = $null
$worksheet = $null
$table = $null
$backupWorkbook = $null
$backupWorksheet = $null
$currentStage = 'Starting'

# Excel constants / colors
$xlPatternNone = -4142
$yellowHighlight = 65535          # RGB(255,255,0)
$greenHighlight = 13561798        # RGB(198,239,206)

function Release-ComObject {
    param($Object)

    if ($null -ne $Object -and [System.Runtime.InteropServices.Marshal]::IsComObject($Object)) {
        try {
            [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($Object)
        }
        catch {}
    }
}

function Set-CellValue {
    param(
        [Parameter(Mandatory = $true)]$Cell,
        $Value,
        [switch]$IsDate
    )

    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) {
        $Cell.ClearContents()
        return
    }

    if ($IsDate) {
        $parsedDate = [DateTime]::Parse(
            [string]$Value,
            [System.Globalization.CultureInfo]::InvariantCulture,
            [System.Globalization.DateTimeStyles]::RoundtripKind
        )
        $Cell.Value2 = [double]$parsedDate.ToOADate()
        return
    }

    if ($Value -is [byte] -or
        $Value -is [int16] -or
        $Value -is [int32] -or
        $Value -is [int64] -or
        $Value -is [single] -or
        $Value -is [double] -or
        $Value -is [decimal]) {
        $Cell.Value2 = [double]$Value
        return
    }

    $Cell.Value2 = [string]$Value
}

function Clear-CommissionHighlights {
    param(
        [Parameter(Mandatory = $true)]$Worksheet
    )

    $usedRange = $null
    try {
        $usedRange = $Worksheet.UsedRange
        $usedRange.Interior.Pattern = $xlPatternNone
        $usedRange.Interior.ColorIndex = $xlPatternNone
    }
    finally {
        Release-ComObject $usedRange
    }
}

function Set-Highlight {
    param(
        [Parameter(Mandatory = $true)]$Range,
        [Parameter(Mandatory = $true)][int]$Color
    )

    $Range.Interior.Pattern = 1
    $Range.Interior.Color = $Color
}

try {
    $currentStage = 'Reading write payload'
    $payload = Get-Content -LiteralPath $PayloadPath -Raw | ConvertFrom-Json

    $currentStage = 'Starting Microsoft Excel'
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.ScreenUpdating = $false
    $excel.EnableEvents = $false

    # Clean the backup too, so it provides a highlight-free baseline for comparison.
    if ($null -ne $payload.backupPath -and
        -not [string]::IsNullOrWhiteSpace([string]$payload.backupPath) -and
        (Test-Path -LiteralPath ([string]$payload.backupPath))) {

        $currentStage = 'Clearing prior highlights from backup baseline'
        $backupWorkbook = $excel.Workbooks.Open([string]$payload.backupPath)
        $backupWorksheet = $backupWorkbook.Worksheets.Item('Commission')
        Clear-CommissionHighlights -Worksheet $backupWorksheet
        $backupWorkbook.Save()
        $backupWorkbook.Close($false)
        Release-ComObject $backupWorksheet
        Release-ComObject $backupWorkbook
        $backupWorksheet = $null
        $backupWorkbook = $null
    }

    $currentStage = 'Opening workbook'
    $workbook = $excel.Workbooks.Open([string]$payload.workbookPath)

    if ($workbook.ReadOnly) {
        throw 'Excel opened the workbook as read-only. Close the workbook in Excel and run the script again.'
    }

    $currentStage = 'Opening Commission table'
    $worksheet = $workbook.Worksheets.Item('Commission')
    $table = $worksheet.ListObjects.Item('Table1')

    $currentStage = 'Clearing prior highlights from Commission'
    Clear-CommissionHighlights -Worksheet $worksheet

    $currentStage = 'Updating existing invoice rows'
    foreach ($update in @($payload.existingRowUpdates)) {
        foreach ($change in @($update.changes)) {
            # Column 9 (Comm) is formula-controlled by Excel and must never be
            # written by this automation. Fail closed if a future payload ever
            # attempts to target it.
            if ([int]$change.column -eq 9) {
                throw 'Safety stop: the automation attempted to write to the Comm column (column 9).'
            }

            $cell = $null
            try {
                $cell = $worksheet.Cells.Item([int]$update.rowNumber, [int]$change.column)
                $isDate = [string]$change.field -eq 'Paid Date'
                Set-CellValue -Cell $cell -Value $change.newValue -IsDate:$isDate
                Set-Highlight -Range $cell -Color $yellowHighlight
            }
            finally {
                Release-ComObject $cell
            }
        }
    }

    $currentStage = 'Filling blank invoice rows'
    foreach ($fill in @($payload.blankRowFills)) {
        $highlightRange = $null
        try {
            for ($column = 6; $column -le 8; $column++) {
                $cell = $null
                try {
                    $cell = $worksheet.Cells.Item([int]$fill.rowNumber, $column)
                    switch ($column) {
                        6 { Set-CellValue -Cell $cell -Value $fill.invoice }
                        7 { Set-CellValue -Cell $cell -Value $fill.paidAmount }
                        8 { Set-CellValue -Cell $cell -Value $fill.paidDate -IsDate }
                    }
                }
                finally {
                    Release-ComObject $cell
                }
            }

            $highlightRange = $worksheet.Range(
                $worksheet.Cells.Item([int]$fill.rowNumber, 6),
                $worksheet.Cells.Item([int]$fill.rowNumber, 8)
            )
            Set-Highlight -Range $highlightRange -Color $greenHighlight
        }
        finally {
            Release-ComObject $highlightRange
        }
    }

    $currentStage = 'Adding new rows to Commission table'
    $groups = @($payload.insertGroups) | Sort-Object -Property afterRowNumber -Descending
    $headerSheetRow = [int]$table.HeaderRowRange.Row

    foreach ($group in $groups) {
        $basePosition = ([int]$group.afterRowNumber - $headerSheetRow) + 1
        $offset = 0

        foreach ($insert in @($group.inserts)) {
            $newListRow = $null
            $newRange = $null
            $previousCell = $null
            $highlightRange = $null

            try {
                $desiredPosition = $basePosition + $offset
                $currentRowCount = [int]$table.ListRows.Count

                if ($desiredPosition -gt $currentRowCount) {
                    $newListRow = $table.ListRows.Add()
                }
                else {
                    $newListRow = $table.ListRows.Add($desiredPosition)
                }

                $newRange = $newListRow.Range
                $previousSheetRow = [int]$newRange.Row - 1
                $previousCell = $worksheet.Cells.Item($previousSheetRow, 1)
                $percentValue = $previousCell.Value2

                Set-CellValue -Cell $newRange.Cells.Item(1, 1) -Value $percentValue
                Set-CellValue -Cell $newRange.Cells.Item(1, 2) -Value $insert.company
                Set-CellValue -Cell $newRange.Cells.Item(1, 3) -Value $insert.projectID
                Set-CellValue -Cell $newRange.Cells.Item(1, 4) -Value $insert.description
                $newRange.Cells.Item(1, 5).ClearContents()
                Set-CellValue -Cell $newRange.Cells.Item(1, 6) -Value $insert.invoice
                Set-CellValue -Cell $newRange.Cells.Item(1, 7) -Value $insert.paidAmount
                Set-CellValue -Cell $newRange.Cells.Item(1, 8) -Value $insert.paidDate -IsDate

                # Intentionally do not write to column 9 (Comm). It is an Excel
                # Table calculated column and Excel propagates its formula.
                $newRange.Cells.Item(1, 10).ClearContents()

                # Highlight the full newly inserted table row so it is unmistakable.
                $highlightRange = $newRange
                Set-Highlight -Range $highlightRange -Color $greenHighlight

                $offset += 1
            }
            finally {
                Release-ComObject $highlightRange
                Release-ComObject $previousCell
                Release-ComObject $newRange
                Release-ComObject $newListRow
            }
        }
    }

    $currentStage = 'Saving workbook'
    $workbook.Save()
    Write-Output 'Workbook updated successfully through Excel. Prior Commission highlights were cleared and current changes were highlighted.'
}
catch {
    $message = "Workbook update failed during stage '$currentStage': $($_.Exception.Message)"
    [Console]::Error.WriteLine($message)
    [Console]::Error.WriteLine($_.ScriptStackTrace)
    exit 1
}
finally {
    if ($null -ne $backupWorkbook) {
        try { $backupWorkbook.Close($false) } catch {}
    }

    if ($null -ne $workbook) {
        try { $workbook.Close($false) } catch {}
    }

    if ($null -ne $excel) {
        try { $excel.Quit() } catch {}
    }

    Release-ComObject $backupWorksheet
    Release-ComObject $backupWorkbook
    Release-ComObject $table
    Release-ComObject $worksheet
    Release-ComObject $workbook
    Release-ComObject $excel

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
