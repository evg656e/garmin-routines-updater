param(
    [string]$BaseDir = (Join-Path $HOME 'Dropbox\sport\swim'),
    [int]$MinDistance = 250,
    [int]$MaxDistance = [int]::MaxValue,
    [string[]]$SwimStrokes = @("BREASTSTROKE", "FREESTYLE"),
    [string]$MSOfficePath = 'C:\Program Files (x86)\Microsoft Visual Studio\Shared\Visual Studio Tools for Office\PIA\Office15',
    [switch]$Silent
)

function log([string]$Message) {
    if (!$Silent) {
        Write-Host $Message
    }
}

function warn([string]$Message) {
    if (!$Silent) {
        Write-Host $Message -ForegroundColor Yellow
    }
}

function done {
    log 'Done.'
}

$IndexFile = Join-Path $BaseDir 'Data\lap_swimming\index.json'
$UpdatedFile = Join-Path $BaseDir 'Data\lap_swimming\updated.json'
$SplitsDir = Join-Path $BaseDir 'Data\lap_swimming\splits'
$ActivityDir = Join-Path $BaseDir 'Data\lap_swimming\activity'
$RoutinesFile = Join-Path $BaseDir 'Routines.xlsx'

Add-Type -Path (Join-Path $MSOfficePath 'Microsoft.Office.Interop.Excel.dll')

function ConvertTo-SentenceCase([string]$str) {
    if ($str.Length -eq 0) {
        return $str
    }
    return $str.Substring(0, 1).ToUpper() + $str.Substring(1).ToLower()
}

function Get-Json([string]$Path, [switch]$AsHashtable) {
    return Get-Content $Path -Raw -Encoding utf8 | ConvertFrom-Json -AsHashtable:$AsHashtable
}

function Get-JsonDefault([string]$Path, [object]$DefaultValue, [switch]$AsHashtable) {
    if (!(Test-Path $Path)) {
        return $DefaultValue
    }
    return Get-Content $Path -Raw -Encoding utf8 | ConvertFrom-Json -AsHashtable:$AsHashtable
}

function Set-Json([string]$Path, [object]$Value) {
    $ParentPath = Split-Path $Path -Parent
    if (!(Test-Path $ParentPath)) {
        mkdir $ParentPath -Force
    }
    Set-Content -Path $Path -Value (ConvertTo-Json -InputObject $Value) -Encoding utf8
}

function Find-ContinousIntervals($Laps) {
    $Results = @()
    $Count = $Laps.Count
    if ($Count -eq 0) {
        return $Results
    }
    $First = 0
    $PrevLap = $Laps[0]
    $TotalDistance = $PrevLap.distance ?? 0
    for ($Index = 1; $Index -lt $Count; $Index++) {
        $CurrLap = $Laps[$Index]
        if ($CurrLap.swimStroke -ne $PrevLap.swimStroke) {
            if (($TotalDistance -ge $MinDistance) -and
                ($TotalDistance -le $MaxDistance) -and
                ($PrevLap.swimStroke -in $SwimStrokes)) {
                $Results += , @($First, $Index)
            }
            $TotalDistance = 0
            $First = $Index
        }
        $TotalDistance += $CurrLap.distance ?? 0
        $PrevLap = $CurrLap
    }
    if (($TotalDistance -ge $MinDistance) -and
        ($TotalDistance -le $MaxDistance) -and
        ($PrevLap.swimStroke -in $SwimStrokes)) {
        $Results += , @($First, $Count)
    }
    return , $Results
}

$DateRow = 1
$LinkRow = 2
$DescriptionRow = 3
$TimeRow = 4
$StrokeRateRow = 5
$LapNumberRow = 6
$DistanceRow = 7
$DurationRow = 8
$StrokesNumberRow = 9
$SpeedRow = 10
$PaceRow = 11
$CadenceRow = 12
$StrokeLengthRow = 13
$SwolfRow = 14
$LastUsedRow = 14

$TitleColumn = 1
$TotalsColumn = 2
$FirstDataColumn = 3

function Update-Routines($Workbook, $ActivityId, $Activity, $Laps, $Results) {
    $ActivityLink = "https://connect.garmin.com/modern/activity/$ActivityId"

    $SummaryDTO = $Activity.summaryDTO
    $StartTimeLocal = $SummaryDTO.startTimeLocal

    log "Updating routines for $ActivityLink ..."
    log "Date: $StartTimeLocal"

    $Description = $Activity.description
    # $StartDate = [DateTime]::Parse($StartTimeLocal, $null, [System.Globalization.DateTimeStyles]::RoundtripKind).ToShortDateString()
    $StartDate = $StartTimeLocal.ToShortDateString()

    $ResultsIndex = 0
    $ResultsCount = $Results.Count
    foreach ($Result in $Results) {
        $ResultsIndex++
        log "Updating interval $ResultsIndex of $ResultsCount..."
        $First = $Result[0]
        $Last = $Result[1]
        $LapCount = $Last - $First
        $FirstLap = $Laps[$First]
        $SwimStroke = ConvertTo-SentenceCase $FirstLap.swimStroke
        log "Stroke: $SwimStroke"
        log "Laps: $LapCount"

        try {
            $RoutinesSheet = $Workbook.Worksheets.Item($SwimStroke + ' Routines')
            $ResultsSheet = $Workbook.Worksheets.Item($SwimStroke + ' Results')
        }
        catch {
            warn "$SwimStroke sheet does not exists!"
            return
        }

        log 'Copying template...'
        $TemplateSheet = $Workbook.Worksheets.Item('Template')
        [void]$TemplateSheet.Range($TemplateSheet.Rows.Item(1), $TemplateSheet.Rows.Item($LastUsedRow + 1)).Copy()
        [void]$RoutinesSheet.Rows.Item(1).Insert([Microsoft.Office.Interop.Excel.XlInsertShiftDirection]::xlShiftDown)
        [void]$ResultsSheet.Rows.Item(2).Insert([Microsoft.Office.Interop.Excel.XlInsertShiftDirection]::xlShiftDown, [Microsoft.Office.Interop.Excel.XlInsertFormatOrigin]::xlFormatFromRightOrBelow)

        log 'Filling summary...'
        $ResultsSheet.Cells.Item(2, $DateRow).Value2 = $RoutinesSheet.Cells.Item($DateRow, $TotalsColumn).Value2 = $StartDate

        [void]$RoutinesSheet.Hyperlinks.Add($RoutinesSheet.Cells.Item($LinkRow, $TotalsColumn), $ActivityLink)
        [void]$ResultsSheet.Hyperlinks.Add($ResultsSheet.Cells.Item(2, $LinkRow), $ActivityLink)

        if (!$Description) {
            $ResultsSheet.Cells.Item(2, $DescriptionRow).Value2 = $RoutinesSheet.Cells.Item($DescriptionRow, $TotalsColumn).Value2 = $Description
        }

        log 'Filling data...'
        for ($Index = $First; $Index -lt $Last; $Index++) {
            $LapNumber = $Index - $First + 1
            $DataColumnIndex = $FirstDataColumn + $LapNumber - 1
            $Lap = $Laps[$Index]
            if ($LapNumber -gt 1) {
                [void]$RoutinesSheet.Range($RoutinesSheet.Cells.Item($LapNumberRow, $DataColumnIndex), $RoutinesSheet.Cells.Item($LastUsedRow, $DataColumnIndex)).FillRight()
            }
            $RoutinesSheet.Cells.Item($LapNumberRow, $DataColumnIndex).Value2 = $LapNumber
            $RoutinesSheet.Cells.Item($DistanceRow, $DataColumnIndex).Value2 = $Lap.distance
            $RoutinesSheet.Cells.Item($DurationRow, $DataColumnIndex).Value2 = $Lap.duration
            $RoutinesSheet.Cells.Item($StrokesNumberRow, $DataColumnIndex).Value2 = $Lap.totalNumberOfStrokes
        }


        log 'Updating totals...'
        if ($LapCount -gt 1) {
            $LastDataColumn = $FirstDataColumn + $LapCount - 1
            Update-FormulaRange $RoutinesSheet $LapNumberRow $TotalsColumn $LapNumberRow $FirstDataColumn $LapNumberRow $LastDataColumn
            Update-FormulaRange $RoutinesSheet $DistanceRow $TotalsColumn $DistanceRow $FirstDataColumn $DistanceRow $LastDataColumn
            Update-FormulaRange $RoutinesSheet $DurationRow $TotalsColumn $DurationRow $FirstDataColumn $DurationRow $LastDataColumn
            Update-FormulaRange $RoutinesSheet $StrokesNumberRow $TotalsColumn $StrokesNumberRow $FirstDataColumn $StrokesNumberRow $LastDataColumn
        }

        log 'Filling results...'
        for ($Row = $TimeRow; $Row -le $SwolfRow; $Row++) {
            $ResultsSheet.Cells.Item(2, $Row).Value2 = $RoutinesSheet.Cells.Item($Row, $TotalsColumn).Value2
        }
    }

    done
}

$FormulaRegExp = [regex]::new('\(.+?\)')

function Update-FormulaRange($Sheet, $FormulaRow, $FormulaColumn, $BeginRangeRow, $BeginRangeColumn, $EndRangeRow, $EndRangeColumn) {
    $Sheet.Cells.Item($FormulaRow, $FormulaColumn).Formula = $FormulaRegExp.Replace(
        $Sheet.Cells.Item($FormulaRow, $FormulaColumn).Formula,
        '(' + $Sheet.Range($Sheet.Cells.Item($BeginRangeRow, $BeginRangeColumn), $Sheet.Cells.Item($EndRangeRow, $EndRangeColumn)).Address($false, $false) + ')'
    )
}

function Update-Activity($Workbook, $ActivityId) {
    log "Updating activity $ActivityId..."

    $SplitsFile = Join-Path $SplitsDir "$ActivityId.json"
    log "Reading $SplitsFile ..."
    $Splits = Get-Json $SplitsFile -AsHashtable
    done

    $Laps = $Splits.lapDTOs[0].lengthDTOs

    log "Finding continous intervals for $ActivityId..."
    $Results = Find-ContinousIntervals $Laps
    log "$($Results.Count) intervals found."

    if ($Results.Count -eq 0) {
        return
    }

    $ActivityFile = Join-Path $ActivityDir "$ActivityId.json"
    log "Reading $ActivityFile ..."
    $Activity = Get-Json $ActivityFile -AsHashtable
    done

    Update-Routines $Workbook $ActivityId $Activity $Laps $Results

    log "Activity $ActivityId updated."
}

log "Reading $IndexFile ..."
$Index = Get-Json $IndexFile
done

log "Reading file $UpdatedFile ..."
$Updated = Get-JsonDefault $UpdatedFile @{ } -AsHashtable
$UpdatedCount = 0
done

log "Updating activities..."
[array]::reverse($Index)
foreach ($Activity in $Index) {
    $ActivityId = [string]$Activity.activityId
    if ($Updated.ContainsKey($ActivityId)) {
        continue;
    }
    $Updated[$ActivityId] = $true
    $UpdatedCount++
    if (!$ExcelApplication) {
        log "Launching Excel..."
        $ExcelApplication = New-Object Microsoft.Office.Interop.Excel.ApplicationClass
        # $ExcelApplication.Visible = $true
        
        log "Opening $RoutinesFile ..."
        $Workbook = $ExcelApplication.Workbooks.Open($RoutinesFile)
        done
    }
    Update-Activity $Workbook $ActivityId
}
log "$UpdatedCount activities updated."

if ($ExcelApplication) {
    log "Saving $RoutinesFile ..."
    $Workbook.Close($true)
    $Workbook = $null
    done

    log "Closing Excel..."
    $ExcelApplication.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ExcelApplication) | Out-Null
    $ExcelApplication = $null
    done
}

if ($UpdatedCount) {
    log "Writing $UpdatedFile ..."
    Set-Json $UpdatedFile $Updated
    done
}

trap {
    if ($Workbook) {
        $Workbook.Close($false)
    }
    if ($ExcelApplication) {
        $ExcelApplication.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ExcelApplication) | Out-Null
    }
    $_
    exit -1
}

# http://import-powershell.blogspot.com/2012/03/excel-part-1.html
# https://docs.microsoft.com/en-us/office/vba/api/overview/excel/object-model
