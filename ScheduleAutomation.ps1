<##
Script takes in a release date from a user to determine that month's release schedule and put it into a new excel spreadsheet.
Can be run from command prompt or PowerShell ISE

Developer: Sydney Brookstein
8/11/2019
#>

#params from command line - takes in just end date. Must be in mm/dd/yyyy format.
param(
    [string] $userEnd
)

###Define fucntion###
#function reutrns true if value is of form mm/dd/yyyy
function Is-Date ($Value) {
    return $Value -match "^[0-1]\d/\d\d/\d\d\d\d$"
}

#checks to make sure user input date in mm/dd/yyyy format. If they didn't, prompt user to try again.
while(-not(Is-Date $userEnd)){ 
    $userEnd = Read-Host -Prompt 'Please input release date: (MM/DD/YY) '
}

###Define Vars###
$EndDate = [DateTime] $userEnd

#check if release date from user is a weekend. User can choose to continue or quit program.
if($EndDate.DayOfWeek -gt 5){
    $ans = Read-Host -Prompt 'The date you chose as the release date is a weekend. Are you sure you want to continue? (Y/N) '
    while($ans -ne "y" -and $ans -ne "n"){
        $ans = Read-Host -Prompt 'Invalid input. Try again. '
    }
    if($ans -eq "n"){
        exit
    }
}

#get month info from release date
$ReleaseMonth = (Get-Culture).DateTimeFormat.GetMonthName($EndDate.Month)
$PrevMonth = (Get-Culture).DateTimeFormat.GetMonthName($EndDate.Month - 1)
$2MoPrev = (Get-Culture).DateTimeFormat.GetMonthName($EndDate.Month - 2)

#titles of activities for the month
$PREScope = "PR&E Scope"
$QUATProcurement = "QUAT Environment Procurement (Includes $2MoPrev Code)"
$QUATDB = "QUAT Database Setup ($2MoPrev)"
$SmokeTest1 = "Smoke Test QUAT ($2MoPrev)"
$DevCodeMerge = "Dev Code Merge - $PrevMonth Release"
$RMBuild = "RM Build Deploy $PrevMonth Release"
$SmokeTest2 = "Smoke Test QUAT ($PrevMonth)"
$DevPRE = "Dev PR&E's - $ReleaseMonth Release"
$QAValidation = "QA Validation PR&E's"
$SoftCodeFreeze = "Soft Code Freeze (All functional changes delivered and Tested in QUAT - Includes project and PRE work) - Changes allowed by approval only"
$ReleaseRegression = "Release Regression Testing"
$UATValidation = "UAT Validation"
$ReleaseSignoff = "Release Signoff (UAT)/Release Notes Due"
$FortifyScans = "Fortify Scans/Release Prep/Release"

#array of activities
$ActivityArray = @($FortifyScans, $ReleaseSignoff, $UATValidation, $ReleaseRegression, $SoftCodeFreeze, $QAValidation, $DevPRE, $SmokeTest2, $RMBuild,
 $DevCodeMerge, $SmokeTest1, $QUATDB, $QUATProcurement, $PREScope)

 #define number of cell with the release in it
 $endCell = $ActivityArray.Count + 2

#end dates array uses WORKDAY function from excel (skips weekends)
#adjust numbers here if the amount of time to work on that activity changes 
$EndDates = @(
    $EndDate,
    "=WORKDAY(C$endCell ,-10)",
    "=WORKDAY(C$endCell ,-11)",
    "=WORKDAY(C$endCell ,-11)",
    $EndDate,
    "=WORKDAY(C$endCell ,-26)",
    "=WORKDAY(C$endCell ,-36)",
    "=WORKDAY(C$endCell ,-41)",
    "=WORKDAY(C$endCell ,-42)",
    "=WORKDAY(C$endCell ,-44)",
    "=WORKDAY(C$endCell ,-46)",
    "=WORKDAY(C$endCell ,-47)",
    "=WORKDAY(C$endCell ,-49)",
    "=WORKDAY(C$endCell ,-51)"
)

#start daes array uses WORKDAY function from excel (skips weekends)
#adjust numbers here if the amount of time to work on that activity changes
$StartDates = @(
    "=WORKDAY(C$endCell ,-9)",
    "=WORKDAY(C$endCell ,-10)",
    "=WORKDAY(C$endCell ,-20)",
    "=WORKDAY(C$endCell ,-20)",
    "=WORKDAY(C$endCell ,-25)",
    "=WORKDAY(C$endCell ,-35)",
    "",
    "=WORKDAY(C$endCell ,-41)",
    "=WORKDAY(C$endCell ,-43)",
    "=WORKDAY(C$endCell ,-45)",
    "=WORKDAY(C$endCell ,-46)",
    "=WORKDAY(C$endCell ,-48)",
    "=WORKDAY(C$endCell ,-53)",
    "=WORKDAY(C$endCell ,-53)"
)

###start writing to excel###
#open excel
$excel = New-Object -ComObject excel.application

#make workbook visible 
$excel.visible = $True

#add sheet to workbook
$workbook = $excel.Workbooks.Add()

#make variable to reference worksheet
$worksheet = $workbook.Worksheets.Item(1)

#label title row
$worksheet.Cells.Item(1,1) = 'Activity'
$worksheet.Cells.Item(1,1).Font.Bold = $True
$worksheet.Cells.Item(1,1).Font.Italic = $True

$worksheet.Cells.Item(1,2) = 'Start Date'
$worksheet.Cells.Item(1,2).Font.Bold = $True
$worksheet.Cells.Item(1,2).Font.Italic = $True

$worksheet.Cells.Item(1,3) = 'End Date'
$worksheet.Cells.Item(1,3).Font.Bold = $True
$worksheet.Cells.Item(1,3).Font.Italic = $True

#Label 2nd row with title
$worksheet.Cells.Item(2,1) = "$year Estimated Release Schedule for $ReleaseMonth Release"
$worksheet.Cells.Item(2,1).Font.Bold = $True
$worksheet.Cells.Item(2,1).Font.Italic = $True

#merge cells in 2nd row and make it gray so it's pretty
$MergeCells = $worksheet.Range("a2:c2")
$MergeCells.Select()
$MergeCells.MergeCells = $true
$MergeCells.Cells.Interior.ColorIndex = 15

#fill out each activity, start date, and end date from respective arrays into excel 
for($i = 0; $i -lt $EndDates.Count; $i++){
    $worksheet.Range('c'+ ($endCell-$i)) = $EndDates[$i]
    $worksheet.Range('c'+ ($endCell-$i)).NumberFormat = "m/d/yyyy"

    $worksheet.Range('b'+ ($endCell-$i)) = $StartDates[$i]
    $worksheet.Range('b'+ ($endCell-$i)).NumberFormat = "m/d/yyyy"

    $worksheet.Range('a'+ ($endCell-$i)) = $ActivityArray[$i]
}

#formatting
$usedRange = $worksheet.UsedRange
$usedRange.EntireColumn.AutoFit() | Out-Null