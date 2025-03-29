param(
    [Parameter(Mandatory = $false, HelpMessage = "Enter the exact recycle bin item name to restore (ex. 'secret_plans.docx')")]
    [string]$ItemNameToRestore,

    [Parameter(Mandatory = $false, HelpMessage = "Enter the verb name to execute (defaults to 'restore')")]
    [string]$Verb = "restore"
)

function Show-Usage {
    Write-Output "`nUsage: .\Unrecycle.ps1 [-ItemNameToRestore <RecycleBinItemName>] [-Verb <VerbName>]"
    Write-Output "Examples:"
    Write-Output "  (Check recycle bin contents)      .\Unrecycle.ps1" 
    Write-Output "  (Restore an item using default)   .\Unrecycle.ps1 -ItemNameToRestore 'secret_plans.docx'"
    Write-Output "  (Use a different verb)            .\Unrecycle.ps1 -ItemNameToRestore 'secret_plans.docx' -Verb 'undelete'"
}

# Create a Shell.Application COM object.
$shell = New-Object -ComObject Shell.Application

# Access the Recycle Bin using namespace 0xA (decimal 10).
$recycleBin = $shell.Namespace(0xA)

# Get all items in the Recycle Bin.
$items = $recycleBin.Items()

if ($items.Count -eq 0) {
    Write-Output "`nThe Recycle Bin is empty.`n"
    exit
}

# Print out the contents of the Recycle Bin with their restore/original location.
Write-Output "`nItems in the Recycle Bin:"
foreach ($item in $items) {
    # Adjust the column index if necessary – index 1 is assumed to be the restore location.
    $restoreLocation = $recycleBin.GetDetailsOf($item, 1)
    Write-Output "  Name: $($item.Name)    ($restoreLocation)"
}

# Check if user provided an item to restore
if (-not $ItemNameToRestore) {
    Write-Output "`nTo restore any of these items, use the -ItemNameToRestore argument"
    Write-Output "ex. .\Unrecycle.ps1 -ItemNameToRestore 'secret_plans.docx'`n"
    exit
}

$foundItem = $null

foreach ($item in $items) {
    if ($item.Name -eq $ItemNameToRestore) {
        $foundItem = $item
        break
    }
}

if ($null -eq $foundItem) {
    Write-Output "`nCould not find an item named '$ItemNameToRestore' in the Recycle Bin.`n"
    exit
}

$restoreExecuted = $false

# Find the verb (default is "restore", or the value of the -Verb argument)
foreach ($verbObj in $foundItem.Verbs()) {
    if ($verbObj.Name -match $Verb) {
        Write-Output "`nExecuting verb '$($verbObj.Name)' on item: $($foundItem.Name)"
        $verbObj.DoIt()
        $restoreExecuted = $true
        break
    }
}

if (-not $restoreExecuted) {
    Write-Output "`nThe command was not executed. The item may have already been restored or the appropriate verb was not found.`n"
    Write-Output "Available verbs for $($foundItem.Name):"
    foreach ($verbObj in $foundItem.Verbs()) {
        Write-Output " - $($verbObj.Name)"
    }
}

