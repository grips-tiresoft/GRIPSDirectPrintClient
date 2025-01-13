# If paper size is standard, set units to "HI" (hundredths of an inch)
# If paper size is custom, you can specify units like "MM" (millimeters)

# Function to get printer capabilities and generate JSON payload
function Get-PrinterJsonPayload {
    param (
        [string]$PrinterName
    )

    # Create an instance of the PrinterSettings class
    $printerSettings = New-Object System.Drawing.Printing.PrinterSettings
    $printerSettings.PrinterName = $PrinterName

    if ($printerSettings.IsValid) {
        # Initialize JSON structure
        $jsonPayload = @{
            "version"       = 1
            "description"   = $printerSettings.PrinterName
            "duplex"        = $printerSettings.CanDuplex
            "color"         = $printerSettings.SupportsColor
            "defaultcopies" = 1  # Default copies set to 1 as per example
            "papertrays"    = @()
        }

        # Loop through the paper sources and paper sizes
        foreach ($source in $printerSettings.PaperSources) {
            foreach ($paperSize in $printerSettings.PaperSizes) {
                # If PaperSize is Custom (PaperKind == 0), we need to specify dimensions
                $paperKind = $paperSize.Kind
                $height = $paperSize.Height
                $width = $paperSize.Width
                $units = "HI"  # Default to hundredths of an inch for standard sizes

                # Create paper tray setup
                $paperTray = @{
                    "papersourcekind" = $source.SourceName
                    "paperkind"       = $paperKind
                }

                # If the paper kind is custom, we include height, width, and units
                if ($paperKind -eq 0) {
                    $units = "MM"  # Assume custom sizes use millimeters for units
                    $paperTray["height"] = $height
                    $paperTray["width"]  = $width
                    $paperTray["units"]  = $units
                }

                # Add paper tray setup to the papertrays array
                $jsonPayload["papertrays"] += $paperTray
            }
        }

        # Convert the PowerShell object to JSON format
        return $jsonPayload | ConvertTo-Json -Depth 5
    } else {
        Write-Error "The printer '$PrinterName' is not valid or not found."
    }
}

# Retrieve installed printers using Get-Printer
$installedPrinters = Get-Printer | Select-Object -ExpandProperty Name

foreach ($printer in $installedPrinters) {
    Write-Host "Getting capabilities for printer: $printer" -ForegroundColor Cyan
    $printerSettings = Get-PrinterJsonPayload -PrinterName $printer
    $printerSettings | Out-File "$printer.json"
    
    Write-Host "`n"
}
