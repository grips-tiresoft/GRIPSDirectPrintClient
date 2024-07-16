# Set PowerShell console to use UTF-8 encoding
$OutputEncoding = [Console]::InputEncoding = [Console]::OutputEncoding =
New-Object System.Text.UTF8Encoding

function DisplayMenuAndReadSelection {
    param (
        [System.Collections.Specialized.OrderedDictionary]$Items,
        [string]$Title,
        [int]$StartIndex = 1
    )

    do {
        Clear-Host  # Clears the console

        # Display the title if provided
        if (-not [string]::IsNullOrWhiteSpace($Title)) {
            Write-Host $Title -ForegroundColor Cyan
            Write-Host ('=' * $Title.Length) -ForegroundColor Cyan  # Optional: underline for the title
        }

        # Display menu options
        $index = $StartIndex
        foreach ($key in $Items.Keys) {
            Write-Host "$index`t$key`t$($Items[$key])"
            $index++
        }

        # Prompt for user input
        $selection = Read-Host "Please select an option ($($StartIndex)-$($StartIndex+$Items.Count-1))"

        # Validate selection
        $selInt = $selection -as [int]
        $isValid = ($selInt -ge $StartIndex) -and ($selInt -le $($StartIndex+$Items.Count-1))

        if (-not $isValid) {
            Write-Host "Invalid selection. Please try again." -ForegroundColor Red
            Start-Sleep -Seconds 2  # Give user time to read the message before clearing
        }

    } while (-not $isValid)

    # Convert selection to corresponding dictionary key
    $selectedKey = ($Items.Keys | Select-Object -Index ($selection - $StartIndex))
    Write-Host "You selected: $selectedKey" -ForegroundColor Green

    # Convert selection to corresponding dictionary key and index
    $selectedKey = ($Items.Keys | Select-Object -Index ($selection - $StartIndex))
    $selectedIndex = [int]$selection - $StartIndex  # Adjust for zero-based index if needed

    # Return both selected key and index
    return @{ "Key" = $selectedKey; "Index" = $selectedIndex }
}
                    
# Function to generate unique shortcut keys for a list of items
function Generate-ChoiceDescriptions {
    param (
        [System.Collections.Specialized.OrderedDictionary]$Items,
        [bool]$UseNumbers = $false
    )

    $usedKeys = @{}
    $choiceDescriptions = [System.Management.Automation.Host.ChoiceDescription[]]@()

    foreach ($item in $Items.Keys) {
        # Find a unique shortcut key for the item
        $shortcutKey = $null
        if ($UseNumbers) {
            $shortcutKey = [string]([array]::IndexOf($Items.Keys, $item) + 1)
        }
        else {
            foreach ($char in $item.ToCharArray()) {
                if (-not $usedKeys.ContainsKey($char) -and $char -match '\w') {
                    $shortcutKey = $char
                    $usedKeys[$char] = $true
                    break
                }
            }
        }

        if ($null -eq $shortcutKey) { $shortcutKey = '_' } # Fallback if no unique character found

        $label = "&$shortcutKey - $item"
        $helpMessage = $Items[$item] # Use the value from the dictionary as the help message
        $choiceDescriptions += New-Object System.Management.Automation.Host.ChoiceDescription $label, $helpMessage
    }

    return $choiceDescriptions
}

# Load the JSON configuration
$configFilePath = ".\Installer\install.json"
$jsonContent = Get-Content -Path $configFilePath -Encoding UTF8 | ConvertFrom-Json 

# Extract countries and prepare the options for PromptForChoice
$Items = [ordered]@{}

foreach ($countryCode in $jsonContent.Countries.PSObject.Properties.Name) {
    $country = $jsonContent.Countries.$countryCode
    $Items.Add("$countryCode", "$($country.Name)")
}

$options = Generate-ChoiceDescriptions -Items $Items

# Present the countries to the user and ask for a selection
$title = "Country Selection"
$message = "Please select a country by typing its code:"
#$selectedOption = $Host.UI.PromptForChoice($title, $message, $options, 0)
$selectionResult = DisplayMenuAndReadSelection -Items $Items -Title $title

# Convert the selection to the country code
$selectedCountryCode = $jsonContent.Countries.PSObject.Properties.Name[$selectionResult.Index]
$selectedCountry = $jsonContent.Countries.$selectedCountryCode

Write-Host "Selected Country: $($selectedCountry.Name) ($selectedCountryCode)"

# Extract database and prepare the options for PromptForChoice
$Items = [ordered]@{}

foreach ($DatabaseName in $selectedCountry.Databases.PSObject.Properties.Name) {
    $Database = $selectedCountry.Databases.$DatabaseName
    $label = "$DatabaseName"
    $helpMessage = ""
    $Items.Add($label, $helpMessage)
}

$options = Generate-ChoiceDescriptions -Items $Items -UseNumbers $true

# Present the database to the user and ask for a selection
$title = "Database Selection"
$message = "Please select a database by typing its code:"
#$selectedOption = $Host.UI.PromptForChoice($title, $message, $options, 0)
if ($selectedCountry.StartIndex) {
    $selectionResult = DisplayMenuAndReadSelection -Items $Items -Title $title -StartIndex $selectedCountry.StartIndex
} else {
    $selectionResult = DisplayMenuAndReadSelection -Items $Items -Title $title
}

# Convert the selection to the country code
$selectedDatabase = $selectedCountry.Databases.PSObject.Properties.Name[$selectionResult.Index]
$selectedBaseURL = $selectedCountry.Databases.$selectedDatabase.BaseURL

Write-Host "Selected Database: $($selectedDatabase) ($selectedBaseURL)"

