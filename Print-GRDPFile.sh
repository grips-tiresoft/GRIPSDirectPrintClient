#!/bin/zsh

# Print-GRDPFile.sh - macOS version using CUPS printing
# Usage: ./Print-GRDPFile.sh -i <InputFile> [-c <configFile>] [-u <userConfigFile>]

# Parse command line arguments
INPUT_FILE=""
CONFIG_FILE=""
USER_CONFIG_FILE=""

while [[ $# -gt 0 ]]; do
    case $1 in
        -i|--input)
            INPUT_FILE="$2"
            shift 2
            ;;
        -c|--config)
            CONFIG_FILE="$2"
            shift 2
            ;;
        -u|--userconfig)
            USER_CONFIG_FILE="$2"
            shift 2
            ;;
        *)
            echo "Unknown option: $1"
            exit 1
            ;;
    esac
done

if [[ -z "$INPUT_FILE" ]]; then
    echo "Error: Input file is required. Usage: $0 -i <InputFile>"
    exit 1
fi

# Get the directory containing the script
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"

# Set up jq - prefer bundled version, fallback to system version
if [[ -f "$SCRIPT_DIR/jq" ]]; then
    JQ="$SCRIPT_DIR/jq"
else
    JQ=$(which jq 2>/dev/null || echo "jq")
fi

# Set default config files if not provided
if [[ -z "$CONFIG_FILE" ]]; then
    CONFIG_FILE="$SCRIPT_DIR/config-macos.json"
fi

if [[ -z "$USER_CONFIG_FILE" ]]; then
    USER_CONFIG_FILE="$HOME/Library/Application Support/GRIPSDirectPrint/userconfig-macos.json"
fi

# Global variables
declare -A CONFIG
declare -A LANGUAGE_STRINGS

# Function to parse JSON (requires jq)
parse_json() {
    local file="$1"
    local key="$2"
    "$JQ" -r "$key" "$file" 2>/dev/null
}

# Function to load configuration
get_config() {
    if [[ ! -f "$CONFIG_FILE" ]]; then
        echo "Error: Config file not found: $CONFIG_FILE"
        exit 1
    fi

    # Load main config
    CONFIG[Version]=$(parse_json "$CONFIG_FILE" '.Version')
    CONFIG[ReleaseApiUrl]=$(parse_json "$CONFIG_FILE" '.ReleaseApiUrl')
    CONFIG[ReleaseCheckDelay]=$(parse_json "$CONFIG_FILE" '.ReleaseCheckDelay')
    CONFIG[TranscriptMaxAgeDays]=$(parse_json "$CONFIG_FILE" '.TranscriptMaxAgeDays')
    CONFIG[UsePrereleaseVersion]=$(parse_json "$CONFIG_FILE" '.UsePrereleaseVersion')
    
    # Merge user config if it exists
    if [[ -f "$USER_CONFIG_FILE" ]]; then
        # Override with user config values
        local user_version=$(parse_json "$USER_CONFIG_FILE" '.Version')
        [[ -n "$user_version" && "$user_version" != "null" ]] && CONFIG[Version]="$user_version"
        
        local user_prerelease=$(parse_json "$USER_CONFIG_FILE" '.UsePrereleaseVersion')
        [[ -n "$user_prerelease" && "$user_prerelease" != "null" ]] && CONFIG[UsePrereleaseVersion]="$user_prerelease"
        
        local user_api_url=$(parse_json "$USER_CONFIG_FILE" '.ReleaseApiUrl')
        [[ -n "$user_api_url" && "$user_api_url" != "null" ]] && CONFIG[ReleaseApiUrl]="$user_api_url"
        
        local user_check_delay=$(parse_json "$USER_CONFIG_FILE" '.ReleaseCheckDelay')
        [[ -n "$user_check_delay" && "$user_check_delay" != "null" ]] && CONFIG[ReleaseCheckDelay]="$user_check_delay"
        
        local user_max_age=$(parse_json "$USER_CONFIG_FILE" '.TranscriptMaxAgeDays')
        [[ -n "$user_max_age" && "$user_max_age" != "null" ]] && CONFIG[TranscriptMaxAgeDays]="$user_max_age"
    fi
}

# Function to get script version
get_script_version() {
    get_config
    echo "Script version: ${CONFIG[Version]}"
}

# Function to check for updates
update_check() {
    echo "Checking for updates..."
    
    local release_api_url="${CONFIG[ReleaseApiUrl]}"
    local use_prerelease="${CONFIG[UsePrereleaseVersion]}"
    
    if [[ "$use_prerelease" == "true" ]]; then
        echo "Checking for latest release (including prereleases)..."
        # Get all releases
        local all_releases=$(curl -s "${release_api_url/\/latest/}" 2>/dev/null)
        local latest_release=$(echo "$all_releases" | "$JQ" -r '.[0]' 2>/dev/null)
    else
        echo "Checking for latest stable release only..."
        local latest_release=$(curl -s "$release_api_url" 2>/dev/null)
    fi
    
    local release_version=$(echo "$latest_release" | "$JQ" -r '.tag_name' 2>/dev/null | sed 's/^v//')
    local current_version="${CONFIG[Version]}"
    
    # Check if we got a valid version
    if [[ -z "$release_version" || "$release_version" == "null" ]]; then
        echo "Unable to check for updates (API error or network issue)"
        return
    fi
    
    # Simple version comparison
    if [[ "$(printf '%s\n' "$release_version" "$current_version" | sort -V | tail -n1)" != "$current_version" ]]; then
        echo "Update available: $release_version (current: $current_version)"
        
        # Look for .pkg asset in release
        local pkg_download_url=$(echo "$latest_release" | "$JQ" -r '.assets[] | select(.name | test("GRIPSDirectPrint.*\\.pkg$"; "i")) | .browser_download_url' 2>/dev/null | head -n1)
        
        if [[ -z "$pkg_download_url" || "$pkg_download_url" == "null" ]]; then
            echo "Error: Could not find .pkg installer in release assets"
            return
        fi
        
        local temp_pkg="/tmp/grdp_update_$$.pkg"
        
        echo "Downloading installer from: $pkg_download_url"
        curl -sL "$pkg_download_url" -o "$temp_pkg"
        
        if [[ ! -f "$temp_pkg" ]]; then
            echo "Error: Failed to download installer"
            return
        fi
        
        # Create update signal file with path to installer
        local update_signal_file="$(cache_dir)/update_ready.txt"
        echo "$temp_pkg" > "$update_signal_file"
        
        echo "Update downloaded successfully. Will install on next run."
    else
        echo "No update required. Current version ($current_version) is up to date."
    fi
}

# Function to perform the update
update_release() {
    local update_signal_file="$(cache_dir)/update_ready.txt"
    
    if [[ ! -f "$update_signal_file" ]]; then
        echo "Error: Update signal file not found: $update_signal_file"
        return
    fi
    
    local pkg_file=$(cat "$update_signal_file")
    
    if [[ -z "$pkg_file" || ! -f "$pkg_file" ]]; then
        echo "Error: Installer package not found: $pkg_file"
        rm -f "$update_signal_file"
        return
    fi
    
    echo "$(date '+%Y-%m-%d %H:%M:%S') Installing update from $pkg_file"
    
    # Remove quarantine attribute
    echo "$(date '+%Y-%m-%d %H:%M:%S') Removing quarantine attribute..."
    xattr -d com.apple.quarantine "$pkg_file" 2>/dev/null || true
    
    # Install using osascript with admin privileges
    echo "$(date '+%Y-%m-%d %H:%M:%S') Launching installer (requires administrator password)..."
    
    osascript -e "do shell script \"installer -pkg '${pkg_file}' -target /\" with administrator privileges with prompt \"GRIPS Direct Print wants to install a new version.\"" 2>&1
    local install_result=$?
    
    if [[ $install_result -eq 0 ]]; then
        echo "$(date '+%Y-%m-%d %H:%M:%S') Update installed successfully."
        rm -f "$update_signal_file"
        rm -f "$pkg_file"
        get_script_version
        echo "$(date '+%Y-%m-%d %H:%M:%S') Script updated to version ${CONFIG[Version]}."
    else
        echo "$(date '+%Y-%m-%d %H:%M:%S') Error: Installation failed or was cancelled by user."
        rm -f "$update_signal_file"
        rm -f "$pkg_file"
        return 1
    fi
}

# Function to get last update check time
get_last_update_check_time() {
    local last_check_file="$(cache_dir)/last_update_check.txt"
    
    if [[ -f "$last_check_file" ]]; then
        cat "$last_check_file"
    else
        echo "0"
    fi
}

# Function to set last update check time
set_last_update_check_time() {
    local last_check_file="$(cache_dir)/last_update_check.txt"
    date +%s > "$last_check_file"
}

# Function to get unique filename
get_unique_filename() {
    local filepath="$1"
    
    if [[ ! -e "$filepath" ]]; then
        echo "$filepath"
        return
    fi
    
    local dir=$(dirname "$filepath")
    local base_name=$(basename "$filepath")
    local name="${base_name%.*}"
    local ext="${base_name##*.}"
    
    # Handle files without extensions
    if [[ "$name" == "$base_name" ]]; then
        ext=""
    fi
    
    local counter=1
    if [[ -n "$ext" ]]; then
        while [[ -e "$dir/$name ($counter).$ext" ]]; do
            ((counter++))
        done
        echo "$dir/$name ($counter).$ext"
    else
        while [[ -e "$dir/$name ($counter)" ]]; do
            ((counter++))
        done
        echo "$dir/$name ($counter)"
    fi
}

# Function to check if printer exists using CUPS
test_printer_exists() {
    local printer_name="$1"
    lpstat -p "$printer_name" &>/dev/null
    return $?
}

# Function to load language strings
get_language_strings() {
    local lang_file="$SCRIPT_DIR/languages.json"
    
    if [[ ! -f "$lang_file" ]]; then
        echo "Warning: Language file not found: $lang_file"
        return
    fi
    
    # Get OS language
    local os_culture="${LANG%%.*}"
    os_culture="${os_culture//_/-}"
    echo "OS Language: $os_culture"
    
    # Try to get language strings
    local lang_data=$("$JQ" -r ".\"$os_culture\"" "$lang_file" 2>/dev/null)
    
    if [[ "$lang_data" == "null" || -z "$lang_data" ]]; then
        # Try language-only match
        local lang_only="${os_culture%%-*}"
        lang_data=$("$JQ" -r "keys[] | select(startswith(\"$lang_only\"))" "$lang_file" | head -n1)
        
        if [[ -n "$lang_data" ]]; then
            echo "Using language strings for: $lang_data"
            "$JQ" -r ".\"$lang_data\"" "$lang_file"
        else
            echo "No matching language found. Using en-US as fallback."
            "$JQ" -r '.["en-US"]' "$lang_file"
        fi
    else
        echo "Using language strings for: $os_culture"
        echo "$lang_data"
    fi
}

# Function to select alternative printer using osascript (AppleScript)
select_alternative_printer() {
    local missing_printer="$1"
    
    # Get list of available printers
    local printers=$(lpstat -p | awk '{print $2}')
    
    if [[ -z "$printers" ]]; then
        osascript -e 'display dialog "No printers available on this system." buttons {"OK"} default button "OK" with icon stop'
        return 1
    fi
    
    # Create printer list for dialog
    local printer_list=$(echo "$printers" | tr '\n' ',' | sed 's/,$//')
    
    # Show selection dialog
    local selected=$(osascript <<EOF
tell application "System Events"
    set printerList to {"$printer_list"}
    set selectedPrinter to choose from list (paragraphs of printerList) with prompt "Printer '$missing_printer' not found.\\n\\nSelect an alternative printer:" default items {item 1 of (paragraphs of printerList)}
    if selectedPrinter is false then
        return ""
    else
        return item 1 of selectedPrinter
    end if
end tell
EOF
)
    
    if [[ -n "$selected" ]]; then
        echo "$selected"
        return 0
    else
        return 1
    fi
}

# Function to print PDF using CUPS lp command
print_pdf_cups() {
    local pdf_file="$1"
    local printer_name="$2"
    local output_bin="$3"
    local additional_args="$4"
    
    # Check if printer exists
    if ! test_printer_exists "$printer_name"; then
        echo "Warning: Printer '$printer_name' not found."
        local alternative_printer=$(select_alternative_printer "$printer_name")
        
        if [[ -z "$alternative_printer" ]]; then
            echo "No alternative printer selected. Skipping print job for $pdf_file"
            return 1
        fi
        
        echo "Using alternative printer: $alternative_printer"
        printer_name="$alternative_printer"
    fi
    
    # Build lp command
    local lp_options=""
    
    # Add output bin if specified
    if [[ -n "$output_bin" ]]; then
        lp_options="-o outputbin=$output_bin"
    fi
    
    # Add additional arguments
    if [[ -n "$additional_args" ]]; then
        lp_options="$lp_options $additional_args"
    fi
    
    # Print the PDF
    echo "Printing $pdf_file to printer '$printer_name' with options: $lp_options"
    
    if [[ -n "$lp_options" ]]; then
        lp -d "$printer_name" $lp_options "$pdf_file"
    else
        lp -d "$printer_name" "$pdf_file"
    fi
    
    local exit_code=$?
    
    if [[ $exit_code -eq 0 ]]; then
        echo "Print job submitted successfully for $pdf_file"
    else
        echo "Error: Print job failed with exit code $exit_code"
        return $exit_code
    fi
}

cache_dir() {
    local dir="$HOME/Library/Caches/com.grips.directprint"
    mkdir -p "$dir"
    echo "$dir"
}

# Main execution
main() {
    # Load configuration
    get_config
    
    echo "Processing file: $INPUT_FILE"
    
    # Check if file exists
    if [[ ! -f "$INPUT_FILE" ]]; then
        echo "Error: Input file not found: $INPUT_FILE"
        exit 1
    fi
    
    # Check if it's a .grdp file
    if [[ "${INPUT_FILE:l}" == *.grdp ]]; then
        # Create temp folder
        local temp_folder="/tmp/grdp_$$"
        mkdir -p "$temp_folder"
        
        # Extract the .grdp (zip) file
        local zip_file="$temp_folder/$(basename "${INPUT_FILE%.grdp}.zip")"
        cp "$INPUT_FILE" "$zip_file"
        unzip -q "$zip_file" -d "$temp_folder"
        
        # Find printsettings.json
        local settings_file="$temp_folder/printsettings.json"
        
        if [[ ! -f "$settings_file" ]]; then
            echo "Error: printsettings.json not found in archive."
            rm -rf "$temp_folder"
            exit 1
        fi
        
        # Process each entry in printsettings.json
        local entries_count=$("$JQ" 'length' "$settings_file")
        
        for ((i=0; i<entries_count; i++)); do
            local pdf_filename=$("$JQ" -r ".[${i}].Filename" "$settings_file")
            local printer=$("$JQ" -r ".[${i}].Printer" "$settings_file")
            local output_bin=$("$JQ" -r ".[${i}].OutputBin" "$settings_file")
            local add_args=$("$JQ" -r ".[${i}].AdditionalArgs" "$settings_file")
            
            # Handle null values
            [[ "$output_bin" == "null" ]] && output_bin=""
            [[ "$add_args" == "null" ]] && add_args=""
            
            local file_path="$temp_folder/$pdf_filename"
            
            if [[ ! -f "$file_path" ]]; then
                echo "Warning: File $pdf_filename not found in archive, skipping."
                continue
            fi
            
            if [[ "${file_path:l}" == *.pdf ]]; then
                # Print PDF using CUPS
                print_pdf_cups "$file_path" "$printer" "$output_bin" "$add_args"
            else
                # Open non-PDF file with default application (e.g., .eml files open with Thunderbird)
                local downloads_folder="$HOME/Downloads"
                # Use only the basename to avoid path issues
                local base_filename=$(basename "$pdf_filename")
                local unique_filepath=$(get_unique_filename "$downloads_folder/$base_filename")
                cp "$file_path" "$unique_filepath"
                echo "Opening file: $unique_filepath"
                
                # Only call open if not running from app bundle
                if [[ -z "$GRDP_NO_OPEN" ]]; then
                    open "$unique_filepath"
                fi
            fi
        done
        
        # Clean up temp folder
        rm -rf "$temp_folder"
        
        # Remove old download files
        local downloads_folder="$HOME/Downloads"
        local max_age_days="${CONFIG[TranscriptMaxAgeDays]:-7}"
        
        # Remove old .eml files
        find "$downloads_folder" -name "NewEmail*.eml" -mtime +$max_age_days -delete 2>/dev/null
        
        # Remove old .sig files
        find "$downloads_folder" -name "*.sig" -mtime +$max_age_days -delete 2>/dev/null
        
        # Remove old .grdp files
        find "$downloads_folder" -name "*.grdp" -mtime +$max_age_days -delete 2>/dev/null
        
    else
        # Normal PDF file - print to default printer
        local default_printer=$(lpstat -d | awk '{print $NF}')
        echo "Printing $INPUT_FILE to default printer: $default_printer"
        lp "$INPUT_FILE"
    fi
    
    # Check for updates
    local update_signal_file="$(cache_dir)/update_ready.txt"
    
    if [[ -f "$update_signal_file" ]]; then
        update_release
    else
        local last_check_time=$(get_last_update_check_time)
        local current_time=$(date +%s)
        local elapsed=$((current_time - last_check_time))
        local release_check_delay="${CONFIG[ReleaseCheckDelay]:-3600}"
        
        if [[ $elapsed -ge $release_check_delay ]]; then
            set_last_update_check_time
            echo "Time since last update check: $elapsed seconds. Checking for updates..."
            update_check
        else
            echo "Last update check was $elapsed seconds ago. Skipping update check."
        fi
    fi
}

# Run main function
main

exit 0
