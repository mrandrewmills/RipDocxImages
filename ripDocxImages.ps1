<#
.SYNOPSIS
    Extracts images from a .docx file.

.DESCRIPTION
    This script extracts all embedded images from a specified Microsoft Word .docx file.
    It creates a new directory named after the source file to store the extracted images.

.PARAMETER Path
    The full path to the .docx file.

.EXAMPLE
    .\ripDocxImages.ps1 -Path "C:\docs\MyDocument.docx"
    This command will extract images from MyDocument.docx and save them in a folder named "C:\docs\MyDocument_images".
#>
param(
    [Parameter(Mandatory=$true, Position=0)]
    [string]$Path
)

# 1. Validate input file
if (-not (Test-Path -Path $Path -PathType Leaf)) {
    Write-Error "File not found: $Path"
    return
}

if ($([System.IO.Path]::GetExtension($Path)) -ne ".docx") {
    Write-Error "Invalid file type. Please provide a .docx file."
    return
}

$fileInfo = Get-Item -Path $Path
$baseName = $fileInfo.BaseName
$directory = $fileInfo.DirectoryName

# 2. Create output directory
$outputDir = Join-Path -Path $directory -ChildPath "${baseName}_images"
if (-not (Test-Path -Path $outputDir)) {
    New-Item -Path $outputDir -ItemType Directory | Out-Null
    Write-Host "Created output directory: $outputDir"
} else {
    Write-Host "Output directory already exists: $outputDir"
}

# 3. Create a temporary directory for extraction
$tempDir = Join-Path -Path $env:TEMP -ChildPath ([System.Guid]::NewGuid().ToString())
New-Item -Path $tempDir -ItemType Directory | Out-Null

try {
    # 4. Expand the .docx file (which is a zip archive)
    Expand-Archive -Path $Path -DestinationPath $tempDir -ErrorAction Stop

    $mediaPath = Join-Path -Path $tempDir -ChildPath "word\media"

    # 5. Check for images and copy them
    if (Test-Path -Path $mediaPath) {
        $images = Get-ChildItem -Path $mediaPath
        if ($images) {
            Copy-Item -Path $images.FullName -Destination $outputDir
            Write-Host "Successfully extracted $($images.Count) image(s) to $outputDir"
        } else {
            Write-Warning "No images found in the document."
            # Clean up the empty output directory if no images were found
            Remove-Item -Path $outputDir -Recurse -Force
            Write-Host "Removed empty output directory."
        }
    } else {
        Write-Warning "No media folder found in the document. No images to extract."
        # Clean up the empty output directory if no media folder was found
        Remove-Item -Path $outputDir -Recurse -Force
        Write-Host "Removed empty output directory."
    }
}
catch {
    Write-Error "An error occurred during extraction: $_"
}
finally {
    # 6. Clean up the temporary directory
    if (Test-Path -Path $tempDir) {
        Remove-Item -Path $tempDir -Recurse -Force
    }
}
