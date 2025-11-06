# Variables
$desiredFontName = "Verdana"    # Change this to the desired font
$desiredFontSize = 9         # Change this to the desired size
$systemLocale = Get-WinSystemLocale

if ($systemLocale -eq "en-US"){
    $templatename = "Normal"
}
elseif ($systemLocale -eq "nl-NL") {
    $templatename = "Standaard"
}
else {
    Write-Host "Language is not en-US or nl-NL"
    exit 0
}

# Check if Word is running
$wordProcesses = Get-Process | Where-Object { $_.ProcessName -eq "WINWORD" }
if ($wordProcesses.Count -gt 0) {
    Write-Host "Word is running."
} else {
    Write-Host "Checking if the correct font is set."
    # Start Word application (hidden)
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false  # Ensure Word is invisible
    # Create a new document
    try {
        $document = $word.Documents.Add()
        if ($null -eq $document) {
            throw "The new document could not be created."
        } else {
            Write-Host "New document created."
        }
    } catch {
        Write-Host "Error creating new document: $_"
        # Close Word application
        $word.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
        exit
    }
    # Check if the current default font differs from the desired font
    $currentFontName = $document.Styles.Item($templatename).Font.Name
    $currentFontSize = $document.Styles.Item($templatename).Font.Size
    if ($currentFontName -ne $desiredFontName -or $currentFontSize -ne $desiredFontSize) {
        try {
            # Get the 'Normal' style of the new document
            $normalStyle = $document.Styles.Item($templatename)
            if ($null -eq $normalStyle) {
                throw "The 'Normal' style could not be found in the document."
            } else {
                Write-Host "Normal style found in the document."
            }
            # Change the font and size in the 'Normal' style
            $normalStyle.Font.Name = $desiredFontName
            $normalStyle.Font.Size = $desiredFontSize
            Write-Host "The default font has been updated in the document."
            # Open NormalTemplate as a document
            $normalTemplatePath = $word.NormalTemplate.FullName
            $normalTemplateDoc = $word.Documents.Open($normalTemplatePath)
            if ($null -eq $normalTemplateDoc) {
                throw "The normal template document could not be opened."
            } else {
                Write-Host "NormalTemplate document opened."
            }
            # Get the 'Normal' style of the NormalTemplate document
            $normalTemplateStyle = $normalTemplateDoc.Styles.Item($templatename)
            if ($null -eq $normalTemplateStyle) {
                throw "The 'Normal' style could not be found in the NormalTemplate document."
            } else {
                Write-Host "Normal style found in the NormalTemplate document."
            }
            # Change the font and size in the 'Normal' style of the NormalTemplate document
            $normalTemplateStyle.Font.Name = $desiredFontName
            $normalTemplateStyle.Font.Size = $desiredFontSize
            Write-Host "The default font has been updated in the NormalTemplate document."
            # Save the NormalTemplate document
            $normalTemplateDoc.Save()
            Write-Host "The default font has been saved in the NormalTemplate document."
            # Close the NormalTemplate document
            $normalTemplateDoc.Close($false)
        } catch {
            Write-Host "An error occurred while updating the font: $_"
        }
    } else {
        Write-Host "The current default font is already set to the desired font. No further action needed."
    }
    
    # Try to close the document without saving, with error handling
    try {
        $document.Close($false)
        Write-Host "Document closed."
    } catch {
        Write-Host "Error closing the document: $_"
    }
    
    # Close Word application
    try {
        $word.Quit()
        Write-Host "Word closed."
    } catch {
        Write-Host "Error closing Word: $_"
    }
    # Clean up
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($document) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
}