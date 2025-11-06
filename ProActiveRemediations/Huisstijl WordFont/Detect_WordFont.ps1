# Variables
$WordEXE = "C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE"
$systemLocale = Get-WinSystemLocale
$desiredFontName = "Verdana"    # Change this to the desired font
$desiredFontSize = 9         # Change this to the desired size
# Check if Word is installed
if (test-path -Path "$WordEXE") {
    Write-Host "Word is installed."
    # Check if current system language is English (US)
    if ($systemLocale -eq "en-US") {
        Write-Host "The system language is English."
        # Check if Word is running
        $wordProcesses = Get-Process | Where-Object { $_.ProcessName -eq "WINWORD" }
        if ($wordProcesses.Count -gt 0) {
            Write-Host "Word is running."
            Exit 0
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
            $currentFontName = $document.Styles.Item("Normal").Font.Name
            $currentFontSize = $document.Styles.Item("Normal").Font.Size
            if ($currentFontName -ne $desiredFontName -or $currentFontSize -ne $desiredFontSize) {
                Write-Host "The font is incorrect."
                
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
                # 2nd Write-Host for visibility in Intune 
                Write-Host "The font is incorrect, proceeding to Remediation"
                Exit 1
            } else {
                Write-Host "The font is correctly set."
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
                # 2nd Write-Host for visibility in Intune
                Write-Host "The font is correctly set."
                Exit 0
            }
        }
    }
    elseif ($systemLocale -eq "nl-NL") {
        Write-Host "The system language is Dutch."
        # Check if Word is running
        $wordProcesses = Get-Process | Where-Object { $_.ProcessName -eq "WINWORD" }
        if ($wordProcesses.Count -gt 0) {
            Write-Host "Word is running."
            Exit 0
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
            $currentFontName = $document.Styles.Item("Standaard").Font.Name
            $currentFontSize = $document.Styles.Item("Standaard").Font.Size
            if ($currentFontName -ne $desiredFontName -or $currentFontSize -ne $desiredFontSize) {
                Write-Host "The font is incorrect."
                
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
                # 2nd Write-Host for visibility in Intune 
                Write-Host "The font is incorrect, proceeding to Remediation"
                Exit 1
            } else {
                Write-Host "The font is correctly set."
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
                # 2nd Write-Host for visibility in Intune
                Write-Host "The font is correctly set."
                Exit 0
            }
        }
    } else {
        Write-Host "The system language is not English or Dutch."
        Exit 0
    }
} else {
    Write-Host "Word is not installed."
    Exit 0
}