# play once
Add-Type -AssemblyName System.Speech
$SpeechSynth = New-Object System.Speech.Synthesis.SpeechSynthesizer
$SpeechSynth.SelectVoice("Microsoft Zira Desktop")
$Browser = New-Object System.Net.WebClient
$Browser.Proxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$CatFact = (ConvertFrom-Json (Invoke-WebRequest -Verbose -Uri https://catfact.ninja/fact -UseBasicParsing))
$CatFact.fact
$SpeechSynth.Speak("Did you know ? $($CatFact.fact)")


#loop cat facts
Add-Type -AssemblyName System.Speech
$SpeechSynth = New-Object System.Speech.Synthesis.SpeechSynthesizer
$SpeechSynth.SelectVoice("Microsoft Zira Desktop")
$Browser = New-Object System.Net.WebClient
$Browser.Proxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$on = $true
while ($true -eq $on) {
    try {
        $CatFact = (ConvertFrom-Json (Invoke-WebRequest -Uri "https://catfact.ninja/fact" -UseBasicParsing)).fact
        Write-Host "Cat Fact: $CatFact"
        $SpeechSynth.Speak("Did you know? $CatFact")
    } catch {
        Write-Host "Failed to retrieve cat fact: $_"
    }
    Start-Sleep -Seconds 30  # Wait 10 seconds before fetching the next fact
}


# dog facts

Add-Type -AssemblyName System.Speech
$SpeechSynth = New-Object System.Speech.Synthesis.SpeechSynthesizer
$SpeechSynth.SelectVoice("Microsoft Zira Desktop")
$Browser = New-Object System.Net.WebClient
$Browser.Proxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$on = $true
while ($true -eq $on) {
    try {
        $DogFact = (ConvertFrom-Json (Invoke-WebRequest -Verbose -Uri https://dogapi.dog/api/v2/facts?limit=1 -UseBasicParsing)).data[0].attributes.body
        $SpeechSynth.Speak("Did you know? $DogFact")
    } catch {
        Write-Host "Failed to retrieve dog fact: $_"
    }
    Start-Sleep -Seconds 30  # Wait 10 seconds before fetching the next fact
}