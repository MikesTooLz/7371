try {
    [System.Media.SystemSounds]::Asterisk.Play()
} catch {
    Write-Host "Failed to play the system sound."
}

Add-Type -AssemblyName System.Speech

# Create a SpeechSynthesizer object
$speechSynthesizer = New-Object -TypeName System.Speech.Synthesis.SpeechSynthesizer

# The message you want to convert to speech
$message = "Hello, this is a test message."

# Speak the message
$speechSynthesizer.Speak($message)
