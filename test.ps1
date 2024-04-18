try {
    [System.Media.SystemSounds]::Asterisk.Play()
} catch {
    Write-Host "Failed to play the system sound."
}
