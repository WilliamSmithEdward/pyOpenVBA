function Take-Screenshot {
    param (
        [Parameter(Mandatory=$true)]
        [string]$OutputPath
    )

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    # Get the screen dimensions
    $screen = [System.Windows.Forms.Screen]::PrimaryScreen
    $bounds = $screen.Bounds

    # Create a bitmap of the screen size
    $bitmap = New-Object System.Drawing.Bitmap $bounds.Width, $bounds.Height

    # Create a graphics object from the bitmap
    $graphic = [System.Drawing.Graphics]::FromImage($bitmap)

    # Copy the screen to the bitmap
    $graphic.CopyFromScreen($bounds.X, $bounds.Y, 0, 0, $bounds.Size)

    # Save the bitmap as a file
    $bitmap.Save($OutputPath)

    # Dispose of the objects
    $graphic.Dispose()
    $bitmap.Dispose()

    Write-Host "Screenshot saved to: $OutputPath"
}
