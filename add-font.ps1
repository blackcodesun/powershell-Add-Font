# add-font.ps1
# Mike Stine, 7/11/2018
# This powershell script installs a font file into Windows
# Window recognizes fonts dragged and dropped into the Windows font folder, but does not recognize fonts copied through cli or script. This is due to the windows explorer shell triggering a font registration process.  This script evokes the same windows explorer shell trigger to install fonts

# This is the folder that contains the font(s) to install, or this can be a direct path to a font
$srcPath = "C:\temp\add-font\*"

# Find files with file extensions *.fon,*.fnt,*.otf,*.ttc,*.ttf
$fontFiles = Get-ChildItem -Path "$srcPath" -Include *.fon,*.fnt,*.otf,*.ttc,*.ttf -Recurse

# Creates new Shell.Application COM object
$shellObject = New-Object -ComObject Shell.Application

# Sets object's namespace to CLSID hex '0X14' (Windows font folder and registration trigger)
$windowsFontFolder = $shellObject.Namespace(0x14)

# Move and register font files.
foreach($fontFile in $fontFiles) {
  # test if font is already installed 
  if ( !(Test-Path "${env:windir}\Fonts\$($fontFile.Name)") ) {
    # install font
    $windowsFontFolder.CopyHere($fontFile.fullname)
  }
}

