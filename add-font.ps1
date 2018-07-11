# add-font.ps1
# Mike Stine, 7/11/2018
# Window recognizes fonts you drag and drop into the windows font folder, but does not recognize fonts copied into the folder through cli or script. BECAUSE, the windows explorer shell triggers a registration process through the Shell.Application COM object.  And so will we.

# This is the folder that contains the font(s) to install, or this can be a direct path to a font
$srcPath = "C:\temp\add-font\*"

# Find files with file extensions *.fon,*.fnt,*.otf,*.ttc,*.ttf
$fontFiles = Get-ChildItem -Path "$srcPath" -Include *.fon,*.fnt,*.otf,*.ttc,*.ttf -Recurse

# Creates new Shell.Application COM object
$shellObject = New-Object -ComObject Shell.Application

# Sets its namespace to the CLSID hex '0X14' (Windows font folder and registration trigger)
$windowsFontFolderTrigger = $shellObject.Namespace(0x14)

# Move and register font files.
foreach($fontFile in $fontFiles) {
  # test if font is already installed 
  if ( !(Test-Path "${env:windir}\Fonts\$($fontFile.Name)") ) {
    # install font
    $windowsFontFolderTrigger.CopyHere($fontFile.fullname)
  }
}

