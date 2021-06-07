New-Item -ItemType directory -Path C:\font

cd c:\font

wget "https://github.com/vegardjpc/font/blob/main/CrimsonText-Bold.otf?raw=true" -outfile "CrimsonText-Bold.otf"
wget "https://github.com/vegardjpc/font/blob/main/CrimsonText-BoldItalic.otf?raw=true" -OutFile "CrimsonText-BoldItalic.otf"
wget "https://github.com/vegardjpc/font/blob/main/CrimsonText-Italic.otf?raw=true" -OutFile "CrimsonText-Italic.otf"
wget "https://github.com/vegardjpc/font/blob/main/CrimsonText-Regular.otf?raw=true" -OutFile "CrimsonText-Regular.otf"
wget "https://github.com/vegardjpc/font/blob/main/CrimsonText-SemiBold.otf?raw=true" -OutFile "CrimsonText-SemiBold.otf"
wget "https://github.com/vegardjpc/font/blob/main/CrimsonText-SemiBoldItalic.otf?raw=true" -OutFile "CrimsonText-SemiBoldItalic.otf"
wget "https://github.com/vegardjpc/font/blob/main/Lexend-Bold.otf?raw=true" -OutFile "Lexend-Bold.otf"
wget "https://github.com/vegardjpc/font/blob/main/Lexend-ExtraBold.otf?raw=true" -OutFile "Lexend-ExtraBold.otf"
wget "https://github.com/vegardjpc/font/blob/main/Lexend-Light.otf?raw=true" -OutFile "Lexend-Light.otf"
wget "https://github.com/vegardjpc/font/blob/main/Lexend-Medium.otf?raw=true" -OutFile "Lexend-Medium.otf"
wget "https://github.com/vegardjpc/font/blob/main/Lexend-Regular.otf?raw=true" -OutFile "Lexend-Regular.otf"
wget "https://github.com/vegardjpc/font/blob/main/Lexend-SemiBold.otf?raw=true" -OutFile "Lexend-SemiBold.otf"
wget "https://github.com/vegardjpc/font/blob/main/Lexend-Thin.otf?raw=true" -OutFile "Lexend-Thin.otf"

$SourceDir   = "C:\font"
$Source      = "C:\font*"
$Destination = (New-Object -ComObject Shell.Application).Namespace(0x14)
$TempFolder  = "C:\Windows\Temp\Fonts"


New-Item -ItemType Directory -Force -Path $SourceDir

New-Item $TempFolder -Type Directory -Force | Out-Null

Get-ChildItem -Path $Source -Include '*.ttf','*.ttc','*.otf' -Recurse | ForEach {
    If (-not(Test-Path "C:\Windows\Fonts\$($_.Name)")) {

        $Font = "$TempFolder\$($_.Name)"
        
        # Copy font to local temporary folder
        Copy-Item $($_.FullName) -Destination $TempFolder
        
        # Install font
        $Destination.CopyHere($Font,0x10)

        # Delete temporary copy of font
        Remove-Item $Font -Force
    }
}