#Loading Variables to Profile
. $PSScriptRoot\local_variables.ps1

$LocalVariableProfile = "$PSScriptRoot\local_variables.ps1" 

function Print-Variables {
	# This function prints the loaded variables automatically, whenever needed.			
	Write-Host "["
	(Get-Content $LocalVariableProfile | Select-String -Pattern '^\$.\w*[^ =]').Matches.Value | ForEach-Object {
		Write-Host "`n{`n`tVariable: `t$_ `n`tPath: `t`t$(Invoke-Expression $_)`n},"	
	}	
	Write-Host "]"
}

function Copy-ProfileObjects {
    [alias("cpprof")]
    param (
        [switch]$push
    )
    Set-Content -Path $CopyProfilePath -Value (Get-Content $Profile)

    if ($push) {
        $oldDirectory = (Get-Location).path
        Set-Location (Split-Path $CopyProfilePath)
        
        git add .
        Write-Host "Write commit message for pushing to origin" -ForegroundColor Yellow
        $commitMessage = Read-Host
        git commit -m ($commitMessage)
        git push origin main
        
        Set-Location $oldDirectory
    }


}

function Edit-EnvironmentalVariable() {
    [alias("edit_env")]
        param()
	cmd /c "sysdm.cpl"
}

function Open-Zoom {
    [alias("zoom")]
    param()
    Start-Process $ZoomPath
}

function Open-Wolfram {
    [alias("wolfram")]
    param (
        [string]$v,
		[string]$file,
		[switch]$listversions
    )
	
	if($listversions){
		Write-Host "Printing available versions..."
		Write-Host "Mathematica" -ForegroundColor Blue
		Get-ChildItem "C:\Program Files\Wolfram Research\Mathematica"
		Write-Host "Wolfram" -ForegroundColor Blue
		Get-ChildItem "C:\Program Files\Wolfram Research\Wolfram"
		break
	}
	
	
	
    $versionTuple = $v.Split(".")
    
    if (
        ($versionTuple[0] -lt 14) -or
        ($versionTuple[0] -eq 14 -and $versionTuple[1] -eq 0)
    ) {
        $path = (Join-Path -Path "C:\Program Files\Wolfram Research\Mathematica" -ChildPath "$v\Mathematica.exe")
    } else
    {
        $path = (Join-Path -Path "C:\Program Files\Wolfram Research\Wolfram" -ChildPath "$v\WolframNB.exe")
    }
	
	if ($file){
    	Write-Host "Opening: $file with $open" -ForegroundColor Blue
    	Start-Process "$path $file" 
	} else {
    	Write-Host "Starting: $path" -ForegroundColor Blue
    	Start-Process $path
	}

}

function Save-ClipboardImage {
    [alias("ss")]
    param(
        [string]$dir = (Get-Location).Path,
        [bool]$setPathToClipboard = $true
    )

    $origLoc = (Get-Location).Path

    # Function to get image from clipboard
    function Get-ClipboardImage {
        Add-Type -AssemblyName System.Windows.Forms
        if ([System.Windows.Forms.Clipboard]::ContainsImage()) {
            $img = [System.Windows.Forms.Clipboard]::GetImage()
            return $img
        }
        else {
            Write-Error "No image found in clipboard."
            return $null
        }
    }

    # Prompt user for a file name
    $FileName = Read-Host "Enter the file name (without extension)"
    $FileName = (Split-Path $dir -leaf) + (Get-Date -Format "_MM_dd_yyyy_") + $FileName

    # Write-Host $FileName

    # Check if the file name is empty
    if (-not [string]::IsNullOrEmpty($FileName)) {
        Set-Location $dir        
        $FilePath = Join-Path -Path (Get-Location).Path -ChildPath "$FileName.png"
        # Write-Host $FilePath 
        (Get-ClipboardImage).Save($FilePath, [System.Drawing.Imaging.ImageFormat]::Png)
        Write-Output "Image saved to $FilePath"

    }
    else {
        Write-Error "File name cannot be empty."
    }

    if ($setPathToClipboard) {
        Set-Clipboard $FilePath
        Write-Host "`nPath copied, ready to paste as path`n" -ForegroundColor Green
    } else {
        Write-Host "You can always copy the path by selecting the path above."
    }

    Set-Location $origLoc

}


function Open-EdgeFileExplorer{ 
    [alias("browse")]
    param(
        [string] $path
    )

    if (-not $path) {
        $path = (Resolve-Path .)
        Write-Host "`nOpening Edge file browser in current directory: $path.`n" -ForegroundColor Blue
    } else {
        $path = (Resolve-Path $path)
        Write-Host "`nOpening Edge file browser path: $path.`n" -ForegroundColor Blue
    }
    Start-Process "msedge.exe" -ArgumentList $path
    
}

 
function Set-PowershellTheme {
    [alias("setPowTheme")]
    param (
        [string]$theme
    )
    $json = Get-Content $LocalPowershellSettings | ConvertFrom-Json
    $json.profiles.list[0].colorScheme = $theme
    $json | ConvertTo-Json -depth 100 | Set-Content $LocalPowershellSettings
}

function Set-Workspace {
    [alias("workspace")]
    param(
        [string]$val,
        [switch]$officeMode
    ) 
    $registryPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize"
    $isDarkModeEnabled = ((Get-ItemPropertyValue -Path $registryPath -Name AppsUseLightTheme) -eq 0) -and ((Get-ItemPropertyValue -Path $registryPath -Name SystemUsesLightTheme) -eq 0)
	
    if ($val -in @('w', 'wolfram','light-mode') ){
        $pict = Join-Path -Path $WallPaperImagesDirectory -ChildPath "working.jpg"

        if ($isDarkModeEnabled){
            Write-Host "`nSetting light mode"
            Start-Process ms-settings:colors
            Write-Host "`nSetting light mode on pow"
            Set-PowershellTheme 'One Half Light (Copy)'
            
			if($officeMode){
				Write-Host "`nMoving... "
            	Set-Location $WorkDirectory
				Write-Host "`nOpening Mitel..."
            	Mitel.exe
            	Write-Host "`nOpening Zoom...`n"
            	Open-Zoom
			}


        } else {
            Write-Host "Computer is already set on light mode. No changes will apply"
        }
        
    } elseif ($val -in @('o', 'out','dark','darkmode','dark-mode')) {
        $wallpapers = Get-ChildItem (Join-Path -Path $WallPaperImagesDirectory -ChildPath "real_wallpapers")
        $rand = Get-Random -Maximum ($wallpapers).count
        $pict = $wallpapers[$rand].FullName
        
        if (-not ($isDarkModeEnabled)){
            Write-Host "Setting dark mode"
            Start-Process ms-settings:colors
            Write-Host "Setting dark mode on pow"
            Set-PowershellTheme 'One Half Dark'
            Set-Location $Env:HOMEPATH

            if ($kill) {
                Get-Process | Where-Object {$_.ProcessName -match 'mitel'} | Stop-Process
                Get-Process | Where-Object {$_.ProcessName -match 'zoom'} | Stop-Process
                openvpn --command disconect_all
            }

        } else {
            Write-Host "Computer is already set on dark mode. Changing desktop wallpaper accordingly." -ForegroundColor Blue
        }

    } else {
        Write-Host "`nWrong Flag. Run Again`n" -ForegroundColor Red
    }
	# From https://www.powershellgallery.com/Packages/VirtualDesktop/1.5.7
    Set-AllDesktopWallpapers $pict
}

function Show-Screenshot {
    [alias("showss")]
    param(
        [int]$int
    )
    
    $screenshots = (Get-ChildItem $Env:OneDrive | Select-Object -Index 3).FullName + "\Capturas de pantalla\"

    Get-ChildItem $screenshots | ForEach-Object {
        $newname = ($_.Name -replace ' ','_')
        Move-Item $_.FullName -Destination ( Join-Path $NewScreenshotsDirectory $newname)
    }

    $files = Get-ChildItem $NewScreenshotsDirectory | Select-Object -Last $int

    foreach ($file in $files) { Start-Process "msedge.exe" -ArgumentList $file.FullName } 

}

function Copy-Path {
    [alias("cpa")] 
    param(
        [string]$path
    )
    Set-Clipboard (Resolve-Path $path).ToString() 
} 

function Get-YouTube {
    [alias("yt")]
    param(
        [string]$flag
    )

    if ($flag.Contains("https://") -or (-not $flag)) {
        if(-not $flag){
            $flag = Get-Clipboard
        }
        if ($flag -match 'v=([^&]+)') {
            $videoId = $matches[1]
            $url = -join ("microsoft-edge:https://www.youtube.com/embed/", $videoId)
            Start-Process $url  
        } else {
            Write-Error "No valid link. Closing."
            break
        }

    } elseif ( $flag -eq 'h') {
        $url = 'microsoft-edge:https://www.youtube.com/feed/history'
        Start-Process $url 
        break

    } elseif ($flag -eq 'wl'){
        $url = 'microsoft-edge:https://www.youtube.com/playlist?list=WL'
        Start-Process $url 
        break
    }

}

function Start-PowershellAdminMode{
    [alias("powad")]
    param()
    Start-Process powershell -Verb RunAs
}

function Open-MicrosoftOffice {
	[alias("msof")]
    param(
		[string]$val
	)
    $officePath = 'C:\Program Files\Microsoft Office\root\Office16' 
	
    switch($val){
		'word'	{& Join-Path $officePath -ChildPath 'WINWORD.exe'}
		'excel'	{& Join-Path $officePath -ChildPath 'EXCEL.exe'}
		'point'	{& Join-Path $officePath -ChildPath 'POWERPNT.exe'}
		'onote'	{& Join-Path $officePath -ChildPath 'ONENOTE.exe'}
	}
	
}

function Open-GoogleSearch {
    [alias("google")]
    param()
 
    if(-not $args){
        $search = (Get-Clipboard).Replace(" ","+")
    } else {
        $search = $args -join "+"
    }

    $url = "https://www.google.com/search?q=$search&udm=14"

    Start-Process msedge -ArgumentList $url

}

function Write-ToLog {
	[alias("log")]
	param(
		[string] $category, 
		[string] $log,
		[int] $check
	)

	if ($check) {
		Write-Host "`nFrom latest to earliest" -ForegroundColor Blue
		"SELECT * FROM log ORDER BY timestamp DESC LIMIT $check" | sqlite3 $LogDatabase
		Write-Host "`n"
		break
	}
	
	$validCategoryList = @('w','o','s', 'h')
	if (-not ( ($category) -and ($category -in $validCategoryList) )) {
		:validation do {
		$category = Read-Host "Is this (w)ork, (o)ut, (h)ome or (s)tudying?"
		if (-not ($category -in $validCategoryList)){
			Write-Host "This is not a valid option!" -ForegroundColor Red				
			Write-Host "Validating again`n" -ForegroundColor Yellow
		} else {
			break validation	
		}	
		} while ($true)
	}

	if(-not $log){
		$log = Read-Host "Enter a brief description of what you want to log"
	}
	
	#Making sure that $log does not contains "'" (escape characters by default on sqlite3)
	$log = $log.Replace("'","''")


	$command = "INSERT INTO log VALUES (strftime('%Y-%m-%d %H:%M:%S','now','localtime'),'$category','$log')"
	
	#Warning! This is insecure since it's prone to SQL inkections. 
	$command | sqlite3.exe $LogDatabase  
}

function Set-HideItem {
	[alias("hide")]	
	param(
		[string] $path
	)
		
	Write-Host "Hidding file: $path"
	$item = Get-Item $path
	$item.Attributes = $item.Attributes -bor "Hidden"

}

function Set-LocationModified {
	[alias("mcd")] param(
		[string] $path
	)
		
	if (-not $path){
		$path = Get-Clipboard
	}
	
	$path = Resolve-Path $path
	
	if(-not (Test-Path $path)){
		$opt = $null 
		Write-Host "'$path' is not a valid path."
		Write-Host @"
Select any of the following options:
	{
		c: create directory '$path'
		f: create file '$path'
		q: quit
	}
"@
		$opt = Read-Host "[c/f/[q]]"
		switch ($opt) {
			'c' {New-Item $path -ItemType "File"}
			'f' {New-Item $path -ItemType "Directory"}
			'q' {return}
			Default {Write-Host "Unknown option. Breaking"; return}
		}	
		return
	}

	if ($path.PSIsContainer) {
		# If it's a directory, move to directory.
		Set-Location ($path)
	} else {
		# If it's a file, move to directory where it is located, move to directory where it is located..
		Set-Location (Split-Path $path)
	}
	
}
