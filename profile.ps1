# Loading Variables to Profile
. $PSScriptRoot\local_variables.ps1

# Getting rid of Microsoft Copyright
Clear-Host

# Loading variables, perhaps there is a better and more elegant way?
$LocalVariableProfile = "$PSScriptRoot\local_variables.ps1" 



# This function prints the loaded variables automatically, whenever needed.			
function Print-Variables {
	[alias("pvar")]
	param(

	)
	Write-Host "["
	(Get-Content $LocalVariableProfile | Select-String -Pattern '^\$.\w*[^ =]').Matches.Value | ForEach-Object {
		Write-Host "`n{`n`tVariable: `t$_ `n`tPath: `t`t$(Invoke-Expression $_)`n},"	
	}	
	Write-Host "]"
}


# Function that copies profile to ~\projects\profile
# if -push, it asks for a commit message to push to git
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


# function to manually edit environmental variables
function Edit-EnvironmentalVariable() {
    [alias("edit_env")]
        param()
	cmd /c "sysdm.cpl"
}


# opens zooms duh
function Open-Zoom {
    [alias("zoom")]
    param(
	
	)
    Start-Process $ZoomPath
}


# opens wolfram mathematica software, specifying version
function Open-Wolfram {
    [alias("wolfram")]
    param (
        [string]$v,
		[switch]$listversions
    )

	$GeneralDirectory = Join-Path -Path $env:PROGRAMFILES -ChildPath "Wolfram Research"
	$MathematicaDir = Join-Path $GeneralDirectory -ChildPath "Mathematica"
	$WolframDir = Join-Path $GeneralDirectory -ChildPath "Wolfram"
	
	if($listversions){
		Write-Host "Printing available versions..."
		Write-Host "Mathematica" -ForegroundColor Blue
		Get-ChildItem $MathematicaDir
		Write-Host "Wolfram" -ForegroundColor Blue
		Get-ChildItem $WolframDir
		return
	}
	
	
    $versionTuple = $v.Split(".")
    if (
        ($versionTuple[0] -lt 14) -or
        ($versionTuple[0] -eq 14 -and $versionTuple[1] -eq 0)
    ) {
        $path = (Join-Path -Path $MathematicaDir -ChildPath "$v\Mathematica.exe")
    } else
    {
        $path = (Join-Path -Path $WolframDir -ChildPath "$v\WolframNB.exe")
    }
	
	Write-Host "Starting: $path" -ForegroundColor Blue
	Start-Process $path


}


# copy a image, from any source, and paste it on a specified path
function Save-ClipboardImage {
    [alias("ss")]
    param(
        [string]$dir = (Get-Location).Path,
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
            return
        }
    }

    # Prompt user for a file name
    $FileName = Read-Host "Enter the file name (without extension)"
    $FileName = (Split-Path $dir -leaf) + (Get-Date -Format "_MM_dd_yyyy_") + $FileName

    # Check if the file name is empty
	# change this chatgpt validation wtf
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

    Set-Clipboard $FilePath
    Write-Host "`nPath copied, ready to paste as path`n" -ForegroundColor Green
    Set-Location $origLoc

}


# open file explorer on browser, ideal for low-resources pcs
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


 # set powershell theme
function Set-PowershellTheme {
    [alias("setPowTheme")]
    param (
        [string]$theme
    )
    $json = Get-Content $LocalPowershellSettings | ConvertFrom-Json
    $json.profiles.list[0].colorScheme = $theme
    $json | ConvertTo-Json -depth 100 | Set-Content $LocalPowershellSettings
}


# change workspace accordingly
# switching from dark mode to light mode and viceversa
function Set-Workspace {
    [alias("workspace")]
    param(
		[Parameter(Mandatory=$true)]
        [string]$val,
        [switch]$officeMode,
		[switch]$kill

    ) 
	
	$registryPath = Registry::HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize\
	$appsInDarkMode = Get-ItemPropertyValue $registryPath AppsUseLightTheme
	$systemInDarkMode = Get-ItemPropertyValue $registryPath SystemUsesLightTheme
    $isDarkModeEnabled = $appsInDarkMode -and $systemInDarkMode

	
    if ($val -in @('w', 'wolfram','light-mode') ){
        $pict = Join-Path -Path $WallPaperImagesDirectory -ChildPath "working.jpg"
	
		if($officeMode){
			Write-Host "Office Mode selected. Openning Zoom and Mitel."
           	Set-Location $WorkDirectory
           	Mitel.exe
           	Open-Zoom
			Write-Host "Current directory changed to $WorkDirectory."
		}

        if ($isDarkModeEnabled){
    		# as per the reddit post:
			# https://www.reddit.com/r/Windows11/comments/11h4p5c/programatically_change_to_dark_mode_in_windows_11/
			# seems impossible to do it from powershell
            Start-Process ms-settings:colors
            Write-Host "Setting light mode on Powershell."
            Set-PowershellTheme 'One Half Light (Copy)'
        } else {
            Write-Host "Computer is already set on light mode. No changes will be applied." -ForegroundColor Yellow
        }
        
		
    } elseif ($val -in @('o', 'out','dark','darkmode','dark-mode')) {

		# select random wallpaper from wallpaper directory
        $wallpapers = Get-ChildItem (Join-Path -Path $WallPaperImagesDirectory -ChildPath "real_wallpapers")
        $rand = Get-Random -Maximum ($wallpapers).count
        $pict = $wallpapers[$rand].FullName

        if ($kill) {
			Write-Warning "You may lose data. Proceed cautiously."
            Get-Process | Where-Object {$_.ProcessName -match 'mitel'} | Stop-Process
            Get-Process | Where-Object {$_.ProcessName -match 'zoom'} | Stop-Process
			$openVPNPath = Join-Path $env:PROGRAMFILES -ChildPath "OpenVPN\bin\openvpn-gui.exe"
			if (Test-Path $openVPNPath) {
				& $openVPNPath --command disconect_all
			}
 
        
        if (-not ($isDarkModeEnabled)){
			# as per the reddit post:
			# https://www.reddit.com/r/Windows11/comments/11h4p5c/programatically_change_to_dark_mode_in_windows_11/
			# seems impossible to do it from powershell
            Start-Process ms-settings:colors
            Write-Host "Setting dark mode on Powershell."
            Set-PowershellTheme 'One Half Dark'
            Set-Location $Env:HOMEPATH
           }
        } else {
            Write-Host "Computer is already set on dark mode. Changing desktop wallpaper accordingly." -ForegroundColor Blue
        }

    } else {
        Write-Error -Message "Invalid Argument." -Category InvalidArgument
    }
	# TODO: implement try - catch for profiles not having imported VirtualDesktop
	# or implement from scratch
	# From https://www.powershellgallery.com/Packages/VirtualDesktop/1.5.7
    Set-AllDesktopWallpapers $pict
}


# function that shows screenshots, same idea as file explorer on browser,
# made for low-resource computers.
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


# function that copies the current path to clipboard
function Copy-Path {
    [alias("cpa")] 
    param(
        [string]$path
    )
    Set-Clipboard (Resolve-Path $path).ToString() 
} 


# function that opens yt videos as /embed/
# probably deprecated
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


# opens powershell in admin mode 
function Start-PowershellAdminMode{
    [alias("powad")]
    param()
    Start-Process powershell -Verb RunAs
}


# opens microsoft office software from commandline
function Open-MicrosoftOffice {
	[alias("msof")]
    param(
		[string]$val
	)

	$officePath = Join-Path $env:PROGRAMFILES -ChildPath "Microsoft Office\root\Office16"
    switch($val){
		'word'	{Start-Process (Join-Path $officePath -ChildPath 'WINWORD.exe') }
		'excel'	{Start-Process (Join-Path $officePath -ChildPath 'EXCEL.exe') }
		'point'	{Start-Process (Join-Path $officePath -ChildPath 'POWERPNT.exe') }
		'onote'	{Start-Process (Join-Path $officePath -ChildPath 'ONENOTE.exe') }
	}
	
}


# searches on the udm=14 interface of chrome
function Open-GoogleSearch {
    [alias("google")]
    param(

	)
 
    if(-not $args){
        $search = (Get-Clipboard).Replace(" ","+")
    } else {
        $search = $args -join "+"
    }

    $url = "https://www.google.com/search?q=$search&udm=14"
    Start-Process msedge -ArgumentList $url

}


# hides folders / files from a simple 'ls'
# https://stackoverflow.com/a/67226308/29272030
function Set-HideItem {
	[alias("hide")]	
	param(
		[string] $path
	)
		
	Write-Host "Hidding file: $path"
	try {
		$item = Get-Item $path -ErrorAction Stop
		$item.Attributes = $item.Attributes -bor "Hidden"
	} catch {
		Write-Host "'$path' does not exist."
	}

}


# function that cd's to a file directory
# if file does not exist, creates it
# 	choose between directory or file
# TODO: check if still works
function Set-LocationModified {
	[alias("mcd")] 

	param(
		[string] $path
	)
	
	if(-not (Test-Path $path)){ # if file doesn't exist
		$opt = $null 
		Write-Host "'$path' is not a valid path."
		Write-Host @"
Select any of the following options:
	{
		d: create directory '$path'
		f: create file '$path'
		q: quit
	}
"@
		$opt = Read-Host "[d/f/[q]]"
		switch ($opt) {
			'd' {New-Item $path -ItemType "Directory"}
			'f' {New-Item $path -ItemType "File"}
			'q' {return}
			Default {Write-Host "Unknown option. Breaking" -ForegroundColor Red; break}
		}	
	}

	$path = Get-Item (Resolve-Path $path)

	if ($path.PSIsContainer) {
		# If it's a directory, move to directory.
		Set-Location ($path)
	} else {
		# If it's a file, move to directory where it is located, move to directory where it is located..
		Set-Location (Split-Path $path)
	}
	
}


# print time in mini-hash time zones
function Get-Time {
	[alias("time")]
	param(

	)

	$timezones = [ordered]@{
		"chicago"	=	-1
		"lima"		= 	0
		"madrid"	=	6
	}
	
	foreach ( $timezone in $timezones.GetEnumerator() ) {
		$time = (Get-Date).AddHours($timezone.Value).ToShortTimeString()
		$city = $timezone.Name.ToUpper()
		Write-Host "Current time in $city : $time"
	}

}


# small todo management
function New-Todo{
	[alias("todo")]
	param(
		[string] $name,
		[switch] $move
	)
	
	if ($move){
		# first search for TODOs on current directory
		$todos = Get-ChildItem . -Filter 'todo*'
		# move them to TODO's dir
		foreach ($todo in $todos) {
			Move-Item -Path $todo.FullName -Destination $TODODirectory
		}
		Write-Host "All TODO files have been transferred to '$TODODirectory'."
		return
	}

	if (-not $name) {
		$name = (Get-Date).ToString("yyyy-MM-dd")
	}
	$name = "todo-" + $name

	New-Item -Name $name -ItemType "File"
	Write-Host "New todo file: $name"
}


# custom prompt for powershell
function prompt {
	$path = (Get-Location).Path.Replace("$env:USERPROFILE", "~") 
	$user = $env:USERNAME
	$computer = ($env:COMPUTERNAME).ToLower()
	$ssh = "[$user@$computer]"
	$time = get-date -Format "HH:mm:ss"

	Write-Host "PS " -NoNewline -ForegroundColor Yellow
	Write-Host "($time) " -NoNewline -ForegroundColor Yellow
	Write-Host "${ssh}: " -NoNewline -ForegroundColor Green
	Write-Host "$path" -ForegroundColor Cyan
	return "> "

}
