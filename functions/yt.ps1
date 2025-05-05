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
            $url = -join ("https://www.youtube.com/embed/", $videoId)
            Start-Process $url  
        } else {
            Write-Error "No valid link. Closing."
            break
        }

    } elseif ( $flag -eq 'h') {
        $url = 'https://www.youtube.com/feed/history'
        Start-Process $url 
        break

    } elseif ($flag -eq 'wl'){
        $url = 'https://www.youtube.com/playlist?list=WL'
        Start-Process $url 
        break
    }

}