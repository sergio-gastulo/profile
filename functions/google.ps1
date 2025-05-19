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