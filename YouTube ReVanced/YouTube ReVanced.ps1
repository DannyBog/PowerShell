#References:
#https://www.reddit.com/r/revancedapp/comments/11b6zy4/how_to_use_revanced_cli_for_nonroot_users_step_by/
#https://revanced.app/patches

function Get-LatestVersion {
    param (
        [Parameter (Mandatory = $true)] [string]$Repo,
        [string]$Tag,
        [string]$File,
        [switch]$Stable
    )

    if ($Stable) {
        $releases = "https://api.github.com/repos/$Repo/releases/latest"
    } else {
        $releases = "https://api.github.com/repos/$Repo/releases"
    }

    $jsons = Invoke-WebRequest -Uri $releases | ConvertFrom-Json

    if ($Tag) {
        if ($File) {
            $assets = $jsons | Where-Object {$_.name -cmatch $Tag} | ForEach-Object {$_.assets | Where-Object {$_.name -like $File}}
        } else {
            $assets = $jsons | Where-Object {$_.name -cmatch $Tag}
        }
    } elseif ($File) {
        $assets = $jsons[0].assets | Where-Object {$_.name -like $File}
    } else {
        $assets = $jsons[0].assets
    }

    foreach ($asset in $assets) {
        $path = (Split-Path $PSCommandPath) + "\" + $($asset.name)
        Invoke-WebRequest -Uri $($asset.browser_download_url) -OutFile $path
    }

    return $path
}

if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Start-Process "powershell.exe" -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
}

if (-not(Test-Path -Path "C:\Program Files\Zulu\zulu-17\bin\java.exe")) {
    $url = "https://cdn.azul.com/zulu/bin/?C=M;O=D"
    $java = (Split-Path $PSCommandPath) + "\" + "java.msi"
    $download = (Invoke-WebRequest -Uri $url | Select-Object -ExpandProperty Links | Where-Object {$_.href -like "*zulu17*win_x64.msi"} | Select-Object -First 1).href
    $url = (Split-Path $url -Parent).Replace("\", "/")
    Invoke-WebRequest -Uri "$url/$download" -OutFile $java
    Start-Process "msiexec.exe" -ArgumentList "/i `"$java`" /passive" -Wait
    Remove-Item -Path $java

    $Env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine")
}

$microg = Get-LatestVersion -Repo "revanced/gmscore" -File "app.revanced*.apk" -Stable
$cli = Get-LatestVersion -Repo "revanced/revanced-cli" -File "revanced-cli*.jar" -Stable
$patches = Get-LatestVersion -Repo "revanced/revanced-patches" -File "revanced-patches*.jar" -Stable
$integrations = Get-LatestVersion -Repo "revanced/revanced-integrations" -File "revanced-integrations*.apk" -Stable

# Not a reliable way to get the most up-to-date version anymore :(
#$url = "https://raw.githubusercontent.com/ReVanced/revanced-patches/main/patches.json"
#$packages = (Invoke-WebRequest -Uri $url | ConvertFrom-Json).compatiblePackages
#$version = $packages | Where-Object {$_.name -eq "com.google.android.youtube"} | Select-Object -ExpandProperty versions -Unique | Sort-Object | Select-Object -Last 1
$version = (java.exe -jar "$cli" list-versions "$patches" -f "com.google.android.youtube" | Select-Object -Skip 2 | Sort-Object | Select-Object -Last 1).Split(" ")[0].Trim()

$url = "https://apkpure.com/youtube/com.google.android.youtube/versions"
$youtube = (Split-Path $PSCommandPath) + "\" + "youtube.apk"
$headers = @{"User-Agent" = 'Mozilla/5.0 (Windows NT; Windows NT 10.0; en-US) AppleWebKit/534.6 (KHTML, like Gecko) Chrome/7.0.500.0 Safari/534.6'}
$download = (Invoke-WebRequest -Uri $url -Headers $headers -UseBasicParsing | Select-Object -ExpandProperty Links | Where-Object {$_."data-dt-version" -eq $version}).href
$download = (Invoke-WebRequest -Uri $download -Headers $headers -UseBasicParsing | Select-Object -ExpandProperty Links | Where-Object {$_.class -eq "btn download-start-btn"}).href
Invoke-WebRequest -Uri $download -Headers $headers -OutFile $youtube

$cd = Split-Path $PSCommandPath -Parent
$youtube = $cd + "\youtube.apk"
$revanced = $cd + "\YouTube ReVanced ($version).apk"
Start-Process "java.exe" -ArgumentList "-jar `"$cli`" patch `"$youtube`" -m `"$integrations`" -b `"$patches`" -o `"$revanced`" -p" -Wait
Remove-Item -Path $youtube, $cli, $patches, $integrations, "$cd\YouTube ReVanced ($version)-options.json", "$cd\YouTube ReVanced ($version).keystore"
Rename-Item -Path $microg -NewName "microg.apk"