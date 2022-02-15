if($PSVersionTable.PSVersion.Major -lt 7)
{
    Write-Error -Message "This script supports PowerShell 7 or later." -ErrorAction Stop
}
Set-Variable -Name ProductionName -Value "InspectAppsUtils" -Option ReadOnly
Set-Variable -Name ProductionPath -Value (Join-Path -Path $PSScriptRoot -ChildPath $ProductionName) -Option ReadOnly
Set-Variable -Name FilesMapping -Value @{
} -Option ReadOnly
if((git tag --list --contains HEAD) -match "^v(?<version>[0-9]+\.[0-9]+\.[0-9])")
{
    $version = $Matches["version"]
}
else
{
    Write-Warning "The current Commit does not contain versioning tag."
    $version = "0.0.1"
}

Set-Variable -Name TextPath -Value (Join-Path -Path $PSScriptRoot -ChildPath "src/txt") -Option ReadOnly
Set-Variable -Name TemporaryDirectory -Value (Join-Path -Path (Get-Item -Path Env:\TEMP).Value -ChildPath $ProductionName) -Option ReadOnly

if(Test-Path -Path $ProductionPath){
    Write-Warning -Message "Old Production was removed."
    Remove-Item -Path $ProductionPath -Recurse
}
New-Item -ItemType Directory -Path $ProductionPath > $null
if(Test-Path -Path $TemporaryDirectory){
    Remove-Item -Path $TemporaryDirectory -Recurse
}
New-Item -ItemType Directory -Path $TemporaryDirectory > $null

$languages = Get-ChildItem -Path $TextPath -Directory
Get-ChildItem -Path ./src/*.cls | ForEach-Object -Process {
    $codePath = $_
    $translationPath = Join-Path -Path $TextPath -ChildPath ([System.IO.Path]::ChangeExtension($codePath.Name, "xml"))
    if(-not (Test-Path -Path $translationPath))
    {
        Write-Error -Message "$($translationPath) is not exists." -ErrorAction Stop
    }
    $tempFilePath = Join-Path -Path $TemporaryDirectory -ChildPath $codePath.Name
    Get-Content -Path $codePath | Where-Object -FilterScript { $_ -cnotmatch "^ *'@[A-Za-z]+" } | ForEach-Object -Process {$_ -replace '\$PackageVersion', $version} | Out-File -FilePath $tempFilePath
    Resolve-VBATranslationPlaceHolder -SourcePath $tempFilePath -TranslationPath (Join-Path -Path $TextPath -ChildPath ([System.IO.Path]::ChangeExtension($codePath.Name, "xml"))) -DestinationPath (Join-Path -Path $ProductionPath -ChildPath $codePath.Name)
    $languages | ForEach-Object -Process {
        $language = Split-Path -Path $_ -Leaf
        $translationPath = Join-Path -Path (Join-Path -Path $TextPath -ChildPath $language) -ChildPath ([System.IO.Path]::ChangeExtension($codePath.Name, "xml"))
        if(-not (Test-Path -Path $translationPath))
        {
            Write-Warning -Message "$($translationPath) is not exists."
            return
        }
        Resolve-VBATranslationPlaceHolder -SourcePath $tempFilePath -TranslationPath $translationPath -DestinationPath (Join-Path -Path (Join-Path -Path $ProductionPath -ChildPath $language) -ChildPath $codePath.Name)
    }
}
Compress-Archive -Path $ProductionPath -DestinationPath "$($ProductionPath).zip" -Force
Remove-Item -Path $ProductionPath -Recurse

Write-Output "New Production was created."