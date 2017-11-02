<#
.SYNOPSIS
	Reduce SharePoint Framework SPPKG file
.DESCRIPTION
    Display 

	Comments and suggestions always welcome!  spjeff@spjeff.com or @spjeff
.NOTES
	File Namespace	: Reduce-Sppkg.ps1
	Author			: Jeff Jones - @spjeff
	Version			: 0.10
	Last Modified	: 11-02-2017
.LINK
	Source Code
	http://www.github.com/spjeff/o365/reduce-sppkg
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $False, ValueFromPipeline = $false, HelpMessage = 'Use -p to provide input package.')]
    [Alias("p")]
    [string]$package
)

# params
$tempFolder = $env:TEMP + "\reduce-sppkg"
$global:keep = @()

# unzip function
Add-Type -AssemblyName System.IO.Compression.FileSystem
function Unzip {
    param([string]$zipfile, [string]$outpath)
    Write-Host "Unzip $zipfile to $outpath" -Fore Yellow
    Remove-Item $outpath -Recurse -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
    mkdir $outpath -ErrorAction SilentlyContinue | Out-Null
    try {
        [System.IO.Compression.ZipFile]::ExtractToDirectory($zipfile, $outpath) 
    }
    catch {}
}

function Zip {
    param([string]$folderInclude, [string]$outZip)
    [System.IO.Compression.CompressionLevel]$compression = "Optimal"
    $ziparchive = [System.IO.Compression.ZipFile]::Open( $outZip, "Update" )
    Write-Host "Zip $folderInclude to $outZip" -Fore Yellow

    # loop all child files
    $realtiveTempFolder = (Resolve-Path $tempFolder -Relative).TrimStart(".\")
    foreach ($file in (Get-ChildItem $folderInclude -Recurse)) {
        # skip directories
        if ($file.GetType().ToString() -ne "System.IO.DirectoryInfo") {
            # relative path
            $relpath = ""
            if ($file.FullName) {
                $relpath = (Resolve-Path $file.FullName -Relative)
            }
            if (!$relpath) {
                $relpath = $file.Name
            }
            else {
                $relpath = $relpath.Replace($realtiveTempFolder, "")
                $relpath = $relpath.TrimStart(".\").TrimStart("\\")
            }

            # display
            Write-Host $relpath
            Write-Host $file.FullName

            # add file
            [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($ziparchive, $file.FullName, $relpath, $compression) | Out-Null
        }
    }
    $ziparchive.Dispose()
}

function DisplayFeatures($tempFolder) {
    # Gather feature XML
    $features = Get-ChildItem "$tempFolder\feat*.xml"
    $features = $features |? {$_.Name -notlike "*config*"}

    # Parse Description from XML
    $coll = @()
    foreach ($f in $features) {
        [xml]$xml = Get-Content $f.FullName
        $title = $xml.Feature.Title
        $id = $xml.Feature.Id
        $obj = New-Object -TypeName PSObject -Prop (@{"Title" = $title; "Id" = $id})
        $coll += $obj
    }

    $coll | sort Title | ft -a
    Write-Host "$($features.length) features found" -Fore Green
}

function SelectKeepers() {
    Write-Host "Type in GUID numbers to keep.  Press ENTER for a blank input to finalize."
    $id = "GO"
    while ($id -ne "") {
        $id = Read-Host
        if ($id -ne "") {
            $global:keep += $id
        }
    }
    Write-Host "Keep Feature IDs" -Fore Yellow
    $global:keep | ft -a
}

function RemoveFeatures() {
    # Remove Directory (if not Keep GUID)
    foreach ($d in (Get-ChildItem $tempFolder -Directory)) {
        $keep = $null
        $keep = $global:keep |? {$_ -eq $d.Name}
        if (!$keep -and $d.Name -ne "_rels") {
            Remove-Item $d.FullName -Recurse -Confirm:$false -Force
        }
    }

    # Feature Config XML
    foreach ($f in (Get-ChildItem "$tempFolder\feature*.config.xml")) {
        $keep = $null
        $featureId = $f.Name.Replace(".xml.config.xml", "").Replace("feature_", "")
        $keep = $global:keep |? {$_ -eq $featureId}
        if (!$keep) {
            Remove-Item $f.FullName -Confirm:$false -Force
        }
    }
    
    # Feature XML
    foreach ($f in (Get-ChildItem "$tempFolder\feature*.xml")) {
        $keep = $null
        $featureId = $f.Name.Replace(".xml", "").Replace("feature_", "")
        $keep = $global:keep |? {$_ -eq $featureId}
        if (!$keep) {
            Remove-Item $f.FullName -Confirm:$false -Force
        }
    }

    # Feature Rel XML
    foreach ($f in (Get-ChildItem "$tempFolder\_rels\feature*.xml.rels")) {
        $keep = $null
        $featureId = $f.Name.Replace(".xml.rels", "").Replace("feature_", "")
        $keep = $global:keep |? {$_ -eq $featureId}
        if (!$keep) {
            Remove-Item $f.FullName -Confirm:$false -Force
        }
    }

    # AppManifest.xml
    $filePath = "$tempFolder\_rels\AppManifest.xml.rels"
    [xml]$xml = Get-Content $filePath
    $rels = $xml.Relationships.Relationship
    foreach ($r in $rels) {
        $keep = $null
        $featureId = $r.Target.Replace("/feature_", "").Replace(".xml", "")
        $keep = $global:keep |? {$_ -eq $featureId}
        if (!$keep) {
            # Remove XML Node
            $xml.Relationships.RemoveChild($r) | Out-Null
        }
    }
    $xml.Save($filePath)
}

function Main() {
    # UnZIP
    Unzip $package $tempFolder

    # Display all features
    DisplayFeatures $tempFolder

    # Select keepers
    SelectKeepers

    # Remove excess Features
    RemoveFeatures

    # ZIP
    $timestamp = (get-date).tostring("yyyy-MM-dd-hh-mm-ss")
    $newFilename = $package.Replace(".sppkg", "$timestamp.sppkg")
    Zip $tempFolder $newFilename

    # Done
    Remove-Item $tempFolder -Recurse -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
    Write-Host "Created $newFilename" -Fore Yellow
    Write-Host "DONE" -Fore Green
}
Main