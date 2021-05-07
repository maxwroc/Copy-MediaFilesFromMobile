[CmdletBinding()]
param()

$configFileName = "$($MyInvocation.MyCommand.Name.Substring(0, $MyInvocation.MyCommand.Name.Length - 4))Config.xml"

if(-not (Test-Path $configFileName)) {
    Write-Error "Config file missing ($($configFileName))"
    exit
}

Write-Verbose "Reading config"
try {
    $configFile = [xml] (Get-Content $configFileName)
    $config = $configFile.Config
}
catch {
    Write-Error "Failed to load the config file: $($_.Exception.Message)"
    Exit
}


# http://blogs.technet.com/b/heyscriptingguy/archive/2013/04/26/use-powershell-to-work-with-windows-explorer.aspx
$o = New-Object -com Shell.Application
# https://msdn.microsoft.com/en-us/library/windows/desktop/bb774096(v=vs.85).aspx
# ShellSpecialFolderConstants.ssfDRIVES == 0x11
$folder = $o.NameSpace(0x11)





Function GetSubFolder ($folder, $sourcePathChunks, $sourcePathDepth)
{
    if($sourcePathDepth -ge $sourcePathChunks.Count) {
        Write-Verbose "Reached the end of the path ($($sourcePathDepth))"
        Write-Host $sourcePathChunks
        return
    }

    $searchingFor = $sourcePathChunks[$sourcePathDepth]
    Write-Verbose "Looking for $($searchingFor)"

    foreach ($i in $folder) {
        Write-Verbose "    Checking if $($i.Name) is a folder: $($i.IsFolder)"
        if ($i.IsFolder -and $i.Name -eq $searchingFor) {
            
            $fld = $i.GetFolder()
            if($sourcePathDepth -eq $sourcePathChunks.Count-1) {
                Write-Verbose "Found"
                return $fld
            }
            else {
                return GetSubFolder $fld.Items() $sourcePathChunks ($sourcePathDepth+1)
            }
        }
    }

    Write-Host -NoNewline "This PC"
    for ($i=0; $i -lt $sourcePathChunks.Count; $i++) {
        Write-Host -NoNewline "\"
        if($sourcePathDepth -eq $i) {
            Write-Host -NoNewline -ForegroundColor Yellow $sourcePathChunks[$i]
        }
        else {
            Write-Host -NoNewline $sourcePathChunks[$i]
        }
    }
    Write-Host ""

    Write-Host -ForegroundColor Red "Error: Couldn't find the path specified"

    return
}

Function CopyFiles ($sourceFolder, $destination, $filters) {
    
    Write-Verbose "Fetching files from source (this can take a while for a large directories)"
    $items = $sourceFolder.Items()

    # create empty collection
    $filesToProcess = @()
    Write-Verbose "Iterating over $($items.Count) files"
    for ($i = 0; $i -lt $items.Count; $i++) {
        Write-Progress -Activity "Detecting how many items to copy" -Status "    Files: ($($filesToProcess.Count))" -PercentComplete ($i / $items.Count * 100)

        $isValid = 1
        foreach ($filter in $filters) {
            $isValid = IsValidFile $items.Item($i) $filter
            if (-not $isValid) {
                break
            }
        }

        if (-not $isValid) {
            continue
        }

        $filesToProcess += ,$items.Item($i)
    }

    Write-Host "Files to process: $($filesToProcess.Count)"

    $confirmation = Read-Host "Do you want to proceed and copy these files? [y/n]: "
    if ($confirmation -eq "y") {
        Write-Verbose "Getting temporary directory dir"
        $target = $o.NameSpace($destination)

        for ($i=0; $i -lt $filesToProcess.Count; $i++) {
            $fileName = $filesToProcess[$i].Name
            Write-Progress -Activity "Copying files" -Status "    $fileName" -PercentComplete ($i / $filesToProcess.Count * 100)
            if(Test-Path "$($destination)\$fileName") {
                Write-Host "    Skipping $fileName as it exists already."
            }
            else {
                Write-Verbose "Copying $fileName"
                $target.CopyHere($filesToProcess[$i], 0)
            }
        }
    }
}

Function IsValidFile ($file, $filter) {
    $result = 1
    switch ($filter.Type) {
        "NewerThan" { 
            $namePattern = $filter.ExtractDateFromName
            $date = [DateTime]::ParseExact($filter.Date, "dd/MM/yyyy", $null)
            $fileName = $file.Name
            
            if ($fileName.Length -lt [int]$namePattern.Substring.From + $namePattern.Substring.Length) {
                Write-Verbose "Too short name of the file: $fileName ($($fileName.Length) < $(([int]$namePattern.Substring.From) + $namePattern.Substring.Length))"
                return 0
            }
            
            if($namePattern.SubString) {
                $fileName = $fileName.Substring($namePattern.Substring.From, $namePattern.Substring.Length)
                Write-Verbose "Cutting file name: $($file.Name) -> $fileName"
            }
            Write-Verbose "Is valid file: $($fileName) [filter: $($filter.Type)][pattern: $($namePattern.Pattern)]"
            $parsedDate = [DateTime]::ParseExact($fileName, $namePattern.Pattern, $null)
            
            if ($parsedDate -lt $date) {
                Write-Verbose "File too old: $($fileName)"
                $result = 0
            }
         }
        Default {}
    }

    return $result
}

Function UpdateFilters ($filters, $destinationDir) {
    $updated = 0
    Write-Verbose "Updating filters"
    foreach ($filter in $filters) {
        switch ($filter.Type) {
            "NewerThan" {
                $latest = Get-ChildItem -Path $destinationDir | Sort-Object LastAccessTime -Descending | Select-Object -First 1
                $latestString = $latest.LastAccessTime.ToString("dd\/MM\/yyyy")
                if ($latestString -ne $filter.Date) {
                    $confirmation = Read-Host "The last file in the target dir is from $latestString do you want to update filter value? [y/n]"
                    if ($confirmation -eq "y") {
                        $filter.Date = $latestString
                        $updated = 1
                    }
                }
            }
        }
    }

    if ($updated) {
        Write-Verbose "Updating configuration: $PSScriptRoot\$configFileName"
        $configFile.Save("$PSScriptRoot\$configFileName")
    }
}

Write-Verbose "Iterating over available dives:"
foreach ($device in $folder.Items()) {
    Write-Verbose "    Checking config for device $($device.Name)"
    $deviceConfig = $config.Sources.Source | Where-Object {$_.Name -eq $device.Name}

    if($deviceConfig) {
        Write-Verbose "Loaded config for device $($device.Name)"
        Write-Verbose "    Path: $($deviceConfig.Path)"

        $sourceFolder = GetSubFolder $device.GetFolder().Items() $deviceConfig.Path.Split("\") 1
        if(!$sourceFolder) {
            Write-Error "Source folder not found: $($deviceConfig.Path)"
            exit
        }
        
        Write-Verbose "Check if destination folder exists: $($config.Destination.Temp)"
        if(-not (Test-Path $config.Destination.Temp)) {
            $ans = Read-Host "Destination folder does not exist do you want to create it? [y/n]"
            if($ans -eq "y") {
                New-Item -ItemType directory $config.Destination.Temp
            }
            
            if(-not (Test-Path $config.Destination.Temp)) {
                Write-Error "Destination folder not found ($($config.Destination.Temp))"
                exit
            }
        }

        CopyFiles $sourceFolder $config.Destination.Temp $deviceConfig.Filters.Filter

        UpdateFilters $deviceConfig.Filters.Filter $config.Destination.Temp
        break
    }
}
