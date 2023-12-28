$phoneName = "Apple iPhone"
$phonePath = "Internal Storage\DCIM"

# Ask for target directory and create if not exist
$destinationPath = read-host -Prompt "Zielverzeichnis"
if(-not (Test-Path -Path $destinationPath)){
    mkdir -Path $destinationPath
}

# Just a shell object
# See: https://blog.daiyanyingyu.uk/2018/03/20/powershell-mtp/
function Get-ShellProxy
{
    if( -not $global:ShellProxy)
    {
        $global:ShellProxy = new-object -com Shell.Application
    }
    $global:ShellProxy
}

# Get the phone
function Get-Phone
{
    $shell = Get-ShellProxy
    # 17 (0x11) = ssfDRIVES from the ShellSpecialFolderConstants (https://msdn.microsoft.com/en-us/library/windows/desktop/bb774096(v=vs.85).aspx)
    # => "My Computer" — the virtual folder that contains everything on the local computer: storage devices, printers, and Control Panel.
    # This folder can also contain mapped network drives.
    $shellItem = $shell.NameSpace(17).self
    $phone = $shellItem.GetFolder.items() | Where-Object { $_.name -eq $global:phoneName }
    return $phone
}

# Get the folder content
function Get-Folder
{
    param($parent,[string]$path = "")
    $path = Join-Path $global:phonePath $path 
    $pathParts = @( $path.Split([system.io.path]::DirectorySeparatorChar) )
    $current = $parent

    # Just hangle from tree to tree unit we have nothing left in the path array
    foreach ($pathPart in $pathParts)
    {
        if ($pathPart)
        {
            $current = $current.GetFolder.items() | Where-Object { $_.Name -eq $pathPart }
        }
    }
    return $current.GetFolder.items()
}

# Get the starting points
$phone = Get-Phone
$folders = Get-Folder -parent $phone

# Now do for each IMG folder
foreach($folder in $folders){

    # Status Message
    write-host ("Folder: " + $folder.name) -ForegroundColor Yellow

    # Create Folder if not exist
    $destinationSubPath = Join-Path $global:destinationPath $folder.name
    if(-not (Test-Path -Path $destinationSubPath)){
        mkdir -Path $destinationSubPath
    }

    # Link destionation to MTP
    $shell = Get-ShellProxy
    $destinationMTPlink = $shell.Namespace($destinationSubPath).self

    # Retrive files from IMG Folder
    $files = Get-Folder -parent $phone -path $folder.name

    # Process each file
    foreach($file in $files){

        # Skip if subfolder (should never occur)
        if($file.IsFolder -eq $true){
            continue
        }

        # Skip AAE files (may cause troubble)
        if($file.name -like "*.AAE"){
            continue
        }

        # Check if file exists
        $destinationFilePath = Join-Path $destinationSubPath $file.name
        if(Test-Path -Path $destinationFilePath){
            write-host ("Does exist: " + $file.name + " ") -ForegroundColor Cyan
            continue
        }

        # Status Message
        write-host ("Move: " + $file.name + " ") -ForegroundColor White -NoNewline

        # Move IMG
        $destinationMTPlink.GetFolder.MoveHere($file)

        # Check result
        if(Test-Path -Path $destinationFilePath){
            write-host "OK" -ForegroundColor Green
        }
        else {
            write-host "ERROR" -ForegroundColor Red
        }
    }
}
