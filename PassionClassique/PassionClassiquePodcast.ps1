<#
Author : Robert MILLI 

Goal :
	Query Passion Classique Podcast Web page and retrieve the latest podcast in a mp3 format, name the file in order to know who was the guest 
    and tag the mp3 file itself

    Normally, this script has to be scheduled on a daily basis

    This script has been created for my wife who loves to hear this podcast but it was not that convenient to hear on her mp3 player with a generic file name and no tags


Syntax :
	C:\Users\RMI1\Documents>Powershell -file PassionClassiquePodcast.ps1 

Sources :
    MP3 Tag                 - https://stackoverflow.com/questions/45298447/access-music-file-metadata-in-powershell
    Webpage pasring         - https://4sysops.com/archives/powershell-invoke-webrequest-parse-and-scrape-a-web-page/
    HTTP Download file      - https://4sysops.com/archives/use-powershell-to-download-a-file-with-http-https-and-ftp/
Changes tracking :
    2017/12/18 - 1.00 - Creation

#>

$urlPodcast ="https://www.radioclassique.fr/radio/emissions/passion-classique/"
$mp3Array = @()

# Read the web page in a variable
$WebResponse = Invoke-WebRequest $urlPodcast
# Search the podcast episode
Foreach($Link in $WebResponse.Links) {
    # Get the metadata of the podcast episode
    if($Link.href -match 'livePlayer') {
        $mp3SingleObject = $mp3Array | where {$_.Name -eq $Link.'data-mp3'}
        # If the object doesn't exist, let's create it
        if(!$mp3SingleObject) {
            $mp3Object = New-Object System.Object
            $mp3Object | Add-Member -type NoteProperty -name Name -Value $Link.'data-mp3'
            $mp3Object | Add-Member -type NoteProperty -name RemoteFileName ""
            $mp3Object | Add-Member -type NoteProperty -name Guest $Link.'data-guests'
            $mp3Object | Add-Member -type NoteProperty -name Date $Link.'data-date'
            $mp3Array += $mp3Object
        } else {
            # If the object exist, let's set the properties we want to keep
            $mp3SingleObject.Guests = $Link.'data-guests'
            $mp3SingleObject.Date = $Link.'data-date'
        }
    }

    # Get the mp3 file link of the podcast episode
    if($Link.href -match '.mp3') {
        # Extract the object name from the file name
        $mp3Name = ($Link.href.Split("/")[-1]).split(".")[0]
        $mp3SingleObject = $mp3Array | where {$_.Name -eq $mp3Name}
        # If the object doesn't exist, let's create it
        if(!$mp3SingleObject) {
            $mp3Object = New-Object System.Object
            $mp3Object | Add-Member -type NoteProperty -name Name -Value $mp3Name
            $mp3Object | Add-Member -type NoteProperty -name RemoteFileName $Link.href
            $mp3Object | Add-Member -type NoteProperty -name Guest ""
            $mp3Object | Add-Member -type NoteProperty -name Date ""
            $mp3Array += $mp3Object
        } else {
            # If the object exist, let's set the properties we want to keep
            $mp3SingleObject.Name = $mp3Name
            $mp3SingleObject.RemoteFileName = $Link.href
        }
    }
}

$PSScriptPath = Split-Path $PSCommandPath
$mp3DownloadPath = $PSScriptPath + "\Download\"
[System.Reflection.Assembly]::LoadFile($PSScriptPath+ "\powershell-taglib-master\taglib-sharp.dll" ) 

# Now that we have an array of mp3 objects with all the expected properties, let's check those we need to download 
Foreach($mp3Object in $mp3Array) {
    # Create a significant filename using the properties of the object created before
    $outFileName = $mp3DownloadPath + $mp3Object.Guest+" - " + $mp3Object.Name + ".mp3"
    # Check if the file has already been downloaded and if not, download it
    If(!(Test-Path $outFileName)){
        Invoke-WebRequest -Uri $mp3Object.RemoteFileName -OutFile $outFileName

        # Tag the file
        $mediaFile=[TagLib.File]::Create($outFileName)    
        $mediaFile.Tag.Album = "Passion Classique" 
        $mediaFile.Tag.Title = $mp3Object.Date
        $mediaFile.Tag.AlbumArtists = $mp3Object.Guest
        $mediaFile.Save()

    }

}


# That's it !!!



