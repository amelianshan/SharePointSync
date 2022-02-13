$fileExtensionToExclude = "bat","ps1","bin","exe","vbs","vb","com","msi","cmd"

Function LogMessage
{
    param([string]$Message)
    ((Get-Date).ToString() + " - " + $Message) >> $LogFile;
}

Function downloadSPFolder($folderUrl)
{
	$oldLocation = Get-Location
    Set-Location -Path $SharedDriveFolderPath
    $folderColl=Get-PnPFolderItem -FolderSiteRelativeUrl $folderUrl -ItemType Folder

    #Extract all name property
    $folderCollNames =  $folderColl | Select-Object -ExpandProperty Name

    # Loop through the folders

    if([string]::IsNullOrEmpty($folderCollNames)){
        $localfolders = Get-ChildItem -Directory $folderUrl
        foreach($localfolder in $localfolders){

            #If local folder name is not within SharePoint folder name, Remove it from Local
            if(!($folderCollNames -contains $localfolder)){

                $folderpath = $SharedDriveFolderPath + "/" + $folderUrl + "/" + $localfolder

                Remove-Item -Recurse $folderpath

                echo "Remove Folder: $folderpath"
                LogMessage -Message "Removed Folder: $folderpath"
            }
        }


    }


    foreach($folder in $folderColl)
    {
       $newFolderURL= $folderUrl+"/"+$folder.Name
       # Call the function to get the folders inside folder
	    if (-Not (Test-Path -Path $newFolderURL)) {
            $newFolder = new-item $newFolderURL -itemtype directory
            echo newFolder=$newFolder
            LogMessage -Message "newFolder: $newFolder"
            $newFolder.CreationTime=$folder.TimeLastModified
        }


        #Get All folders under current directories
        $localfolders = Get-ChildItem -Directory $folderUrl

        foreach($localfolder in $localfolders){

            #If local folder name is not within SharePoint folder name, Remove it from Local
            if(!($folderCollNames -contains $localfolder)){

                $folderpath = $SharedDriveFolderPath + "/" + $folderUrl + "/" + $localfolder

                Remove-Item -Recurse $folderpath

                echo "Remove Folder: $folderpath"
                LogMessage -Message "Removed Folder: $folderpath"
            }
        }

       downloadSPFolder($newFolderURL)
    }
    $SPFiles = Get-PnPFolderItem -FolderSiteRelativeUrl $folderUrl -ItemType File

    #Extract Name property from all the files
    

    foreach($SPFile in $SPFiles) {
        
        $SPfileExtension = $SPFile.Name.split('.')[1]

        #Exclude potential harmful file extensions

        if(!($fileExtensionToExclude.Contains($SPfileExtension)) ){
        
        $SPFilePath = $SharedDriveFolderPath+"/"+$folderUrl+"/"+$SPFile.Name
        if (Test-Path -Path $SPFilePath) {
            $localFile = Get-Item $SPFilePath
            if (($localFile.LastWriteTime -ge $SPFile.TimeLastModified.AddHours(-5) ) -and ($localFile.Length -eq $SPFile.Length)) {
                continue
            }
        }
		$SPFileSize = $SPFile.Length
		LogMessage -Message "Downloading File: $SPFilePath [$SPFileSize]";
		echo "Downloading File: $SPFilePath [$SPFileSize]"
		Get-PnPFile -Url $SPFile.ServerRelativeUrl -Path $folderUrl -FileName $SPFile.Name -AsFile -Force
		$localFile = Get-Item $SPFilePath
		# TODO convert SP File UTC time to US East
		$localFile.LastWriteTime = $SPFile.TimeLastModified.AddHours(0)

        }
	}

    #Get Local Directory Files
    $SPFileNames =  $SPFiles | Select-Object -ExpandProperty Name
    $localFiles =  Get-ChildItem -File $folderUrl

    foreach($localFile in $localFiles){
        

        #Remove file if name is not in SharePoint file name collection
        if(!($SPFileNames -contains $localFile)){

             $filepath = $SharedDriveFolderPath + "/" + $folderUrl + "/" + $localfile
            Remove-Item $filepath

            echo "Remove File: $filepath"
            LogMessage -Message "Removed File: $filepath"
            

        }

    }
    
	Set-Location -Path $oldLocation
}

#Net Use
$SharePointSiteURL = "https://xxx.sharepoint.com/sites"  
# Change this SharePoint Site URL  
$SharedDriveFolderPath = "E:\Sharepoint"  
# Change this Network Folder path  
#$SharePointFolderPath = "Shared Documents/TP"
$SharePointFolderPath = "Shared Documents/TP"
$timestamp =  get-date -f yyyy-MM-dd
$LogFile = "E:\SharepointLogs\FromSharepointToE\SharepointSyncLog_$timestamp.log"

$username = "username@email.com"

$password = "xxxxxxxxx"
$password = ConvertTo-SecureString -String $password -AsPlainText -Force

$Credential = New-Object System.Management.Automation.PSCredential $username,$password

Connect-PnPOnline -Url $SharePointSiteURL -Credentials $Credential
downloadSPFolder($SharePointFolderPath)

$timestamp =  get-date -f yyyy-MM-dd

$directories = Import-Csv E:\Scripts\robocopy.csv

foreach($directory in $directories){

    robocopy $directory.Source $directory.Destination /S /E /DCOPY:T /COPY:DATOU /MIR /B /J /NP /PF /R:0 /W:0 /log:E:\SharepointLogs\$timestamp.log

}

