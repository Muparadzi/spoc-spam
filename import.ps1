# Replace "C:\path\to\pst\file.pst" with the actual path to your PST file
$PSTFilePath = "C:\Users\Rory\Desktop\final.pst" # add in a varable to select which one you want to use


$OutlookProfileName = "Outlook Profile"

# Create a new Outlook Application object
$Outlook = New-Object -ComObject Outlook.Application

# Open the specified PST file
$PST = $Outlook.Session.AddStoreEx($PSTFilePath, 3)

# Get the root folder of the PST file
$RootFolder = $PST.GetRootFolder()

# Replace "SPAM" with the name of the folder where you want to import the PST contents
$DestinationFolder = $Outlook.Session.GetDefaultFolder([Microsoft.Office.Interop.Outlook.OlDefaultFolders]::olFolderInbox).Folders.Item("SPAM")

# Copy the contents of the PST file to the destination folder
$RootFolder.CopyTo($DestinationFolder)

# Close the PST file
$PST.Remove()

# Release the Outlook COM object
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Outlook) | Out-Null
