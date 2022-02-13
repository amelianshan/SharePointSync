# SharePointSync
Sync Sharepoint Site to the Local File Share. It can exclude some specific file type, like exe file, bat file for security reason.
Before running the script, you need to install sharepoint pnp powershell module.
You need a user account who has full access to the SharePoint Site and no expired password.
The file will be updated when the last edit time is later than local file's last edit time. Or it's a new file/folder.
