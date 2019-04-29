# GetDiskInfo

Description:

	- This script is an easy to use method to query remote machines for disk space utilization based on Active Directory OU membership.
	
Parameters:

	- The main wrapper file is called Get-DiskInfo.ps1. It is the main file that multi-threads the update process and generates the Excel reports. It takes the following parameters:
		-SearchBase: Mandatory string parameter. This is the distinguished name of the OU in Active Directory housing the target machines.
		-Credential: Mandatory PSCredential object. The PSCredential object used to run the jobs. Must be an admin on both the controller server and target machines.
		-NumJobs: Optional positive integer value. Optional parameter to set number of simultaneous jobs. If not set, the script will determine the number of simultaneous jobs based on the number of logical cores.

Pre-reqs:

	- This has only been tested with PoSH 4 and higher. PoSH v3.0 "should" work, but I haven't tested it. With that said, don't expect it to run on Server 2003/2008/XP/etc.
	- The controller server/workstation running this script should have internet access, specifically to the Microsoft PSGallery. If internet access from this machine exists, the appropriate modules will be installed automatically.
		- If internet access is not possible from the controller node, you will need to manually install the ImportExcel module from the PSGallery.
	- PS Remoting ports need to be open from the controller server/workstation to all target servers in order to pass the Invoke-Command cmdlet.
	- If running from the main launcher script, the account specified in the PSCredential object must have the ability to read from AD (no local accounts).
	- The PSCredential object passed to the main script must have admin rights on the target servers/workstations.
	- The script ignores Disabled computer accounts, so it doesn't try to connect to them.
	- The script does test connectivity to computer accounts, so if a computer account is enabled but not responding (i.e. powered off), it will also be ignored.
	
Files:

	- The file structure must remain as is in order to run the script. 
	- Test-TCPPort.ps1 tests a TCP socket connect to a specified TCP port. While not as robust as the built-in Test-NetConnection function, it returns much faster if the connection fails (important for scripts). It uses the System.Net.Sockets.TcpClient class to create the connection.
			Example 1:
			
				PS C:\Scripts\Projects\WindowsUpdate> Test-TCPport -ComputerName andy-2k16-vmm2 -TCPport 5985

				hostname         port open
				--------         ---- ----
				{andy-2k16-vmm2} 5985 True
				
Main Script Usage

	- Usage of the script is pretty easy. Simply pass the DN, the credential object and the number of simultaneous jobs to use.
		Example 1:
		
			PS> .\Get-DiskInfo.ps1 -SearchBase "ou=servers,dc=domain,dc=com" -Credential (Get-Credential) -NumJobs 16
				- The above command will query the disk drives on all enabled computer accounts in the Servers OU (maximum 16 jobs at a time), and it will export the list of available updates into an Excel spreadsheet.