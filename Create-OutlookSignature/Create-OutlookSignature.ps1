<#
  .SYNOPSIS
	Created outlook signatures based on supplied AD information.

  .DESCRIPTION
	This script will query the domain to find organizational information within the context of the user who runs it. Best used as a logon script via GPO.
	For example, it can be used to find their name, their phone number, their job title, or custom extensionAttribute's
	It will then create a standardized outlook signature using the supplied $Body data and AD info, which can be applied it to their outlook and OWA settings via args.

	You can change some of the behaviors by changing the registry variables.
	Currently supports Outlook 2016/2019.
	
	This is meant to be edited for your own use.
  .PARAMETER NoWriteReg
	The script will still generate the files but it will not write the settings to the registry.
	Useful if you just want to generate a standardized signature without setting it in stone on your user's profile.
  
  .INPUTS
	None. You cannot pipe objects to Create-OutlookSignature.ps1.

  .OUTPUTS
	Will copy needed signature files from the set directory, as well as generate a signature htm file with it's related filelist.xml file.

  .EXAMPLE
	PS> .\Create-OutlookSignature.ps1 -NoWriteReg
#>

param (
	[switch]$NoWriteReg = $False
)

# Created by: ShortToGround
# Date created: 03/15/2021


# First let's find the user's information
# I did it this way to avoid installing the AD module on each machine
$Username = "$env:USERNAME" # Important that this is ran in user's context
$SearcherFilter = "(&(objectCategory=User)(samAccountName=$UserName))"
$Searcher = New-Object System.DirectoryServices.DirectorySearcher
$Searcher.Filter = $SearcherFilter
$ADUserPath = $Searcher.FindOne()
$UserInfo = $ADUserPath.GetDirectoryEntry()


# Now we pull the desired information
# EDIT TO YOUR NEEDS
$Name = $UserInfo.Name
$JobTitle = $UserInfo.Title
$OutsidePhoneNumber = $UserInfo.Telephonenumber
# If you store custom attributes in their extensionAttribute you can pull it here as well
#$Example1 = $UserInfo.extensionAttribute1


# If you signature contains images or logos or etc they will be stored in a dir with this name
$UsernameFiles = "$Username"+"_files"

# Name of the actual signature htm file
$HTMFileName = "$Username"+".htm"


# In my case, we stored the acronym of the certifications each user's holds in their extensionAttribute1 attributes
# This is an example of how I would edit the html data depending on if they had certifications or not
# Feel free to delete this if you don't need it.
# In my case it would add the acronyms to the end of their name like "John Doe, ABC, DEF"
# This would have been inserted into the $Body section where their name would have been 
# If ($Example1 -ne $Null){
	# $NameHTML = "<span
  # style='font-size:14.0pt;color:#365F91'>$Name, </span></b></span><span
  # style='mso-bookmark:_MailAutoSig'></span><a
  # href=""http://TheirAcronymsHadLinks.Too""><span
  # style='mso-bookmark:_MailAutoSig'><b><span style='font-size:12.0pt;
  # color:#376092;text-decoration:none;text-underline:none'>$Certifications</span></b></span></a><span
  # style='mso-bookmark:_MailAutoSig'><span style='font-size:12.0pt;mso-no-proof:
  # yes'>"
# }
# Else{
	# $NameHTML = "<span
  # style='font-size:14.0pt;color:#365F91'>$Name</span></b></span>
  # style='mso-bookmark:_MailAutoSig'></span>"
# }


# In our case I needed a single standard signature template. So I took a known good signature, and extracted the html data here.
# It's quite long so I didn't feel like sanitizing it and including it
# This is also an example of how easy it is to add in the html output from the example above
$Body = @"
  $NameHTML
  
"@






# This is usually the default place Outlook saves signatures, at least for 2016/2019.
$SignatureFolder = "$env:APPDATA\Microsoft\Signatures"
$SignatureFilesFolder = "$env:APPDATA\Microsoft\Signatures\$UsernameFiles"

# You will need to point this to a share that can be accessed by the user. I chose a SYSVOL location personally, since this was running as a logon script.
# I'm not totally sure but I think all html signatures will have this folder
$FilesLocation = ""



# Let's get a list of all of the files included in the $FilesLocation dir
# then we can being creating a valid filelist XML file
$Files = Get-ChildItem $FilesLocation
$XMLData = ""
ForEach ($Item in $Files){
	$XMLData += "`r`n<o:File HRef=`"$Item`"/>"
}

# I noticed that each signature had a "unique" XML file in their file directory
# So this block is used to generate said XML file
# This should always match your files in $FilesLocation due to the above operation
$filelistXML = @"
 <xml xmlns:o="urn:schemas-microsoft-com:office:office">
 <o:MainFile HRef="../$HTMFileName"/> $XMLData
</xml>
"@





# Really quick and dirty error checking. Don't want to create a signature file if some of the key information is missing.
# Don't judge for the bad error checking :)
If (($Name -eq $Null) -or ($Name -eq "") -or ($JobTitle -eq $Null) -or ($JobTitle -eq "") -or ($OutsidePhoneNumber -eq $Null) -or ($OutsidePhoneNumber -eq "")){
	Write-Host "Required information is missing, aborting..."
	Exit
}


Else{
	### Next version I'm going to move this up in the chain, no need to do all of this processing if they already have the correct files and everything. ###
	# Let's check to see if the files already exist on the user's profile
	If (!(Test-Path -Path $SignatureFilesFolder)){
		New-Item -Path "$SignatureFolder" -Name "$UsernameFiles" -ItemType "Directory"
		Copy-Item "$FilesLocation\*" -Destination "$SignatureFilesFolder" -Recurse
		$filelistXML | Out-File "$SignatureFilesFolder\filelist.xml"
	}
	# If they do have it, lets compare file hashes so we don't have to copy files again
	Else{
		
		# Probably a better way to do this
		# I'm just adding all the hashes into a single string and comparing at the end
		$HashList1 = ""
		$HashList2 = ""

		$Files1 = Get-ChildItem $FilesLocation
		ForEach ($Item1 in $Files1){
			if ($Item1.Name -ne "filelist.xml"){
				$Hash1 = Get-FileHash "$FilesLocation\$Item1" | ForEach { $_.Hash }
				$HashList1 += "$Hash1"
			}
		}
			
		$FilesLocation2 = $SignatureFilesFolder
		$Files2 = Get-ChildItem $SignatureFilesFolder
		ForEach ($Item2 in $Files2){
			if ($Item2.Name -ne "filelist.xml"){
				$Hash2 = Get-FileHash "$SignatureFilesFolder\$Item2" | ForEach { $_.Hash }
				$HashList2 += "$Hash2"
			}
		}	

		if ($HashList1 -ne $HashList2){
			Copy-Item "$FilesLocation\*" -Destination "$SignatureFilesFolder" -Recurse
			$filelistXML | Out-File "$SignatureFilesFolder\filelist.xml"
		}
	
	}
	
	
	
	
	
	# I decided to rewrite the signature.htm every logon to make sure its up to date with AD info
	$Body | Out-File "$env:APPDATA\Microsoft\Signatures\$HTMFileName"
	
	If (!$NoWriteReg){
		# Setting the common office settings means the users will not be able to change this signature or select a new one
		# Should probably just check these before writing but oh well 
		Set-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Office\16.0\Common\MailSettings" -Name "NewSignature" -Value "$env:USERNAME" -Type "ExpandString"
		Set-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Office\16.0\Common\MailSettings" -Name "ReplySignature" -Value "$env:USERNAME" -Type "ExpandString"
	}
}






