# Create-OutlookSignature

This script will query the domain to find organizational information within the context of the user who runs it. 
Best used as a logon script via GPO.


For example, it can be used to find their name, their phone number, their job title, or custom extensionAttribute's.
It will then create a standardized outlook signature using the supplied $Body data and AD info, which can be applied it to their outlook and OWA settings via args.


You can change some of the behaviors by changing the registry variables.
Currently supports Outlook 2016/2019.
	
	
This is meant to be edited for your own use.
