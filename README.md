# Export-ActiveDirectoryVisioMap
Exports AD OUs and GPOs to a Visio Map

# Version 1.1
Editor : Kyle Schuler </br>
Original Author: tcox8

# Requirements
RSAT tools for Active Directory and GPO. </br>
VisioAutomation - https://github.com/saveenr/VisioAutomation  (this module is imported in the script but I wanted to give mention to Saveenr and all his hard work). </br>
A working copy of Visio installed.

</br>
</br>
</br>

Example:
![Example Picture](ExampleImages/ExamplePicture.PNG?raw=true)


All shapes in visio will have details in the Shape Data

Example OU:
![Example OU Details](ExampleImages/ExampleOUdetails.png?raw=true)

Example GPO:
![Example GPO Details](ExampleImages/ExampleGPOdetails.png?raw=true)


# Change Log: 
Ver 1.0 - Initial release</br>
Ver 1.1 - Fixed Visio Cmdlet Parameters, Adjusted for Azure AD joined devices, Fixed issue with importing Visio module, Reduced output to console
