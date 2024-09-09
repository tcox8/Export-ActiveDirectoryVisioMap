# Export-ActiveDirectoryVisioMap
Exports AD OUs and GPOs to a Visio Map


# Version 1.3
Editor : Kyle Schuler</br>
Fork of : https://github.com/tcox8/Export-ActiveDirectoryVisioMap


# Requirements
* RSAT tools for Active Directory and GPO. </br>
* VisioAutomation - https://github.com/saveenr/VisioAutomation  (this module is imported in the script but I wanted to give mention to Saveenr and all his hard work). </br>
* A working copy of Visio installed.</br>
* Active Directory Visio Stencil (Possibly, need to look into this)


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
Ver 1.1 - Fixed Visio Cmdlet Parameters, Adjusted for Azure AD joined devices, Fixed issue with importing Visio module, Reduced output to console</br>
Ver 1.2 - Refactored, reformatted, added outputs and more error handling</br>
Ver 1.3 - Added options for user to include or exclude GPOs, and choose the direction of the layout</br>

# Known Issues
Visio AD Stencil not autoloading properly or not loading if not installed