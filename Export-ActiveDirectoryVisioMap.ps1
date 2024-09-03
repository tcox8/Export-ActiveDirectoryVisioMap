#############################################################################
# Author  : Tyler Cox
#
# Version : 1.0
# Created : 11/2/2021
# Modified : 
#
# Purpose : This script will build an inventory of all GPOs and their links.
#
# Requirements: A computer with Active Directory Admin Center (ADAC) installed and a 
#               user account with enough privileges 
#             
# Change Log: Ver 1.0 - Initial release
#
#############################################################################


#Import the modules
Try 
    {   
        Import-Module ActiveDirectory -ErrorAction Stop
    }
Catch 
    {
        Write-Host "Error! Could not import ActiveDirectory module! Please make sure you are running this as an Administrator and that RSAT tools are installed!"
        break
    }
Try 
    {   
        Import-Module GroupPolicy -ErrorAction Stop
    }
Catch 
    {
        Write-Host "Error! Could not import GroupPolicy module! PLease make sure you are running this as an Administrator and that RSAT tools are installed!"
        break
    }
Try 
    {   
        Import-Module Visio -ErrorAction Stop
    }
Catch 
    {
        Write-Host "Error! Could not import Visio module!"
        break
    }


#Create the Visio Application
New-VisioApplication
#Create the Visio Document
$VisioDoc = New-VisioDocument
#Create the Visio Page
$Page = $VisioDoc.Pages[1]
#Create the Visio Point at 1,1
$Point_1_1 = New-VisioPoint -X 1.0 -Y 1.0

#Set our counters
$nodeCount = 0
$conCount = 0
$gpoCount = 0

#Get our root domain from the current logged on user
$DNSDomain = $env:USERDNSDOMAIN 

#Get all OUs except LostAndFound
$OUs = Get-ADOrganizationalUnit -Server $DNSDomain -Filter 'Name -like "*"' -Properties Name, DistinguishedName, CanonicalName, LinkedGroupPolicyObjects | `
    Where-Object {$_.canonicalname -notlike "*LostandFound*"} | Select-Object Name, Canonicalname, DistinguishedName, LinkedGroupPolicyObjects | `
    Sort-Object CanonicalName # | Select -First 50


#Gather our shapes from Visio's stencils
$ADO_u = Open-VisioDocument "ADO_U.vss"
$connectors = Open-VisioDocument "Connectors.vss"
$masterOU = Get-VisioMaster "Organizational Unit" -Document $ADO_u
$connector = Get-VisioMaster "Dynamic Connector" -Document $Connectors
$masterDomain = Get-VisioMaster "Domain" -Document $ADO_u
$masterGPO = Get-VisioMaster "Policy" -Document $ADO_u

#Create our first shape. This is the root domain node
$n0 = New-VisioShape -Master $MasterDomain -Position $Point_1_1
#Set shape properties
$n0.Text = $DNSDomain
$n0.Name = "n" + $DNSDomain

#Get Root Domain linked GPOs and process them accordingly
$RootGPOs = Get-ADObject -Server $DNSDomain -Identity (Get-ADDomain -Identity $DNSDomain).distinguishedName -Properties name, distinguishedName, gPLink, gPOptions
#Loop through each root GPO
ForEach ($gpolink in $RootGPOs.gPlink -split "\]\[")
    {
        #Add to our counters (for naming)
        $gpoCount += 1 
        $conCount += 1 
        #get only the GUID of the gpo
        $gpoGUID = ([Regex]::Match($gpoLink,'{[a-zA-Z0-9]{8}[-][a-zA-Z0-9]{4}[-][a-zA-Z0-9]{4}[-][a-zA-Z0-9]{4}[-][a-zA-Z0-9]{12}}')).Value 
        #pull details for the GPO based on the GUID
        $gpo = Get-GPO -GUID $gpoGUID -Domain $DNSDomain

        #declare what we'll call the gpo shape 
        $shapename = "g" + $gpoCount 
        #Create the GPO shape
        $shapeGPO = New-VisioShape -Master $MasterGPO -Position $Point_1_1 
        #Set the shape properties
        $ShapeGPO.Text = $GPO.DisplayName
        $ShapeGPO.Name = $shapename
        #Set the shape's custom properties
        $GUID = "{" + $gpo.id.guid + "}"
        If ($GPO.DisplayName) {Set-VisioCustomProperty -Shape $ShapeGPO -Name "GPOName" -Value $GPO.DisplayName}
        If ($GPO.Description) {Set-VisioCustomProperty -Shape $ShapeGPO -Name "Description" -Value $GPO.Description}
        If ($GPO.ID.Guid) {Set-VisioCustomProperty -Shape $shapeGPO -Name "GUID" -Value $GUID}
        If ($GPO.GPOStatus) {Set-VisioCustomProperty -Shape $shapeGPO -Name "Status" -Value $GPO.GpoStatus.ToString()}
        If ($GPO.CreationTime) {Set-VisioCustomProperty -Shape $shapeGPO -Name "CreationTime" -Value $GPO.CreationTime.ToString()}
        If ($GPO.ModificationTime) {Set-VisioCustomProperty -Shape $shapeGPO -Name "ModifiedTime" -Value $GPO.ModificationTime.ToString()}
        If ($GPO.WmiFilter) {Set-VisioCustomProperty -Shape $shapeGPO -Name "WMIFilterName" -Value $GPO.WMIFilter.Name}
        #Create the shape's connections
        $con = Connect-VisioShape -From $n0 -To $shapeGPO -Master $connector 
        #Set the connections custom properties
        $con.text = "GPO"
        $con.name = "gcon" + $conCount #We name it like this so that later we can identify all GPO connections for formatting of the connector's text
        $con_cells = New-VisioShapeCells
        $con_cells.LineColor = "rgb(0,175,240)"
        $con_cells.LineEndArrowSize = "3"
        $con_cells.LineBeginArrowSize = "2"
        $con_cells.LineEndArrow = "42"
        $con_cells.LineBeginArrow = "4"
        $con_cells.CharColor = "rgb(0,175,240)"
        #Set the shape properties
        Set-VisioShapeCells -Cells $con_cells -Shape $con     
    }




#Loop through each OU
ForEach ($ou in $OUs)
    {
        #Add to our counters
        $nodeCount += 1
        $conCount += 1
     
        #Massage the OU details to get the name
        $OUName = $OU.Name
        #Massage the OU details to get the Canonical name. We use this to get the previous OU name
        $OUConName = $OU.Canonicalname
        $nameSplit = $ou.CanonicalName -split '(?<!\\)/'
        $nameRecombined = $nameSplit[0..($nameSplit.length-2)] -join "/"
        #If the previous OU name is the root domain..
        If ($nameSplit[$index-2] -eq $DNSDomain)
            {
                #declare what we'll call the shape
                $shapename = "n" + $OUConName
                #Create the new shape
                $shape = New-VisioShape -Master $MasterOU -Position $Point_1_1
                #Set the shape details
                $Shape.Text = $OUName
                $Shape.Name = $shapename
                
                #Set custom properties of the shape
                Set-VisioCustomProperty -Shape $shape -Name "OU_Name" -Value $OU.Name
                Set-VisioCustomProperty -Shape $shape -Name "DistinguishedName" -Value $OU.DistinguishedName
                Set-VisioCustomProperty -Shape $shape -Name "Linked_GPOs" -Value $OU.LinkedGroupPolicyObjects.Count
                
                #Connect the shape to the root domain shape
                Connect-VisioShape -From $n0 -To $shape -Master $connector | Out-Null

            }
        #If it's not the root domain, then do this..
        else 
            {
                #Set the name of the previous shape 
                $prevOUName = "n" + $nameRecombined

                #Get the previous shape from Visio based on the name
                $prevOUshape = Get-VisioShape -Name * | Where-Object {$_.Nameu -eq $prevOUName}

                #Set the name of the new shape
                $shapename = "n" + $OUConName
                #Create the new shape
                $shape = New-VisioShape -Master $MasterOU -Position $Point_1_1
                #Set the shape properties
                $Shape.Text = $OUName
                $Shape.Name = $shapename

                #Set custom properties of the shape
                Set-VisioCustomProperty -Shape $shape -Name "OU_Name" -Value $OU.Name
                Set-VisioCustomProperty -Shape $shape -Name "DistinguishedName" -Value $OU.DistinguishedName
                Set-VisioCustomProperty -Shape $shape -Name "Linked_GPOs" -Value $OU.LinkedGroupPolicyObjects.Count

                #Connect the shape to the previous shape
                Connect-VisioShape -From $prevOUshape -To $shape -Master $connector | Out-Null
            }

        #If the OU had linked GPOs..
        If ($OU.LinkedGroupPolicyObjects)
            {
                #Loop through each GPO
                Foreach ($gpoLink in $OU.LinkedGroupPolicyObjects)
                    {
                        #increase our counters
                        $gpoCount += 1
                        $conCount += 1
                        #get only the GUID of the gpo
                        $gpoGUID = ([Regex]::Match($gpoLink,'{[a-zA-Z0-9]{8}[-][a-zA-Z0-9]{4}[-][a-zA-Z0-9]{4}[-][a-zA-Z0-9]{4}[-][a-zA-Z0-9]{12}}')).Value
                        #Create the GPO shape
                        $gpo = Get-GPO -GUID $gpoGUID -Domain $DNSDomain

                        #declare what we'll call the gpo shape 
                        $shapename = "g" + $gpoCount
                        #Create the GPO shape
                        $shapeGPO = New-VisioShape -Master $MasterGPO -Position $Point_1_1
                        #Set the shape properties
                        $ShapeGPO.Text = $GPO.DisplayName
                        $ShapeGPO.Name = $shapename
                        $GUID = "{" + $gpo.id.guid + "}"
                        If ($GPO.DisplayName) {Set-VisioCustomProperty -Shape $ShapeGPO -Name "GPOName" -Value $GPO.DisplayName}
                        If ($GPO.Description) {Set-VisioCustomProperty -Shape $ShapeGPO -Name "Description" -Value $GPO.Description}
                        If ($GPO.ID.Guid) {Set-VisioCustomProperty -Shape $shapeGPO -Name "GUID" -Value $GUID}
                        If ($GPO.GPOStatus) {Set-VisioCustomProperty -Shape $shapeGPO -Name "Status" -Value $GPO.GpoStatus.ToString()}
                        If ($GPO.CreationTime) {Set-VisioCustomProperty -Shape $shapeGPO -Name "CreationTime" -Value $GPO.CreationTime.ToString()}
                        If ($GPO.ModificationTime) {Set-VisioCustomProperty -Shape $shapeGPO -Name "ModifiedTime" -Value $GPO.ModificationTime.ToString()}
                        If ($GPO.WmiFilter) {Set-VisioCustomProperty -Shape $shapeGPO -Name "WMIFilterName" -Value $GPO.WMIFilter.Name}

                        #Create the shape's connections
                        $con = Connect-VisioShape -From $shape -To $shapeGPO -Master $connector
                        $con.text = "GPO"
                        $con.Name = "gcon" + $conCount #We name it like this so that later we can identify all GPO connections for formatting of the connector's text
                        $con_cells = New-VisioShapeCells
                        $con_cells.LineColor = "rgb(0,175,240)"
                        $con_cells.LineEndArrowSize = "3"
                        $con_cells.LineBeginArrowSize = "2"
                        $con_cells.LineEndArrow = "42"
                        $con_cells.LineBeginArrow = "4"
                        $con_cells.CharColor = "rgb(0,175,240)"
                        #Set the shape properties
                        Set-VisioShapeCells -Cells $con_cells -Shape $con                      
                    }
            }
    }

#Create a new layout object
$ls = New-Object VisioAutomation.Models.LayoutStyles.hierarchyLayoutStyle
#set object properties (this is how we format the page)
$ls.AvenueSizeX = 1
$ls.AvenueSizeY = 1
$ls.LayoutDirection = "ToptoBottom"
$ls.ConnectorStyle = "Simple"
$ls.ConnectorAppearance = "Straight"
$ls.horizontalAlignment = "Left"

#Apply the layout object to the page
Format-VisioPage -LayoutStyle $ls 

#Change the page's size to match the new data
Format-VisioPage -FitContents -BorderWidth 1.0 -BorderHeight 1.0

#This section is to set text for the GPO shapes based on the length of the line. We had to move the shapes around first before we could run this part.
#Create a new Shape Cell Object
$con_cells = New-VisioShapeCells
#Set the location of the text based on the length of the line
$con_cells.TextFormPinX = "=POINTALONGPATH(Geometry1.Path,1)"
$con_cells.TextFormPinY = "=POINTALONGPATH(Geometry1.Path,.75)"
#Get all gpo connections
$gpoShapes = Get-VisioShape -Name * | Where-Object {$_.Nameu -like "gcon*"}
#Loop through each connection
ForEach($shape in $gpoShapes)
    {
        #Set the shape from the shape cell object
        Set-VisioShapeCells -Cells $con_cells -Shape $shape   
    }
