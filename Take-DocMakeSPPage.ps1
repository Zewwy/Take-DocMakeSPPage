#region Author Information
########################################################################################################################## 
# Author: Zewwy (Aemilianus Kehler)
# Date:   May 8, 2020
# Script: Take-DocMakeSPPage
# This script takes a Word Doc and makes it a web page without SharePoint features.
# 
# Required parameters: SharePoint PnP Modules + Word Doc + Office Applications Installed
#   
#    
##########################################################################################################################
#endregion

#region Variables
##########################################################################################################################
#   Variables
##########################################################################################################################
$BadfileQry = "Sorry mate but it seems the dingo ate your file."
#MyLogoArray
$MylogoArray = @("#####################################","# This script is brought to you by: #","#                                   #","#             Zewwy                 #","#                                   #","#####################################"," ")
#Static Variables
$ScriptName = "Take-DocMakeSPPage; Cause the SharePoint Feature doesn't work.`n"
$TOCTemplate = @("​​​​<style>","H1 A, H2 A, H3 A {","color:inherit !important;","cursor:default;","text-decoration:none !important;","}","",".outer {","position:fixed;","margin:0px auto;","width:200px;","height:300px;","overflow:auto;","top:300px;","right:9%;","}","",".inner {","border-bottom:white 1px solid;","position:fixed;","overflow:auto;","border-left:white 1px solid;","width:200px;","}","</style>","<div class=""outer"">","   <div class=""inner"" id=""toc"">","      <h2>Table of Contents </h2>")
$pswheight = (get-host).UI.RawUI.MaxWindowSize.Height
$pswwidth = (get-host).UI.RawUI.MaxWindowSize.Width
#endregion

#region Functions
##########################################################################################################################
#   Functions
##########################################################################################################################

#function call to validate that required powershell mods exist on system
function Check4Mods()
{
    if ((Get-Module -ListAvailable -Name "SharePointPnPPowerShell2016") -or (Get-Module -ListAvailable -Name "SharePointPnPPowerShell2013") -or (Get-Module -ListAvailable -Name "SharePointPnPPowerShellOnline")) 
    {
        #Write-Host "Module exists"
        #If you need to come for something based on the modules actually existing add it here.
    } 
    else 
    {
        Centeralize "No SharePoint PnP Modules located on machine. This is a required module. Please install one." "red";exit
    }
}

#function call to validate that required powershell mods are loaded in the current session
function Check4Load()
{
    if (!(Get-Module "SharePointPnPPowerShell2016") -and !(Get-Module "SharePointPnPPowerShell2013") -and !(Get-Module "SharePointPnPPowerShellOnline")) 
    {
        Centeralize "Module Not Loaded!!! Load a module man, I'm too lazy to code for it right now" "Yellow";exit
    }
    else
    {
        #Write-Host "You sexy beast always got stuff ready to go!"
    }
}

#Validate System has assemblies, else exit
function ValidateAssemblies()
{
    try{$WordType = Add-Type -AssemblyName 'Microsoft.Office.Interop.Word' -Passthru}
    catch{Centeralize "This script requires Word assemblies; This means you're missing a required DLL AKA Install Word!" "Red";exit}
    try{[void][System.Reflection.Assembly]::LoadWithPartialName('Microsoft.Office.Interop.Word.WdSaveFormat')}
    catch{Centeralize "I have no clue...$_" "red";exit}
}

#Validate String is a valid URL
function isURL($URL) 
{
	($URL -as [System.URI]).AbsoluteURI -ne $null -and ($URL -as [System.URI]).Scheme -match "http|https"
}

#function takes in a name to alert confirmation of deletion of a web part, returns true or false
function confirm($name)
{
    #function variables, generally only the first two need changing
    $title = "Confirm Action!"
    $message = "$name"

    $yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", "This means Yes"
    $no = New-Object System.Management.Automation.Host.ChoiceDescription "&No", "This means No"

    $options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)

    $result = $host.ui.PromptForChoice($title, $message, $Options, 0)
    Write-Host " "
    Switch ($result)
        {
              0 { Return $true }
              1 { Return $false }
        }
}

#Function to Centeralize Write-Host Output, Just take string variable parameter and pads it
function Centeralize()
{
  param(
  [Parameter(Position=0,Mandatory=$true)]
  [string]$S,
  [Parameter(Position=1,Mandatory=$false,ParameterSetName="color")]
  [string]$C
  )
    $sLength = $S.Length
    $padamt =  "{0:N0}" -f (($pswwidth-$sLength)/2)
    $PadNum = $padamt/1 + $sLength #the divide by one is a quick dirty trick to covert string to int
    $CS = $S.PadLeft($PadNum," ").PadRight($PadNum," ") #Pad that shit
    if ($C) #if variable for color exists run below
    {    
        Write-Host $CS -ForegroundColor $C #write that shit to host with color
    }
    else #need this to prevent output twice if color is provided
    {
        $CS #write that shit without color
    }
}

#Function to ask for input color coded, and return results
function Get-UserInput($InputRequestTxt)
{
    Write-host $InputRequestTxt": " -ForegroundColor Magenta -NoNewline
    $UserInput = Read-Host
    Return $UserInput
}

#function to valid user input is a doc
function GetSourceDoc()
{
  $global:SourceDoc = Get-UserInput ("(E.G: C:\word.docx) Source Doc File")
  if([string]::IsNullOrWhiteSpace($SourceDoc) `
  -or !(Test-Path $SourceDoc) `
  -or !(Test-Path $SourceDoc -PathType Leaf) `
  -or (Get-Item $SourceDoc).Extension -notmatch "doc|docx"){GetSourceDoc}
  #-or !(Get-Item $SourceDoc).Extension match "doc|docx" **DOES NOT WORK HAVE TO USE -notmatch**
}

#function to validate userinput is a folder
function GetDestinationPath()
{
  $global:DesPath = Get-UserInput ("(E.G: C:\temp) Destination Path")
  if([string]::IsNullOrWhiteSpace($DesPath) `
  -or !(Test-Path $DesPath) `
  -or !(Test-Path $DesPath -PathType Container)){GetDestinationPath}
}

#Request URL for sharepoint upload and validate it is a URL using isURL
function GetSPSitePath()
{
  $global:SPPath = Get-UserInput ("(E.G: https://sharepoint.domain.com/subsite) SharePoint Path")
  if(!(isURL $SPPath)){GetSPSitePath}
}

function SaveDocAsHTML($DaDoc, $DaDestination)
{
    #Set the save type to HTML create the base HTML code and extracts the images to a folder
    $saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatHTML");
    #Create a Word App Object to open the doc file for processing
    $wordApp = new-object -comobject word.application
    #Ensure it doesn't open on the running desktop
    $wordApp.Visible = $false
    #Open it for processing
    $openDoc = $wordApp.documents.open($DaDoc.FullName)
    if($DaDestination.FullName -notmatch '\\$'){$GoodPath = $DaDestination.FullName+'\'}
    $BaseName = $DaDoc.BaseName
    $SaveFileName = $GoodPath + $BaseName+".html"
    #$SaveFileName
    $openDoc.saveas([ref]"$SaveFileName", [ref]$saveFormat);
    $wordApp.quit()
}

function ConnectToSPandValidatePaths()
{
    #Ask for the SharePoint Site which will be used to upload the images and create the page on
    GetSPSitePath
    #Try to connect to the SharePoint Site
    Try{Connect-PnPOnline SPPath}
    Catch{Centeralize "Something went wrong connecting to the SharePoint Site... $?" "Red";exit}
    GetSPImageFolder
}

function GetSPImageFolder()
{
    $global:SPImageUploadDocLibrary = Get-UserInput ("(E.G: PublishingImages) Destination Path")
    Try{$ValidSPDocLib = Get-PnPFolder -Url $SPImageUploadDocLibrary -ErrorAction Stop}
    Catch{Centeralize "Looks like that Doc Library doesn't exist, try again" "Yellow";GetSPImageFolder}
}

function UploadImagesToSP($DaImages)
{
    foreach ($image in $DaImages)
    {
        Write-Host $image.Name" Upladed to "$TheRightPath
        #Uncomment this one line to actually upload the files
        Add-PnPFile -Folder "$CmdReqPath" -Path $image.FullName
    } 
    Write-Host "Done"
}

function FixLocalHTMLImgPaths()
{
    $SPSourceHTML = $TheDoc.FullName -replace ($TheDoc.Extension,".html")
    Write-host $SPSourceHTML
    Write-Host $TheRightPath
    Centeralize "Verifying File Path, Please Wait...`n" "White"
    if (Test-Path $SPSourceHTML) #Check if file exists
    {
        $SourceFile = Get-ChildItem $SPSourceHTML
        Centeralize "Scanning txt file for Headers and Image links`n" "Cyan"
        $DataContent = @() #Ensures these variables are instantiated as arrays, and are global outside the foreach loop
        foreach($Line in Get-Content $SPSourceHTML) #Reads each line of the source file one by one
        {
            if($Line -NotMatch "_files") #Checks the line for the headers by looking for the known HTML header tags, if the line doesn't contain a _files extention, create the data content as normal
            {
            $DataContent = $DataContent + $line
            }
            else
            {
                $ArrayOfPath = ($TheRightPath -split "/")
                $ArrayOfPath = $ArrayOfPath[3..($ArrayOfPath.length -1)]
                foreach($Item in $ArrayOfPath)
                {
                    $FullLine = $FullLine + "/" + $Item
                }
                $DaVars = $SourceFile.BaseName + "_files"
                $EncodeSource = [uri]::EscapeUriString($DaVars)
                $EncodeTarget = [uri]::EscapeUriString($FullLine)
                $FullLine = ""
                $CMONFFS = $line.replace($EncodeSource,$EncodeTarget)
                $DataContent = $DataContent + $CMONFFS #Else we remove this garbage from MS
            }
        }
        #$DataContent
        $DataContent | Out-File $TheDestination\HTMLCleanImagePaths.txt
    }
}
#endregion

#region Running Code
##########################################################################################################################
#   Start of script
##########################################################################################################################

#region Show Logo
#Start actual script by posting and asking user for responses
foreach($L in $MylogoArray){Centeralize $L "green"}
Centeralize $ScriptName "White"
#endregion

#region Validation Function Calls
#Validate we have SharePoint Modules Installed
Check4Mods

#Validate that we have SharePoint Modules loaded
Check4Load

#Validate we can work with doc files
ValidateAssemblies
#endregion

#region Ask for Source File and Destination Folder
#Congrats you have what it takes to move on, where's your file?
GetSourceDoc
$TheDoc = Get-Item $SourceDoc
Write-Host ""

#Ask for converted files location destination
GetDestinationPath
$TheDestination = Get-Item $DesPath
Write-Host ""
#endregion

#region Confirm Conversion to HTML
Centeralize "All required parameters have been met, proceed to convert Doc to HTML?" "Green"
Centeralize "This step simply uses word dll's to create the base source HTML and extract the images to folder." "White"
if(confirm "Save $TheDoc as HTML file to $TheDestination`?")
{
    SaveDocAsHTML $TheDoc $TheDestination
}
else
{
    Centeralize "Without converting the document to HTML the rest of this script is useless." "Red"
    exit
}
#endregion

#region Connect to SP and Validate Doc Library Provided
if(confirm "Connect to SharePoint Site?")
{
    ConnectToSPandValidatePaths
    #region Populate Information and Confirm Upload to SP
    #We need the base name of the doc we converted to get the images folder name which always ends with _files
    $AYBO=$TheDoc.BaseName
    $ImageFolder = Get-ChildItem $TheDestination | ?{$_.Name -match "$AYBO`_files"}
    #We also only care about the images so we filter on them cause they are always named image#
    $Images = Get-ChildItem $ImageFolder | ?{$_.Name -match "image"}
    $CmdReqPath = $SPImageUploadDocLibrary+"/"+$AYBO
    $TheRightPath = $SPPath+"/"+$CmdReqPath
    $WordUP = "Upload the total of " + $Images.Count + " images to " + $TheRightPath + "?"
    If(confirm $WordUp){UploadImagesToSP $Images}
    #endregion
}
else
{
    Centeralize "Doc converted to HTML but only on local system, No SharePoint Page created." "Yellow"
}
#endregion

#region Alter Local HTML Image Paths
if(confirm "Alter local HTML image Paths?"){FixLocalHTMLImgPaths}
#endregion

#region Ask to Delete local created files
$AYBO=$TheDoc.BaseName
$OldFolder = "$TheDestination"+"\"+$AYBO+"_files"
$OldFile = $TheDoc.FullName -replace ($TheDoc.Extension,".html")
if(confirm "Delete local HTML file and folder?"){(Get-Item $OldFolder).Delete($true);(Get-Item $OldFile).Delete()} 
else{Centeralize "Script is done but old files remain on your local machine. $OldFolder and $OldFile" "Yellow"}
#endregion

#endregion
