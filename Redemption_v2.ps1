function Connect-Exchange
{ 
    param( 
    	[Parameter(Position=0, Mandatory=$true)] [string]$MailboxName
    )  
 	Begin
		 {
		## Load Managed API dll  
		###CHECK FOR EWS MANAGED API, IF PRESENT IMPORT THE HIGHEST VERSION EWS DLL, ELSE EXIT
		#$EWSDLL = (($(Get-ItemProperty -ErrorAction SilentlyContinue -Path Registry::$(Get-ChildItem -ErrorAction SilentlyContinue -Path 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Exchange\Web Services'|Sort-Object Name -Descending| Select-Object -First 1 -ExpandProperty Name)).'Install Directory') + "Microsoft.Exchange.WebServices.dll")
		$EWSDLL = "C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll"
        if (Test-Path $EWSDLL)
		    {
		    Import-Module $EWSDLL
		    }
		else
		    {
		    "$(get-date -format yyyyMMddHHmmss):"
		    "This script requires the EWS Managed API 1.2 or later."
		    "Please download and install the current version of the EWS Managed API from"
		    "http://go.microsoft.com/fwlink/?LinkId=255472"
		    ""
		    "Exiting Script."
		    $exception = New-Object System.Exception ("Managed Api missing")
			throw $exception
		    } 
  
		## Set Exchange Version  
		$ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013 
		  
		## Create Exchange Service Object  
		$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($ExchangeVersion)  
		  
		## Set Credentials to use two options are availible Option1 to use explict credentials or Option 2 use the Default (logged On) credentials  
		  
		#Credentials Option 1 using UPN for the windows Account  
		#$psCred = Get-Credential  
		$creds = New-Object System.Net.NetworkCredential('redemption','redemption')  
		$service.Credentials = $creds      
		#Credentials Option 2  
		#service.UseDefaultCredentials = $true  
		 #$service.TraceEnabled = $true
		## Choose to ignore any SSL Warning issues caused by Self Signed Certificates  
		  
		## Code From http://poshcode.org/624
		## Create a compilation environment
		$Provider=New-Object Microsoft.CSharp.CSharpCodeProvider
		$Compiler=$Provider.CreateCompiler()
		$Params=New-Object System.CodeDom.Compiler.CompilerParameters
		$Params.GenerateExecutable=$False
		$Params.GenerateInMemory=$True
		$Params.IncludeDebugInformation=$False
		$Params.ReferencedAssemblies.Add("System.DLL") | Out-Null

$TASource=@'
  namespace Local.ToolkitExtensions.Net.CertificatePolicy{
    public class TrustAll : System.Net.ICertificatePolicy {
      public TrustAll() { 
      }
      public bool CheckValidationResult(System.Net.ServicePoint sp,
        System.Security.Cryptography.X509Certificates.X509Certificate cert, 
        System.Net.WebRequest req, int problem) {
        return true;
      }
    }
  }
'@ 
		$TAResults=$Provider.CompileAssemblyFromSource($Params,$TASource)
		$TAAssembly=$TAResults.CompiledAssembly

		## We now create an instance of the TrustAll and attach it to the ServicePointManager
		$TrustAll=$TAAssembly.CreateInstance("Local.ToolkitExtensions.Net.CertificatePolicy.TrustAll")
		[System.Net.ServicePointManager]::CertificatePolicy=$TrustAll

		## end code from http://poshcode.org/624
		  
		## Set the URL of the CAS (Client Access Server) to use two options are availbe to use Autodiscover to find the CAS URL or Hardcode the CAS to use  
		  
		#CAS URL Option 1 Autodiscover  
		$service.AutodiscoverUrl($MailboxName,{$true})  
		Write-host ("Using CAS Server : " + $Service.url)   
		   
		#CAS URL Option 2 Hardcoded  
		  
		#$uri=[system.URI] "https://casservername/ews/exchange.asmx"  
		#$service.Url = $uri    
		  
		## Optional section for Exchange Impersonation  
		  
		#$service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName) 
		if(!$service.URL){
			throw "Error connecting to EWS"
		}
		else
		{		
			return $service
		}
	}
}
####################### 
<# 
.SYNOPSIS 
 Creates a Contact in a Contact folder in a Mailbox using the  Exchange Web Services API 
 
.DESCRIPTION 
  Creates a Contact in a Contact folder in a Mailbox using the  Exchange Web Services API 
  
  Requires the EWS Managed API from https://www.microsoft.com/en-us/download/details.aspx?id=42951
.EXAMPLE
	Example 1 To create a contact in the default contacts folder 
	Create-Contact -Mailboxname mailbox@domain.com -EmailAddress contactEmai@domain.com -FirstName John -LastName Doe -DisplayName "John Doe"
	
	Example 2 To create a contact and add a contact picture
	Create-Contact -Mailboxname mailbox@domain.com -EmailAddress contactEmai@domain.com -FirstName John -LastName Doe -DisplayName "John Doe" -photo 'c:\photo\Jdoe.jpg'
	Example 3 To create a contact in a user created subfolder 
	Create-Contact -Mailboxname mailbox@domain.com -EmailAddress contactEmai@domain.com -FirstName John -LastName Doe -DisplayName "John Doe" -Folder "\MyCustomContacts"
    
	This cmdlet uses the EmailAddress as unique key so it wont let you create a contact with that email address if one already exists.
#> 
########################
function Create-Contact 
{ 
    [CmdletBinding()] 
    param( 
    	[Parameter(Position=1, Mandatory=$true)]  [string]$MailboxName,
        [Parameter(Position=2, Mandatory=$true)]  [Xml.XmlElement[]]$ContactObj,
        [Parameter(Position=3, Mandatory=$false)] [string]$Folder,
        [Parameter(Position=4, Mandatory=$false)] [switch]$useImpersonation,
        [Parameter(Position=5, Mandatory=$true)]  [Microsoft.Exchange.WebServices.Data.ExchangeService]$service

		
    )  
 	Begin
	{
		#Connect
		#$service = Connect-Exchange -MailboxName $MailboxName

		if($useImpersonation.IsPresent){
			$service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName)
		}

		$folderid= new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Contacts,$MailboxName)  

        if($Folder){
			$Contacts = Get-ContactFolder -service $service -FolderPath $Folder -SmptAddress $MailboxName
		}
		else{
			$Contacts = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid)
		}
		if($service.URL){
			$type = ("System.Collections.Generic.List"+'`'+"1") -as "Type"
			$type = $type.MakeGenericType("Microsoft.Exchange.WebServices.Data.FolderId" -as "Type")
			$ParentFolderIds = [Activator]::CreateInstance($type)
			$ParentFolderIds.Add($Contacts.Id)
			$Error.Clear();
			$cnpsPropset= new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)  
			$ncCol = $service.ResolveName($ContactObj.Email1,$ParentFolderIds,[Microsoft.Exchange.WebServices.Data.ResolveNameSearchLocation]::DirectoryThenContacts,$true,$cnpsPropset);
			$createContactOkay = $false
			if($Error.Count -eq 0){
				if ($ncCol.Count -eq 0) {
					$createContactOkay = $true;	
				}
				else{
					foreach($Result in $ncCol){
						if($Result.Contact -eq $null){
							Write-host "Contact already exists " + $Result.Mailbox.Name
							throw ("Contact already exists")
						}
						else{
							if((Validate-EmailAddres -EmailAddress $ContactObj.Email1)){
								if($Result.Mailbox.MailboxType -eq [Microsoft.Exchange.WebServices.Data.MailboxType]::Mailbox){
									$UserDn = Get-UserDN -Credentials $Credentials -EmailAddress $Result.Mailbox.Address
									$cnpsPropset= new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties) 
									$ncCola = $service.ResolveName($UserDn,$ParentFolderIds,[Microsoft.Exchange.WebServices.Data.ResolveNameSearchLocation]::ContactsOnly,$true,$cnpsPropset);
									if ($ncCola.Count -eq 0) {  
										$createContactOkay = $true;		
									}
									else
									{
										Write-Host -ForegroundColor  Red ("Number of existing Contacts Found " + $ncCola.Count)
										foreach($Result in $ncCola){
											Write-Host -ForegroundColor  Red ($ncCola.Mailbox.Name)
										}
										throw ("Contact already exists")
									}
								}
							}
							else{
								Write-Host -ForegroundColor Yellow ("Email Address is not valid for GAL match")
							}
						}
					}
				}
				if($createContactOkay){
					$Contact = New-Object Microsoft.Exchange.WebServices.Data.Contact -ArgumentList $service 

                    #   See all the fields here:
                    #   http://msdn.microsoft.com/en-us/library/microsoft.exchange.webservices.data.contact_members.aspx 
                    #   see here for clues
                    #    http://msdn.microsoft.com/en-us/library/aa563318.aspx
                    #   and
                    #    http://msdn.microsoft.com/en-us/library/dd636262.aspx

                    $Contact.GivenName = $ContactObj.GivenName
                    $Contact.Surname = $ContactObj.SurName
                    $Contact.NickName = $ContactObj.NickName
                    $Contact.FileAs = $ContactObj.CompleteName

                    $vorname = $ContactObj.GivenName
                    $nachname = $ContactObj.SurName

                    #Per Regex aus dem CompleteName den Title, weitere Vornamen und Suffix finden
                    $regex = [regex]"^.*(?=$vorname)"
                    $b = $regex.Match($ContactObj.CompleteName) 
                    $Title = $b.Value.Trim()

                    $regex = [regex]"(?<=$vorname).*(?=$nachname)"
                    $b = $regex.Match($ContactObj.CompleteName) 
                    $Contact.MiddleName = $b.Value.Trim()

                    $regex = [regex]"(?<=$nachname).*$"
                    $b = $regex.Match($ContactObj.CompleteName) 
                    $Contact.Generation = $b.Value.Trim()


                    if($Title -ne ""){
                        $PR_DISPLAY_NAME_PREFIX_W = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x3A45,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String);  
                        $Contact.SetExtendedProperty($PR_DISPLAY_NAME_PREFIX_W,$Title)						
                    }

                    $Contact.CompanyName = $ContactObj.CompanyName
                    $Contact.JobTitle = $ContactObj.JobTitle

                    $Contact.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress1] = $ContactObj.Email1
                    $Contact.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress2] = $ContactObj.Email2
                    $Contact.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress3] = $ContactObj.Email3
                    $Contact.BusinessHomePage = $ContactObj.BusinessHomePage

                    $Contact.PhoneNumbers[[Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::AssistantPhone] = $ContactObj.AssistantPhone
                    $Contact.PhoneNumbers[[Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::BusinessFax] = $ContactObj.BusinessFax
                    $Contact.PhoneNumbers[[Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::BusinessPhone] = $ContactObj.BusinessPhone
                    $Contact.PhoneNumbers[[Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::BusinessPhone2] = $ContactObj.BusinessPhone2
                    $Contact.PhoneNumbers[[Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::CompanyMainPhone] = $ContactObj.CompanyMainPhone
                    $Contact.PhoneNumbers[[Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::HomePhone] = $ContactObj.HomePhone
                    $Contact.PhoneNumbers[[Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::MobilePhone] = $ContactObj.MobilePhone
                    $Contact.PhoneNumbers[[Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::PrimaryPhone] = $ContactObj.PrimaryPhone

                    $objNewPhysicalAddress1 = New-Object Microsoft.Exchange.WebServices.Data.PhysicalAddressEntry
                    $objNewPhysicalAddress2 = New-Object Microsoft.Exchange.WebServices.Data.PhysicalAddressEntry

                    # BUSINESS
                    $objNewPhysicalAddress1.Street          = $ContactObj.BusinessStreet
                    $objNewPhysicalAddress1.City            = $ContactObj.BusinessCity
                    $objNewPhysicalAddress1.State           = $ContactObj.BusinessState
                    $objNewPhysicalAddress1.PostalCode      = $ContactObj.BusinessPostalCode
                    $objNewPhysicalAddress1.CountryOrRegion = $ContactObj.BusinessCountryOrRegion

                    # HOME
                    $objNewPhysicalAddress2.Street          = $ContactObj.HomeStreet
                    $objNewPhysicalAddress2.City            = $ContactObj.HomeCity
                    $objNewPhysicalAddress2.State           = $ContactObj.HomeState
                    $objNewPhysicalAddress2.PostalCode      = $ContactObj.HomePostalCode
                    $objNewPhysicalAddress2.CountryOrRegion = $ContactObj.HomeCountryOrRegion

                    # enum! see http://msdn.microsoft.com/en-us/library/exchangewebservices.physicaladdressdictionaryentrytype_members.aspx
                    #           http://msdn.microsoft.com/en-us/library/exchangewebservices.physicaladdresskeytype.aspx
 
                    #[enum]::getvalues([Microsoft.Exchange.WebServices.Data.PhysicalAddressKey])
                    $enumBusinessValue = [Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Business
                    $enumHomevalue     = [Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Home
                    $enumOtherValue    = [Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Other

                    $Contact.PhysicalAddresses[$enumBusinessValue] = $objNewPhysicalAddress1
                    $Contact.PhysicalAddresses[$enumHomevalue]    = $objNewPhysicalAddress2

                    $Contact.Department = $ContactObj.Department
                    $Contact.Profession = $ContactObj.Profession
                    $Contact.Birthday = $ContactObj.Birthday
                    $Contact.WeddingAnniversary = $ContactObj.WeddingAnniversary

                    $Contact.Body = $ContactObj.Note
 
			   		$Contact.Save($Contacts.Id)				
					Write-Host ("Contact Created: " + $ContactObj.CompleteName)
				}
			}
		}
	}
}


####################### 
<# 
.SYNOPSIS 
 Deletes a Contact in a Contact folder in a Mailbox using the  Exchange Web Services API 
 
.DESCRIPTION 
  Deletes a Contact in a Contact folder in a Mailbox using the  Exchange Web Services API 
  
  Requires the EWS Managed API from https://www.microsoft.com/en-us/download/details.aspx?id=42951
.EXAMPLE 
	Example 1 To delete a contact from the default contacts folder
	Delete-Contact -MailboxName mailbox@domain.com -EmailAddress email@domain.com 
	Example2 To delete a contact from a non user subfolder
	Delete-Contact -MailboxName mailbox@domain.com -EmailAddress email@domain.com -Folder \Contacts\Subfolder
#> 
########################
function Delete-Contact 
{

   [CmdletBinding()] 
    param( 
    	[Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
		[Parameter(Position=1, Mandatory=$true)] [string]$EmailAddress,
		[Parameter(Position=2, Mandatory=$true)] [System.Management.Automation.PSCredential]$Credentials,
		[Parameter(Position=3, Mandatory=$false)] [switch]$force,
		[Parameter(Position=4, Mandatory=$false)] [string]$Folder,
		[Parameter(Position=5, Mandatory=$false)] [switch]$Partial
    )  
 	Begin
	{
		#Connect
		$service = Connect-Exchange -MailboxName $MailboxName -Credential $Credentials
		if($useImpersonation.IsPresent){
			$service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName)
		}
		$folderid= new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Contacts,$MailboxName)   
		if($Folder){
			$Contacts = Get-ContactFolder -service $service -FolderPath $Folder -SmptAddress $MailboxName
		}
		else{
			$Contacts = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid)
		}
		if($service.URL){
			$type = ("System.Collections.Generic.List"+'`'+"1") -as "Type"
			$type = $type.MakeGenericType("Microsoft.Exchange.WebServices.Data.FolderId" -as "Type")
			$ParentFolderIds = [Activator]::CreateInstance($type)
			$ParentFolderIds.Add($Contacts.Id)
			$Error.Clear();
			$cnpsPropset= new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties) 
			$ncCol = $service.ResolveName($EmailAddress,$ParentFolderIds,[Microsoft.Exchange.WebServices.Data.ResolveNameSearchLocation]::DirectoryThenContacts,$true,$cnpsPropset);
			if($Error.Count -eq 0){
				if ($ncCol.Count -eq 0) {
					Write-Host -ForegroundColor Yellow ("No Contact Found")		
				}
				else{
					foreach($Result in $ncCol){
						if($Result.Contact -eq $null){
							$contact = [Microsoft.Exchange.WebServices.Data.Contact]::Bind($service,$Result.Mailbox.Id) 
							if($force){
								if(($Result.Mailbox.Address.ToLower() -eq $EmailAddress.ToLower())){
									$contact.Delete([Microsoft.Exchange.WebServices.Data.DeleteMode]::SoftDelete)  
									Write-Host ("Contact Deleted " + $contact.DisplayName + " : Subject-" + $contact.Subject + " : Email-" + $Result.Mailbox.Address)
								}
								else
								{
									Write-Host ("This script won't allow you to force the delete of partial matches")
								}
							}
							else{
								if(($Result.Mailbox.Address.ToLower() -eq $EmailAddress.ToLower()) -bor $Partial.IsPresent){
								    $yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes",""  
		                            $no = New-Object System.Management.Automation.Host.ChoiceDescription "&No","" 
		                           
		                            $choices = [System.Management.Automation.Host.ChoiceDescription[]]($yes,$no)  
		                            $message = "Do you want to Delete contact with DisplayName " + $contact.DisplayName + " : Subject-" + $contact.Subject + " : Email-" + $Result.Mailbox.Address
		                            $result = $Host.UI.PromptForChoice($caption,$message,$choices,1)  
		                            if($result -eq 0) {                       
		                                $contact.Delete([Microsoft.Exchange.WebServices.Data.DeleteMode]::SoftDelete) 
										Write-Host ("Contact Deleted")
		                            } 
									else{
										Write-Host ("No Action Taken")
									}
								}
								
							}
						}
						else{
							if((Validate-EmailAddres -EmailAddress $Result.Mailbox.Address)){
							    if($Result.Mailbox.MailboxType -eq [Microsoft.Exchange.WebServices.Data.MailboxType]::Mailbox){
									$UserDn = Get-UserDN -Credentials $Credentials -EmailAddress $Result.Mailbox.Address
									$cnpsPropset= new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties) 
									$ncCola = $service.ResolveName($UserDn,$ParentFolderIds,[Microsoft.Exchange.WebServices.Data.ResolveNameSearchLocation]::ContactsOnly,$true,$cnpsPropset);
									if ($ncCola.Count -eq 0) {  
										Write-Host -ForegroundColor Yellow ("No Contact Found")			
									}
									else
									{
										Write-Host ("Number of matching Contacts Found " + $ncCola.Count)
										$rtCol = @()
										foreach($aResult in $ncCola){
											if(($aResult.Mailbox.Address.ToLower() -eq $EmailAddress.ToLower()) -bor $Partial.IsPresent){
												$contact = [Microsoft.Exchange.WebServices.Data.Contact]::Bind($service,$aResult.Mailbox.Id) 
												if($force){
													if($aResult.Mailbox.Address.ToLower() -eq $EmailAddress.ToLower()){
														$contact.Delete([Microsoft.Exchange.WebServices.Data.DeleteMode]::SoftDelete)  
														Write-Host ("Contact Deleted " + $contact.DisplayName + " : Subject-" + $contact.Subject + " : Email-" + $Result.Mailbox.Address)
													}
													else
													{
														Write-Host ("This script won't allow you to force the delete of partial matches")
													}
												}
												else{
												    $yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes",""  
						                            $no = New-Object System.Management.Automation.Host.ChoiceDescription "&No","" 
						                            $choices = [System.Management.Automation.Host.ChoiceDescription[]]($yes,$no)  
						                            $message = "Do you want to Delete contact with DisplayName " + $contact.DisplayName + " : Subject-" + $contact.Subject + " : Email-" + $Result.Mailbox.Address 
						                            $result = $Host.UI.PromptForChoice($caption,$message,$choices,1)  
						                            if($result -eq 0) {                       
						                                $contact.Delete([Microsoft.Exchange.WebServices.Data.DeleteMode]::SoftDelete) 
														Write-Host ("Contact Deleted ")
						                            } 
													else{
														Write-Host ("No Action Taken")
													}
													
												}
											}
											else{
												Write-Host ("Skipping Matching because email address doesn't match address on match " + $aResult.Mailbox.Address.ToLower())
											}
										}								
									}
								}
							}
							else
							{
								Write-Host -ForegroundColor Yellow ("Email Address is not valid for GAL match")
							}
						}
					}
				}
			}	
			
		}
	}
}

function Make-UniqueFileName{
    param(
		[Parameter(Position=0, Mandatory=$true)] [string]$FileName
	)
	Begin
	{
	
	$directoryName = [System.IO.Path]::GetDirectoryName($FileName)
    $FileDisplayName = [System.IO.Path]::GetFileNameWithoutExtension($FileName);
    $FileExtension = [System.IO.Path]::GetExtension($FileName);
    for ($i = 1; ; $i++){
            
            if (![System.IO.File]::Exists($FileName)){
				return($FileName)
			}
			else{
					$FileName = [System.IO.Path]::Combine($directoryName, $FileDisplayName + "(" + $i + ")" + $FileExtension);
			}                
            
			if($i -eq 10000){throw "Out of Range"}
        }
	}
}

function Get-ContactFolder{
	param (
	        [Parameter(Position=0, Mandatory=$true)] [string]$FolderPath,
			[Parameter(Position=1, Mandatory=$true)] [string]$SmptAddress,
			[Parameter(Position=2, Mandatory=$true)] [Microsoft.Exchange.WebServices.Data.ExchangeService]$service
		  )
	process{
		## Find and Bind to Folder based on Path  
		#Define the path to search should be seperated with \  
		#Bind to the MSGFolder Root  
		$folderid = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot,$SmptAddress)   
		$tfTargetFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid)  
		#Split the Search path into an array  
		$fldArray = $FolderPath.Split("\") 
		 #Loop through the Split Array and do a Search for each level of folder 
		for ($lint = 1; $lint -lt $fldArray.Length; $lint++) { 
            $folderName = $fldArray[$lint]
	        #Perform search based on the displayname of each folder level 
	        $fvFolderView = new-object Microsoft.Exchange.WebServices.Data.FolderView(1) 
	        $SfSearchFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName,$fldArray[$lint]) 
	        $findFolderResults = $service.FindFolders($tfTargetFolder.Id,$SfSearchFilter,$fvFolderView) 
	        if ($findFolderResults.TotalCount -gt 0){ 
	            foreach($folder in $findFolderResults.Folders){ 
	                $tfTargetFolder = $folder                
	            } 
	        } 
	        else{ 
	            $newFolder = New-Object Microsoft.Exchange.WebServices.Data.Folder($service)
                $newFolder.DisplayName = $folderName
                $newFolder.FolderClass = "IPF.Contact"
                $newFolder.Save($tfTargetFolder.Id)
	            $tfTargetFolder = $newFolder  
	        }     
	    }  
		if($tfTargetFolder -ne $null){
			return [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$tfTargetFolder.Id)
		}
		else{
			throw ("Folder Not found")
		}
	}
}


####################### 
<# 
.SYNOPSIS 
 Creates Contacts from a XML in a Contact folder in a Mailbox using the  Exchange Web Services API 
 
.DESCRIPTION 
  Creates Contacts from a XML in a Contact folder in a Mailbox using the  Exchange Web Services API 
  
  Requires the EWS Managed API from https://www.microsoft.com/en-us/download/details.aspx?id=42951
.EXAMPLE
	Example 1 To create a contact in the default contacts folder 
	Create-Contacts-from-XML -MailboxName walzenbach@arges.de -Folder "\Kontakte\TestKontakte" -XMLPath "D:\ARGESContactList.xml"
    
	This cmdlet uses the EmailAddress as unique key so it wont let you create a contact with that email address if one already exists.
#> 
########################
function Create-Contacts-from-XML 
{
    [CmdletBinding()] 
    param( 
    	[Parameter(Position=1, Mandatory=$true)]  [string]$MailboxName,
        [Parameter(Position=2, Mandatory=$true)]  [string]$XMLPath,
        [Parameter(Position=3, Mandatory=$false)] [string]$Folder
    )  
 	Begin
	{
        #Connect
		$service = Connect-Exchange -MailboxName $MailboxName


        # XML Hardcoded
        #[xml]$xml = get-content "D:\ARGESContactList.xml"
        # XML per Parameter
        [xml]$xml = get-content $XMLPath

        $contacts = $xml.Report.Tablix1.Details_Collection
        $contactArr = New-Object System.Collections.Generic.List[System.Object]

        foreach ($contact in $contacts.ChildNodes)
        {
            Create-Contact -MailboxName $MailboxName -useImpersonation -Folder $Folder -ContactObj $contact -service $service
        }
    }
}




##### TESTAUFRUF
Create-Contacts-from-XML -MailboxName walzenbach@arges.de -Folder "\Kontakte\TestKontakte" -XMLPath "D:\ARGESContactList.xml"
