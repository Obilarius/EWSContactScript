[xml]$xml = get-content "D:\ARGESContactList.xml"

$contacts = $xml.Report.Tablix1.Details_Collection
$contactArr = New-Object System.Collections.Generic.List[System.Object]

foreach ($contact in $contacts.ChildNodes)
{
    if($contact.GivenName -eq "Vorname") {
        #Create-Contact -MailboxName walzenbach@arges.de -ContactObj $contact
        $contactout = New-Object Microsoft.Exchange.WebServices.Data.Contact 
        Write-Output $contactout
    }
}




function Connect-Exchange
{ 
    param( 
    	[Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
		[Parameter(Position=1, Mandatory=$true)] [System.Management.Automation.PSCredential]$Credentials
    )  
 	Begin
		 {
		## Load Managed API dll  
		###CHECK FOR EWS MANAGED API, IF PRESENT IMPORT THE HIGHEST VERSION EWS DLL, ELSE EXIT
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
		$ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013_SP1  
		  
		## Create Exchange Service Object  
		$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($ExchangeVersion)  
		  
		## Set Credentials to use two options are availible Option1 to use explict credentials or Option 2 use the Default (logged On) credentials  
		  
		#Credentials Option 1 using UPN for the windows Account  
		#$psCred = Get-Credential  
		$creds = New-Object System.Net.NetworkCredential($Credentials.UserName.ToString(),$Credentials.GetNetworkCredential().password.ToString())  
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
		#$service.AutodiscoverUrl($MailboxName,{$true})  
		#Write-host ("Using CAS Server : " + $Service.url)   
		   
		#CAS URL Option 2 Hardcoded  
		  
		$uri=[system.URI] "https://helios.arges.local/ews/exchange.asmx"  
		$service.Url = $uri    
		  
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

function Create-Contact 
{ 
    [CmdletBinding()] 
    param( 
    	[Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
		[Parameter(Position=6, Mandatory=$true)] [System.Management.Automation.PSCredential]$Credentials,
		[Parameter(Position=7, Mandatory=$false)] [object]$ContactObj,
		[Parameter(Position=26, Mandatory=$false)] [switch]$useImpersonation
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
            $EmailAddress = $ContactObj.Email1
			$ncCol = $service.ResolveName($EmailAddress,$ParentFolderIds,[Microsoft.Exchange.WebServices.Data.ResolveNameSearchLocation]::DirectoryThenContacts,$true,$cnpsPropset);
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
							if((Validate-EmailAddres -EmailAddress $EmailAddress)){
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
					$Contact = New-Object Microsoft.Exchange.WebServices.Data.Contact($service)

					
					$Contact.GivenName = $ContactObj.GivenName
					$Contact.Surname = $ContactObj.SurName
					#if($ContactObj.CompleteName) { $Contact.CompleteName = $ContactObj.CompleteName }
					$Contact.FileAs = $ContactObj.CompleteName
					if($Title -ne ""){
						$PR_DISPLAY_NAME_PREFIX_W = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x3A45,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String);  
						$Contact.SetExtendedProperty($PR_DISPLAY_NAME_PREFIX_W,$Title)						
					}
					if($ContactObj.CompanyName) { $Contact.CompanyName = $ContactObj.CompanyName }
					if($ContactObj.DisplayName) { $Contact.DisplayName = $ContactObj.DisplayName }
					if($ContactObj.Department) { $Contact.Department = $ContactObj.Department }
					if($ContactObj.Office) { $Contact.OfficeLocation = $ContactObj.Office }
					if($ContactObj.CompanyName) { $Contact.CompanyName = $ContactObj.CompanyName }
					if($ContactObj.BusinssPhone) { $Contact.PhoneNumbers[[Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::BusinessPhone] = $ContactObj.BusinssPhone }
					if($ContactObj.MobilePhone) { $Contact.PhoneNumbers[[Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::MobilePhone] = $ContactObj.MobilePhone }
					if($ContactObj.HomePhone) { $Contact.PhoneNumbers[[Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::HomePhone] = $ContactObj.HomePhone }
                    
					$Contact.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Business] = New-Object  Microsoft.Exchange.WebServices.Data.PhysicalAddressEntry
					if($ContactObj.Street) { $Contact.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Business].Street = $ContactObj.Street }
					if($ContactObj.State) { $Contact.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Business].State = $ContactObj.State }
					if($ContactObj.City) { $Contact.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Business].City = $ContactObj.City }
					if($ContactObj.Country) { $Contact.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Business].CountryOrRegion = $ContactObj.Country }
					if($ContactObj.PostalCode) { $Contact.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Business].PostalCode = $ContactObj.PostalCode }

					$Contact.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress1] = $ContactObj.EmailAddress
					if($ContactObj.FileAs) { $Contact.FileAs = $ContactObj.FileAs }
					if($ContactObj.Notes) { $Contact.Body = $ContactObj.Notes }


					if($Photo){
						$fileAttach = $Contact.Attachments.AddFileAttachment($Photo)
						$fileAttach.IsContactPhoto = $true
					}


			   		$Contact.Save($Contacts.Id)				
					Write-Host ("Contact Created")
				}
			}
		}
	}
}