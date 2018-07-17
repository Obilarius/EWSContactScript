$MailboxName = "walzenbach@arges.de"


Import-Module "C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll"
$ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013_SP1  
$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($ExchangeVersion)

$uri=[system.URI] "https://helios.arges.local/ews/exchange.asmx"  
$service.Url = $uri 


$folderid = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Contacts,$MailboxName)
$Contacts = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid)



### XML
[xml]$xml = get-content "D:\ARGESContactList.xml"

$xml_contacts = $xml.Report.Tablix1.Details_Collection
$xml_contactArr = New-Object System.Collections.Generic.List[System.Object]





foreach ($ContactObj in $xml_contacts.ChildNodes)
{
    $Contact = New-Object Microsoft.Exchange.WebServices.Data.Contact($service)

    
    ### Scan Complete Name from XML to Title, MiddleName and Suffix
    $regex = "^(.*) " + $ContactObj.GivenName
    $Title = $ContactObj.CompleteName -Split $regex
    $Title = $CompleteName[1]
        
    $regex = $ContactObj.GivenName + " (.*) " + $ContactObj.SurName
    $MiddleName = $ContactObj.CompleteName -Split $regex
    $MiddleName = $MiddleName[1]

    $regex = $ContactObj.SurName + " (.*)"
    $Suffix = $ContactObj.CompleteName -Split $regex
    $Suffix = $Suffix[1]
        

    $Contact.GivenName = $ContactObj.GivenName
    $Contact.Surname = $ContactObj.SurName
    $Contact.NickName = $ContactObj.NickName
    $Contact.FileAs = $ContactObj.CompleteName
    if($Title -ne ""){
	    $PR_DISPLAY_NAME_PREFIX_W = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x3A45,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String);  
	    $Contact.SetExtendedProperty($PR_DISPLAY_NAME_PREFIX_W,$Title)						
    }
    $Contact.MiddleName = $MiddleName
    $Contact.Generation = $Suffix

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
    $Contact.PhoneNumbers[[Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::HomePhone2] = $ContactObj.HomePhone2
    $Contact.PhoneNumbers[[Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::MobilePhone] = $ContactObj.MobilePhone
    $Contact.PhoneNumbers[[Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::PrimaryPhone] = $ContactObj.PrimaryPhone

    $Contact.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Business].Street = $ContactObj.BusinessStreet
    $Contact.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Business].City = $ContactObj.BusinessCity
    $Contact.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Business].CountryOrRegion = $ContactObj.BusinessCountryOrRegion
    $Contact.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Business].PostalCode = $ContactObj.BusinessPostalCode
    $Contact.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Business].State = $ContactObj.BusinessState
 
    $Contact.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Home].Street = $ContactObj.HomeStreet
    $Contact.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Home].City = $ContactObj.HomeCity
    $Contact.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Home].CountryOrRegion = $ContactObj.HomeCountryOrRegion
    $Contact.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Home].PostalCode = $ContactObj.HomePostalCode
    $Contact.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Home].State = $ContactObj.HomeState

    $Contact.Department = $ContactObj.Department
    $Contact.Profession = $ContactObj.Profession
    $Contact.Birthday = $ContactObj.Birthday
    $Contact.WeddingAnniversary = $ContactObj.WeddingAnniversary

    $Contact.Body = $ContactObj.Note
    
    $Contact.Save($Contacts.Id)				
	Write-Host ($ContactObj.CompleteName + " Contact Created")
}
