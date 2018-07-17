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
    #if($ContactObj.GivenName -eq "Vorname") {
        $Contact = New-Object Microsoft.Exchange.WebServices.Data.Contact($service)

        $Contact.GivenName = $ContactObj.GivenName
        $Contact.Surname = $ContactObj.SurName
        $Contact.FileAs = $ContactObj.CompleteName
        #if($Title -ne ""){
	    #    $PR_DISPLAY_NAME_PREFIX_W = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x3A45,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String);  
	    #    $Contact.SetExtendedProperty($PR_DISPLAY_NAME_PREFIX_W,$Title)						
        #}
        if($ContactObj.CompanyName) { $Contact.CompanyName = $ContactObj.CompanyName }
        if($ContactObj.BusinssPhone) { $Contact.PhoneNumbers[[Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::BusinessPhone] = $ContactObj.BusinssPhone }
        if($ContactObj.MobilePhone) { $Contact.PhoneNumbers[[Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::MobilePhone] = $ContactObj.MobilePhone }
        if($ContactObj.HomePhone) { $Contact.PhoneNumbers[[Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::HomePhone] = $ContactObj.HomePhone }
                    
        $Contact.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Business] = New-Object  Microsoft.Exchange.WebServices.Data.PhysicalAddressEntry
        if($ContactObj.BusinessStreet) { $Contact.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Business].Street = $ContactObj.BusinessStreet }
        if($ContactObj.BusinessState) { $Contact.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Business].State = $ContactObj.BusinessState }
        if($ContactObj.BusinessCity) { $Contact.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Business].City = $ContactObj.BusinessCity }
        if($ContactObj.BusinessCountryOrRegion) { $Contact.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Business].CountryOrRegion = $ContactObj.BusinessCountryOrRegion }
        if($ContactObj.BusinessPostalCode) { $Contact.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Business].PostalCode = $ContactObj.BusinessPostalCode }

        $Contact.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress1] = $ContactObj.Email1
        if($ContactObj.Notes) { $Contact.Body = $ContactObj.Notes }

        $Contact.Save($Contacts.Id)				
		Write-Host ($ContactObj.CompleteName + " Contact Created")
    #}
}
