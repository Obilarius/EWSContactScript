function EmptyFunction
{

   [CmdletBinding()] 
    param( 
		[Parameter(Position=0, Mandatory=$true)]  [string]$FolderName,
        [Parameter(Position=1, Mandatory=$true)]  [Microsoft.Exchange.WebServices.Data.ExchangeService]$service
    )  
 	Begin
	{
        
    }
}


function getEWSConnect
{

   [CmdletBinding()] 
    param( 
		[Parameter(Position=0, Mandatory=$true)]  [string]$MailboxName
    )  
 	Begin
	{
        ##Define the SMTP Address of the mailbox to impersonate
        $MailboxToImpersonate = $MailboxName

        ## Load Exchange web services DLL
        ## Download here if not present: http://go.microsoft.com/fwlink/?LinkId=255472
        $dllpath = "C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll"
        Import-Module $dllpath

        ## Set Exchange Version
        $ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013

        ## Create Exchange Service Object
        $service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($ExchangeVersion)

        #Get valid Credentials using UPN for the ID that is used to impersonate mailbox
        $creds = New-Object System.Net.NetworkCredential('redemption','redemption')
        $service.Credentials = $creds

        ## Set the URL of the CAS (Client Access Server)
        #$service.AutodiscoverUrl($AccountWithImpersonationRights ,{$true})
        $service.Url = New-Object URI("https://helios.arges.local/EWS/Exchange.asmx")


        ##Login to Mailbox with Impersonation
        $service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress,$MailboxToImpersonate );

        return $service
    }
}


function getPublicFolderId
{

   [CmdletBinding()] 
    param( 
		[Parameter(Position=0, Mandatory=$true)]  [string]$FolderName,
        [Parameter(Position=1, Mandatory=$true)]  [Microsoft.Exchange.WebServices.Data.ExchangeService]$service
    )  
 	Begin
	{
        ## PUBLIC FOLDER
        $itemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView(1000)
        $PublicFolderRoot = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::PublicFoldersRoot)
        $PublicFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$PublicFolderRoot)

        ## finde den Public Contact Folder "Arges Intern"
        $folderView = New-Object Microsoft.Exchange.WebServices.Data.FolderView(1)
        $searchFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName, $FolderName)
        $findPublicContactFolder = $service.FindFolders($PublicFolder.Id, $searchFilter, $folderView)
        $PublicContactsFolder = $findPublicContactFolder[0]

        return $PublicContactsFolder.Id
    }
}


function getPublicFolderItems
{

   [CmdletBinding()] 
    param( 
		[Parameter(Position=0, Mandatory=$true)]  [Microsoft.Exchange.WebServices.Data.FolderId]$FolderId
    )  
 	Begin
	{
        ## PUBLIC FOLDER
        $itemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView(1000)

        $PublicFindItemResults = $service.FindItems($FolderId, $itemView) 

        return $PublicFindItemResults
    }
}


function getLocalFolderID
{

   [CmdletBinding()] 
    param( 
		[Parameter(Position=0, Mandatory=$true)]  [string]$FolderName,
        [Parameter(Position=1, Mandatory=$true)]  [Microsoft.Exchange.WebServices.Data.ExchangeService]$service
    )  
 	Begin
	{
        ## LOCAL FOLDER
        $itemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView(1000)
        $mbRootFolderId = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot)
        $mbFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service, $mbRootFolderId)

        $destination = $FolderName
        $arrMbPath = $destination.Split("\")
        for ($i = 1; $i -lt $arrMbPath.length; $i++)
        {
          $folderName = $arrMbPath[$i]
          $mbFolderView = New-Object Microsoft.Exchange.WebServices.Data.FolderView(1)
          $mbSearchFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName, $arrMbPath[$i])
          $mbFindFolderResults = $service.FindFolders($mbFolder.Id, $mbSearchFilter, $mbFolderView)
          if ($mbFindFolderResults.TotalCount -gt 0)
          {
            $mbFolder = $mbFindFolderResults.Folders[0]
          }
          else
          {
            $newFolder = New-Object Microsoft.Exchange.WebServices.Data.Folder($service)
            $newFolder.DisplayName = $folderName
            $newFolder.FolderClass = "IPF.Contact"
            $newFolder.Save($mbFolder.Id)
            $mbFolder = $newFolder
          }
        }
        $LocalContactsFolderId = $mbFolder.Id

        return $LocalContactsFolderId
    }
}


function getLocalFolderItems
{

   [CmdletBinding()] 
    param( 
		[Parameter(Position=0, Mandatory=$true)]  [Microsoft.Exchange.WebServices.Data.FolderId]$FolderID,
        [Parameter(Position=1, Mandatory=$true)]  [Microsoft.Exchange.WebServices.Data.ExchangeService]$service
    )  
 	Begin
	{
        # Define Property Set 
        $itemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView(1000)
        $Propset = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties) 

        #$PidTagSubject = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x804F,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String);  
        #$Propset.Add($PidTagSubject)  

        # ExtendedPropertys
        $AddressGuid2 = new-object Guid(“00062004-0000-0000-C000-000000000046”)
        $User3 = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition($AddressGuid2,0x8051,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String);
        $User4 = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition($AddressGuid2,0x8052,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String);  
        $Propset.add($User3) 
        $Propset.add($User4) 
        
        #### TEST
        #$cmpName = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition([Microsoft.Exchange.WebServices.Data.DefaultExtendedPropertySet]::PublicStrings,"User1",[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String);
        #$cmpName2 = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition([Microsoft.Exchange.WebServices.Data.DefaultExtendedPropertySet]::PublicStrings,"Testfeld",[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String);
        #$cmpName3 = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition([Microsoft.Exchange.WebServices.Data.DefaultExtendedPropertySet]::PublicStrings,"testfeldid",[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String);
        
        #$Propset.add($cmpName)
        #$Propset.add($cmpName2)
        #$Propset.add($cmpName3)
        ####


        $itemView.PropertySet = $Propset 

        $LocalFindItemResults = $service.FindItems($FolderID, $itemView)

        #try {
	    #    [Void]$service.LoadPropertiesForItems($LocalFindItemResults,$Propset)
        #}
        #catch {
	    #    write-warning "SKIP:Unable to load User1"
        #}

        return $LocalFindItemResults
    }
}


#noch nicht auf Funktion umgebaut
function copyContact
{

   [CmdletBinding()] 
    param( 
		[Parameter(Position=0, Mandatory=$true)]  [Microsoft.Exchange.WebServices.Data.Item]$pItem,
        [Parameter(Position=1, Mandatory=$true)]  [Microsoft.Exchange.WebServices.Data.ExchangeService]$service,
        [Parameter(Position=2, Mandatory=$true)]  [Microsoft.Exchange.WebServices.Data.FolderId]$localFolderID
    )  
 	Begin
	{
        $pItem.copy($localFolderID)
        Write-Host "Kontakt kopiert: " $pItem.DisplayName

        # Suche Localen Kontakt der noch keine Id in Feld User3, User4 hat
        $localItems = getLocalFolderItems -FolderID $localFolderID -service $service
        foreach ($lItem in $localItems) {
            # Prüft ob User3 und User4 befüllt sind
            $AddressGuid2 = new-object Guid(“00062004-0000-0000-C000-000000000046”)
            $User3 = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition($AddressGuid2,0x8051,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String);
            $User4 = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition($AddressGuid2,0x8052,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String);

            $exPropUser3 = $null;
            $exPropUser4 = $null;
            $bool = $lItem.TryGetProperty($User3, [ref]$exPropUser3)
            $bool = $lItem.TryGetProperty($User4, [ref]$exPropUser4)

            # Prüfung ob User Felder befüllt sind
            if (!$exPropUser3 -or !$exPropUser4) {

                #Überprüft nochmal sicherheitshalber ob EmailAdressen und Display Name übereinstimmt
                if ($lItem.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress1].Address -eq $pItem.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress1].Address -and 
                    $lItem.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress2].Address -eq $pItem.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress2].Address -and 
                    $lItem.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress3].Address -eq $pItem.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress3].Address -and 
                    $lItem.DisplayName -eq $pItem.DisplayName) {

                    $lItem.SetExtendedProperty($User3, $pItem.id.UniqueId)
                    $lItem.SetExtendedProperty($User4, $pItem.id.ChangeKey)
                    $lItem.Update([Microsoft.Exchange.WebServices.Data.ConflictResolutionMode]::AlwaysOverwrite)

                    Write-Host "Kontakt ID eingetragen: " $lItem.DisplayName
                    Write-Host $lItem.id
                    Write-Host $pItem.Id
                }
            }

        } 
    }
}


function getLocalHashtable
{

   [CmdletBinding()] 
    param( 
		[Parameter(Position=0, Mandatory=$false)] [System.Array]$localItems = @()
    )  
 	Begin
	{
        $localHashtable = @{}
        foreach ($item in $localItems)
        {
            $AddressGuid2 = new-object Guid(“00062004-0000-0000-C000-000000000046”)
            $User3 = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition($AddressGuid2,0x8051,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String);
            $User4 = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition($AddressGuid2,0x8052,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String);

            $exPropUser3 = $null;
            $exPropUser4 = $null;
            $bool = $item.TryGetProperty($User3, [ref]$exPropUser3)
            $bool = $item.TryGetProperty($User4, [ref]$exPropUser4)

            $hashtable = @{id = $item.id.UniqueId; ChangeKey = $item.Id.ChangeKey; publicId = $exPropUser3; publicChangeKey = $exPropUser4}

            try {
                $localHashtable[$exPropUser3] = $hashtable
            } catch {}
        }
        return $localHashtable
    }
}


function getPublicHashtable
{

   [CmdletBinding()] 
    param( 
		[Parameter(Position=0, Mandatory=$false)]  [System.Array]$publicItems = @()
    )  
 	Begin
	{
        $publicHashtable = @{}
        foreach ($pItem in $publicItems)
        {

            $hashtable = @{id = $pItem.id.UniqueId; ChangeKey = $pItem.Id.ChangeKey}
            $publicHashtable[$pItem.id.UniqueId] = $hashtable
        
        }
        return $publicHashtable
    }
}


Function Get-StringHash([String] $String, $HashName = "MD5") 
{ 
    $StringBuilder = New-Object System.Text.StringBuilder 
    [System.Security.Cryptography.HashAlgorithm]::Create($HashName).ComputeHash([System.Text.Encoding]::UTF8.GetBytes($String))|%{ 
    [Void]$StringBuilder.Append($_.ToString("x2")) 
    } 
    $StringBuilder.ToString() 
}



$service = getEWSConnect -MailboxName "walzenbach@arges.de"

$localFolderID = getLocalFolderID -FolderName '\Kontakte\Arges Intern' -service $service
$localItems = getLocalFolderItems -FolderID $localFolderID -service $service
$localHashes = getLocalHashtable($localItems)

$localHashes.Count

$publicFolderID = getPublicFolderId -FolderName 'Arges Intern' -service $service
$publicItems = getPublicFolderItems -FolderId $publicFolderID
$publicHashes = getPublicHashtable($publicItems)



#$localHashes["AAIARgAAAAAAGkRzkKpmEc2byACqAC/EWgkAnAJFgiSNqUmRkyDbF6atEgAAAAK0OAAAAnqhhWM4XEa7wXcwcaolGgAAMGJcGQAALgAAAAAAGkRzkKpmEc2byACqAC/EWgMAnAJFgiSNqUmRkyDbF6atEgAAAAK0OAAA"]

foreach($id in $publicHashes.Keys){ 

    if($id -eq "AAIARgAAAAAAGkRzkKpmEc2byACqAC/EWgkAnAJFgiSNqUmRkyDbF6atEgAAAAK0OAAAAnqhhWM4XEa7wXcwcaolGgAAMGJcGQAALgAAAAAAGkRzkKpmEc2byACqAC/EWgMAnAJFgiSNqUmRkyDbF6atEgAAAAK0OAAA") {
        
        
        #$localHashes[$id]


        #$TargetFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$publicFolderID)
        #$objContact = [Microsoft.Exchange.WebServices.Data.Contact]::Bind($service,$id)
        #Write-Host $objContact.DisplayName "  -  " $objContact.Id

        #copyContact -pItem $objContact -service $service -localFolderID $localFolderID
        
        #$TargetFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$localFolderID)
        #$objContact2 = [Microsoft.Exchange.WebServices.Data.Contact]::Bind($service,$localHashes[$id].id)
        #Write-Host $objContact2.DisplayName "  -  " $objContact2.Id

    }







    # Vorhanden in Local
    if($localHashes.ContainsKey($id)) {

        # Vergleiche ChangeKey
        if($publicHashes[$key].ChangeKey -ne $localHashes[$key].publicChangeKey) {
            # Kontakt hat sich geändert
            # TODO: Kontakt löschen und neu erstellen

            $TargetFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$localFolderID)
            $objContact = [Microsoft.Exchange.WebServices.Data.Contact]::Bind($service,$localHashes[$id].id)

        }

    } else {

        # Kontakt nicht vorhanden
        # TODO: Kontakt erstellen
        
        $TargetFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$publicFolderID)
        $objContact = [Microsoft.Exchange.WebServices.Data.Contact]::Bind($service,$id)
        Write-Host $objContact.DisplayName "  -  " $objContact.Id
        #copyContact -pItem $objContact -service $service -localFolderID $localFolderID
    }

}

