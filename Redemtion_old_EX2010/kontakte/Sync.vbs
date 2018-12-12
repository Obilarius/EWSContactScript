Const LogFile="C:\Program Files (x86)\Redemption\Kontakte\SyncLog.txt"
Const MailBoxFile ="C:\Program Files (x86)\Redemption\Kontakte\MailboxesToSync.txt"

Call WriteToLog ("***************** Start Sync Run ***************************************************")
Call RunSyncForMailboxes(MailBoxFile)
Call WriteToLog ("***************** Done with Sync Run ***************************************************")

'*****************************************************************************
Private Sub RunSyncForMailboxes(txtPath)
' Read txt file with mailboxes and PF_Paths to sync
Dim arrFileLines()

i = 0
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.OpenTextFile(txtPath, 1)
Do Until objFile.AtEndOfStream
     Redim Preserve arrFileLines(i)
     arrFileLines(i) = objFile.ReadLine
     i = i + 1
Loop
objFile.Close


For l = Lbound(arrFileLines)+1 to UBound(arrFileLines)
    'Wscript.Echo arrFileLines(l)
     SyncParas=split(arrFileLines(l),";")
    'Wscript.Echo SyncParas(0)
    Call DoSync (SyncParas(0),SyncParas(1),SyncParas(2))
Next
end sub

'*****************************************************************************
Sub WriteToLog (LogEntry)
Const ForAppending = 8

Set objFSO = CreateObject("Scripting.FileSystemObject")

If objFSO.FileExists(LogFile) then
	Set objFile = objFSO.GetFile(LogFile)
	'Wscript.Echo objFile.Size

'	If objFile.Size > 10000000 then
'  	
'		objFile.delete
'		Delete if logfile larger 1MB
	If objFile.Size > 1000000  then
'		Delete old Synclog.txt.old
		Set fso = CreateObject("Scripting.FileSystemObject")
		fso.GetFile("C:\Program Files (x86)\Redemption\Kontakte\SyncLog.txt.old").Delete
'		Rename 	SyncLog.txt to SyncLog.txt.old
		objFile.Move "C:\Program Files (x86)\Redemption\Kontakte\SyncLog.txt.old"
	end if
end if

Set objTextFile = objFSO.OpenTextFile(LogFile, ForAppending, True)



objTextFile.WriteLine(now() & " - " & LogEntry)
'Wscript.Echo "Test"
objTextFile.Close

end sub

'*****************************************************************************
Sub DoSync (mailbox, mailserver, ExchangePFPath)

MAPI_NO_CACHE = &H200
MAPI_BEST_ACCESS = &H10

Dim Session
Dim Session2
Dim ExchangeStore
Dim PFStore
Dim ArgesKontakte
Dim PFFolder
Dim ArgesKontakt
Dim Synchonizer

Dim CopyDrop
Dim ProfileContact
Dim SyncDataContact
Dim Prop
Dim SynchFolderEntryID
Dim i
Dim syncdata_available
Dim EntryIDPF()
Dim EntryIDContacts()

Set Session = CreateObject("Redemption.RDOSession")
Session.LogonExchangeMailbox mailbox, mailserver
'WriteToLog "angemeldet"
'Wscript.Echo Session.Stores.Count
'Wscript.Echo Session.loggedon
Set ProfileContacts = Session.GetDefaultFolder(10)
'Wscript.Echo ProfileContacts.name

Set PFStore = Session.Stores.FindExchangePublicFoldersStore
a = PFStore.name
Set ArgesKontakte = Session.GetFolderFromID(ExchangePFPath)
'Wscript.Echo ArgesKontakte.name
'Wscript.Echo ArgesKontakte.FolderPath


'Set ExchangeStore = Session.Stores.DefaultStore
''Wscript.Echo ExchangeStore.name


'strSQLRestriction = " MessageClass = 'IPM.Contact' "

SynchFolderEntryID = ArgesKontakte.EntryID
Set Synchonizer = ArgesKontakte.ExchangeSynchonizer

WriteToLog ("*** Start Sync for mailbox " & mailbox & " sync folder "& ExchangePFPath &" with " & ProfileContacts & " ***")


i = -1

'check for existing Sync Data
Set SyncDataContact=Nothing

If Not ProfileContacts.Items(SynchFolderEntryID) Is Nothing Then
        Set SyncDataContact = ProfileContacts.Items(SynchFolderEntryID)
        strPreviousSyncData = SyncDataContact.Body
        syncdata_available = True
Else
		
	Set Item = ProfileContacts.Items.Find("Select LastName FROM Folder WHERE FullName = 'no not remove - Sync Hash'")
	
	If Item is nothing then
		''Wscript.Echo ProfileContacts.Items.Count
		deleted=True
		Do While deleted=True
			deleted=False
			For each Kontakt in ProfileContacts.Items
				If Kontakt.Sensitivity <> 2 then
					WriteToLog (j &" Deleting " & Kontakt & " missing private flag " & Kontakt.Sensitivity)
 					Kontakt.delete
 					deleted=True
				end if
			
			Next
		loop		
	end if
        
End If

''Wscript.Echo ProfileContacts.Items.Count

If syncdata_available = False Then strPreviousSyncData = "" 'no data at first run, this really needs to come from some persistent storage saved after the previous sync
Set SyncItems = Synchonizer.SyncItems(strPreviousSyncData)



If SyncItems.Count = 0 Then
     WriteToLog ("There were no changes in the folder " & ExchangePFPath & " for mailbox: " & mailbox)
Else

    WriteToLog ("There were " & SyncItems.Count & " changes in the folder. The list of changes follows:")
    
    For Each ProfileContact In ProfileContacts.Items
        Set Prop = ProfileContact.UserProperties.Find("EntryID_PF")
        If Not Prop Is Nothing Then
          i = i + 1
          ReDim Preserve EntryIDPF(i)
          EntryIDPF(i) = Prop.Value
          ReDim Preserve EntryIDContacts(i)
          EntryIDContacts(i) = ProfileContact.EntryID
          
        End If
    Next
    
   If  ProfileContacts.Folders.Item("CopyDrop") Is Nothing Then
			Set CopyDrop = Nothing
		Else
			Set CopyDrop = ProfileContacts.Folders.Item("CopyDrop") 
		End If		
    
    For Each Item In SyncItems
    	 
        If Item.Kind = 0 Then 'sikChanged
            'modification
            
            If Item.IsNewMessage Then
              WriteToLog ("New: " & Item.Item.Subject)
              
              Set ArgesKontakt = Item.Item
              'Copy Folder
              If CopyDrop Is Nothing Then
                Set CopyDrop = ProfileContacts.Folders.Add("CopyDrop")
              End If
              ArgesKontakt.CopyTo CopyDrop
              
              Set ProfileContact = CopyDrop.Items.GetLast
              Set Prop = ProfileContact.UserProperties.Add("EntryID_PF", 1)
              Prop.Value = Item.EntryID
              ProfileContact.Save
              ProfileContact.Move ProfileContacts
              
                 
            Else
              WriteToLog ("Modified: " & Item.Item.Subject)
              
              Set ArgesKontakt = Item.Item
              For i = 0 To UBound(EntryIDPF)
                If EntryIDPF(i) = ArgesKontakt.EntryID Then Exit For
              Next
              
              Set ProfileContact = Session.GetMessageFromID(EntryIDContacts(i))
              ProfileContact.Delete
                  
              'Copy Folder
              If CopyDrop Is Nothing Then
                Set CopyDrop = ProfileContacts.Folders.Add("CopyDrop")
              End If
              ArgesKontakt.CopyTo CopyDrop
              
              Set ProfileContact = CopyDrop.Items.GetLast
              Set Prop = ProfileContact.UserProperties.Add("EntryID_PF", 1)
              Prop.Value = CStr(Item.EntryID)
              ProfileContact.Save
              ProfileContact.Move ProfileContacts

            End If
        ElseIf Item.Kind = 1 Then 'sikDeleted		    
            'deletion. Since the Item is gone, RDOSyncMessageItem.Item will be NULL
			
			' Print entryId that shall be deleted
			'Wscript.echo Item.EntryID
			
            WriteToLog ("Deletion. Source key = " & Item.SourceKey)
			
			'Search in ProfileContacts for deleted item by finding the user property EntryID_PF
			For each Kontakt in ProfileContacts.Items			
			  Set prop = Kontakt.UserProperties.Find("EntryID_PF")
			  'Do not care if there is no EntryID_PF user property
			  If Not Prop Is Nothing Then
			    'Found matching EntryID
			    If Prop.Value = Item.EntryID then
				  ' Print name (Subject) of contact
				  'wscript.echo Kontakt.Subject
				  dim name 
				  name = Kontakt.Subject
				  Kontakt.delete
				  WriteToLog ("Deletion of " & name & " sucessfull!")				  
			    end if
			  end if
			next		
        End If
   Next


strPreviousSyncData = SyncItems.SyncData

If SyncDataContact Is Nothing Then
    Set SyncDataContact = ProfileContacts.Items.Add
End If

SyncDataContact.FullName = "no not remove - Sync Hash"
SyncDataContact.FileAs = "zzz_no not remove - Sync Hash " & ArgesKontakte.Name
SyncDataContact.Subject = SynchFolderEntryID
SyncDataContact.Body = strPreviousSyncData
SyncDataContact.Save

' strPreviousSyncData now needs to be persisted to be used in the next sync instead of using an empty string
If Not CopyDrop Is Nothing Then
 ProfileContacts.Folders("CopyDrop").Delete
End If


End If
WriteToLog ("*** Done with sync for mailbox: "& mailbox & " for folder :" & ExchangePFPath & " ***")

' Free all RDO objects
Set ExchangeStore = Nothing
Set PFStore = Nothing
Set ProfileContacts = Nothing
Set SyncItems = Nothing
Set Synchonizer = Nothing
Set Session = Nothing
Set SyncDataContact = Nothing

Set ws = Nothing
Set wb = Nothing

End Sub