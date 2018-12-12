Const LogFile="C:\Program Files (x86)\Redemption\MA Bday Anniversary\LogFile.txt"
Const MailBoxFile ="C:\Program Files (x86)\Redemption\MA Bday Anniversary\MailboxesToAddBdayAnniversary.txt"
Const CalenderItemCategory ="automatically added from PF - Arges Intern"
Const YearsInFuture = 3
Const olFolderCalendar = 9 

Call WriteToLog ("******* Start ADD Run *******")
Call AddMADatesToMailboxes(MailBoxFile)
Call WriteToLog ("******* Done with ADD Run *******")

'*****************************************************************************
Private Sub AddMADatesToMailboxes(txtPath)
' Read txt file with mailboxes and PF_Paths to add Calender Items
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
    SyncParas=split(arrFileLines(l),";")
    Call AddMADates (SyncParas(0),SyncParas(1),SyncParas(2))
Next
end sub

'*****************************************************************************
Sub WriteToLog (LogEntry)
Const ForAppending = 8

Set objFSO = CreateObject("Scripting.FileSystemObject")

If objFSO.FileExists(LogFile) then
	Set objFile = objFSO.GetFile(LogFile)
	'Wscript.Echo objFile.Size

	If objFile.Size > 10000000 then
  	
		objFile.delete
	end if
end if

Set objTextFile = objFSO.OpenTextFile(LogFile, ForAppending, True)
objTextFile.WriteLine(now() & " - " & LogEntry)
objTextFile.Close

end sub

'*****************************************************************************
Sub AddMADates (mailbox, mailserver, KontakteEntryID)

MAPI_NO_CACHE = &H200
MAPI_BEST_ACCESS = &H10

Dim Session
Dim SQLstr
Dim PFEntryID
Dim ExchangeStore
Dim PFstore
Dim Kalender
Dim MitarbeiterKontakte
Dim MA 
Dim AppointmentCat
Dim ToFutureYear
Dim KalenderEintrag
Dim Bday
Dim BdayThisYear
Dim BDayEintrag
Dim AnniThisYear
Dim AnniEintrag
Dim Anniversary
Dim Lebensalter
Dim Firmenzugehoerigkeit
dim Subject

Call WriteToLog ("************** start run: update employee dates to mailbox: " & mailbox & " **************")

PFEntryID = KontakteEntryID

If Not PFEntryID = "" Then
    Set Session = CreateObject("Redemption.RDOSession")
    Session.LogonExchangeMailbox mailbox, mailserver
    Set PFstore = Session.Stores.FindExchangePublicFoldersStore
    Set MitarbeiterKontakte = PFstore.GetFolderFromID(PFEntryID)
    Set Kalender = Session.GetDefaultFolder(olFolderCalendar)
    Set Items = Kalender.Items
        
    SQLstr = ""
    SQLstr = "SELECT EntryID FROM " & Kalender.Name & " WHERE categories='" & CalenderItemCategory & "' AND Start >= '" & YEAR(Date()) & "-" & Pd(Month(date()),2) & "-" & Pd(DAY(date()),2) & "';"
    'Wscript.Echo SQLstr
    Set Recordset = Items.MAPITable.ExecSQL(SQLstr)
    
    'Wscript.Echo Recordset.EOF   
    Do While Not Recordset.EOF
        Set KalenderEintrag = Session.GetMessageFromID(Recordset.Fields("EntryID").Value)
        Subject = KalenderEintrag.Subject
        KalenderEintrag.Delete
        Call WriteToLog ("Calendar Item deleted: " & Subject)
        'Wscript.Echo Subject
        Recordset.MoveNext
    Loop

    
    For Each MA In MitarbeiterKontakte.Items
        'Wscript.Echo MA
        Call CreateCalendarItem(MA, Kalender)
    Next
 end if
Call WriteToLog ("************** stop run: update employee dates for mailbox: " & mailbox & " **************")
End Sub

'****************************************************************************

Sub CreateCalendarItem (MA,Kalender)

    If not year(MA.Birthday)=4501 Then
        Bday = MA.Birthday
        BdayThisYear = Day(Bday) & "." & Month(Bday) & "." & Year(Now)
    Else
        Bday = "01.01.1900"               
    End If
    'WriteToLog (Bday &" "& MA.LastNameAndFirstName)
        
    If Not year(MA.Anniversary)=4501 Then
        Anniversary = MA.Anniversary
        AnniThisYear = Day(Anniversary) & "." & Month(Anniversary) & "." & Year(Now)
    Else
      Anniversary = "01.01.1900"
    End If
    'WriteToLog (Anniversary &" "& MA.LastNameAndFirstName)
    
        'Wscript.Echo (Kalender.Items.count)
        AppointmentCat = CalenderItemCategory
        ToFutureYears = YearsInFuture
        
        For i = 0 To ToFutureYears - 1
             'Geburtsatgseinträge setzten
             If year(Bday)<> 1900 Then
                BDayEintrag = DateAdd("yyyy", i, BdayThisYear)
                'Wscript.Echo BDayEintrag
                If BDayEintrag >= Date() Then
                    Lebensalter = DateDiff("yyyy", Bday, BDayEintrag)
                    Set KalenderEintrag = Kalender.Items.Add
                    KalenderEintrag.Start = BDayEintrag
                    KalenderEintrag.Categories = AppointmentCat
                    KalenderEintrag.AllDayEvent = True
                    KalenderEintrag.Subject = "Geburtstag " & MA.LastNameAndFirstName & " (" & Lebensalter & ") // Geburtstag: "& Bday 
                    KalenderEintrag.ReminderSet = False
                    KalenderEintrag.Save
                    Call WriteToLog ("Calendar Item created for "& KalenderEintrag.Start &": " & KalenderEintrag.Subject)
                End If
             end if   
            
            'Jahrestage setzten
            If year(Anniversary)<> 1900 Then
                AnniEintrag = DateAdd("yyyy", i, AnniThisYear)
                If AnniEintrag >= Date() Then
                    Firmenzugehoerigkeit = DateDiff("yyyy", Anniversary, AnniEintrag)
                    Set KalenderEintrag = Kalender.Items.Add
                    KalenderEintrag.Start = AnniEintrag
                    KalenderEintrag.Categories = AppointmentCat
                    KalenderEintrag.AllDayEvent = True
                    KalenderEintrag.Subject = MA.LastNameAndFirstName & " bei ARGES seit " & Firmenzugehoerigkeit & " Jahren // Einstellungsdatum: " & Anniversary
                    KalenderEintrag.ReminderSet = False
                    KalenderEintrag.Save
                    Call WriteToLog ("Calendar Item created for "& KalenderEintrag.Start &": " & KalenderEintrag.Subject)
                End If
            End If          
        Next
 End sub
 
 '*******************************************************************************
Function pd(n, totalDigits) 
    If totalDigits > len(n) then 
        pd = String(totalDigits-len(n),"0") & n 
    else 
        pd = n 
    End if 
End Function 
