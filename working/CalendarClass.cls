VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CalendarClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit
'
'============================================================================================================================
' Global Objects
'
' 1/17/05 Checkout test
'============================================================================================================================
'
'Private Main As MainClass
Private Main As Object
Private CSV As Object
'
Const cr1 = vbCrLf & "" & vbTab
Const cr2 = cr1 & vbTab
Const cr3 = cr2 & vbTab
Const cr4 = cr3 & vbTab
'
Const e_StartDate = 0
Const e_EndDate = 1
Const e_Name = 2
Const e_PageID = 3
Const e_Copyfilename = 4
Const e_Link = 5
Const e_EventID = 6
Const e_BriefFilename = 7

Const rnTabNumber = "tb"
Const rnCalID = "calendarid"
Const rnCalName = "calname"
Const rnCalStart = "calstrt"
Const rnCalEnd = "calend"
Const rnCalURL = "calurl"
Const rnCalDayCopy = "caldycpy"
Const rnCalDetailCopy = "caldtcpy"
Const rnCals = "cals"
Const rnCalRepeats = "calrpt"
Const rnCalRptStart = "calrptsrt"
Const rnCalRptEnd = "calrptend"
Const rnEventID = "evtid"
Const rnCalPrntID = "calprnt"
Const rnDefaultEventCalendarID = "defaultEventCalendarID"           ' when clicking the add button, this is the id of the calendar they are on, set during event save

Const rnCalORName = "calname_or"
Const rnCalORStart = "calstrt_or"
Const rnCalOREnd = "calend_or"
Const rnCalORURL = "calurl_or"
Const rnCalORDayCopy = "caldycpy_or"
Const rnCalORDetailCopy = "caldtcpy_or"
Const rnCalsOR = "cals_or"
Const rnCalORRepeats = "calrpt_or"
Const rnCalORRptStart = "calrptsrt_or"
Const rnCalORRptEnd = "calrptend_or"

Const btnSave = "Save"
Const btnSaveSeries = "Save This Series"
Const btnCancel = "Cancel"
Const btnDelete = "Delete This Occurrence"
Const btnDeleteSeries = "Delete This Series"

'
'============================================================================================================================
' Global variables
'============================================================================================================================
'
Private VerboseMessage As String
Private adminUrl_local As String
'
Private GlobalRootPageName As String
Private GlobalPageRecordID As Long
Private GlobalContentRecordKey As String
'
Private GlobalMonth As Long
Private GlobalYear As Long
Private GlobalDay As Long
Private GlobalDate As Date
Private GlobalDateLimitNext As Date
Private GlobalDateLimitPrevious As Date
Private GlobalCalendarEventsContentID As Long
Private CalendarCheck As String
'
Private GlobalCalendarName As String
Private GlobalCalendarID As Long
Private GlobalCalendarCaption As String
'Private GlobalCalendarPageContentID As Long ' The ContentID of the Page Content definition for non-linked records
'Private GlobalCalendarPageContentName As String
'
Private tabsDef(4, 2) As String
'
Private RequestLink As String
'
Const RequestNameTargetdate = "targetdate"
Const RequestNameReturnTabNumber = "returntb"
Const RequestNameNoIdea = "i"
Const RequestNameCRKey = "rdky"
Const RequestNamePageID = "pgid"
Const RequestNameEventID = "evtid"

Const rnCalendar = "calnm"

Const buttonUpdate = "  Update Calendar  "

Const TableWidth = "100%"
Const BackGroundColor = "lightgrey" ' this will be the background color of the Content Section, as well as the active tab
Const TabColor_Ghost = "darkgray" ' this will be the background color of the inactive tabs
'
Const FormDay = 1
Const FormWeek = 2
Const FormMonth = 0
Const FormYear = 3
Const FormDetails = 4
Const FormEdit = 5
'
' no idea
'
Private iYear
Private sMonth
Private j
Private k
Private sDate
Private sHref
Private sEvent
Private sScriptName
Private iMonthCols
Private sHTMLOut
Private dValue
Private iStartDay
Private iMonth
Private iDateIncrement
Private iCell
Private bValidCell
Private i
Private lMeetingId
Private SQLQuery
Private RSEvents
Private Conn
Private strConn
Private RS
'Private PageContentID As Long
'
Private AddOnVersion As String

Private calArray() As String
Private calArrayCount As Long
Private calSelected As String
'
Dim aryMeetings(1, 6)

Private calName As String
Private calStart As String
Private calEnd As String
Private calURL As String
Private calDayCopy As String
Private calDetailCopy As String
Private eventID As Long
Private calRepeats As Long
Private calRptStart As String
Private calRptEnd As String
Private calPrntID As Long
Private calDefaultEventCalendarID As Long                   ' calendarid set when adding a new event, to populate the event rule

Private calORName As String
Private calORStart As String
Private calOREnd As String
Private calORURL As String
Private calORDayCopy As String
Private calORDetailCopy As String
Private calORRepeats As Long
Private calORRptStart As String
Private calORRptEnd As String

Private AllowEdit As Boolean

Private btnValue As String

Public Function execute(CsvObject As Object, MainObject As Object, OptionString As String, FilterInput As String) As String
    On Error GoTo ErrorTrap
    
    Set CSV = CsvObject
    
    Call Init(MainObject)
    
    AllowEdit = Main.IsLinkAuthoring("Calendar Events")
    
    execute = GetContent(OptionString)
    
    Exit Function
ErrorTrap:
    Call HandleError("CalendarClass", "Execute", Err.Number, Err.Source, Err.Description, True, False)
    End Function

 Public Sub Init(WebClientMain As Object)
   On Error GoTo ErrorTrap
    Set Main = WebClientMain
    '
    Dim EventPageID As Long
    '
    AddOnVersion = App.Major & "." & App.Minor & "." & App.Revision
    '
    Call UpgradeCalendarEvents
    '
    EventPageID = Main.GetStreamInteger(RequestNamePageID)
    '
    '   Redirect to the event's content page
    '
    If EventPageID <> 0 Then
        Call Main.Redirect(Main.GetPageLink(EventPageID))
    End If
    '
    If Main.RefreshQueryString = "" Then
        RequestLink = Main.ServerPage & "?"
    Else
        RequestLink = Main.ServerPage & "?" & Main.RefreshQueryString & "&"
    End If
    
    Exit Sub
ErrorTrap:
    Call HandleTrapError("Init")
End Sub
'
'============================================================================================================================
'   Aggregate Object Interface
'============================================================================================================================
'
 Public Function GetContent(Options As String)
    On Error GoTo ErrorTrap
    '
    Dim CalendarName As String
    Dim TickTime As Long
    Dim calList As String
    '
    TickTime = GetTickCount
    CalendarName = Trim(Main.GetAggrOption("CalendarName", Options))
    
    calList = Main.GetAggrOption("Calendar List", Options)
    If calList <> "" Then
        If InStr(1, calList, ",", vbTextCompare) <> 0 Then
            calArray = Split(calList, ",", , vbTextCompare)
            calArrayCount = UBound(calArray)
            
            Call setSelectedCalendars
            calSelected = Main.GetMemberProperty("Calendars Selected")
        End If
    End If
    
    If (CalendarName = "") And (calList = "") Then
        CalendarName = "Default"
    End If
    
    If calSelected = "" Then
        calSelected = "," & calList & ","
    End If
    
    GlobalCalendarName = CalendarName
    GetContent = GetCalendar(CalendarName)
    If Main.Testing Then
        Call Main.TestPoint("<div>calendar render time = " & (GetTickCount - TickTime) & "</div>" & VerboseMessage)
    End If
    '
    If InStr(1, GetContent, "<!-- noWrap -->") = 0 Then
        GetContent = "" _
            & cr1 & "<div class=""calendarCon"">" _
            & GetContent _
            & cr1 & "</div>"
    End If
    '
    Exit Function
ErrorTrap:
    Call HandleTrapError("GetContent")
End Function
''
''============================================================================================================================
'' The Calendar Application Developed by Kma Net http://www.contensive.com
'' This Calendar Application Displays the Events Entered.
'' This Standard Calendar Module was developed on June 07 2002.
'' CopyRight(c) Kma Net
'' Last Authored by Santosh, Kma Net
''
''   If LinkAuthoring PageContent, display the Calendars Root page (and child pages)
''       This also covers the bid case, and with Link Authoring, Content Watch is not set
''   if CRKey detected, display the non-mapped ContentPage (convert crkey to BID argument)
''   Otherwise, print the Calendar Root page followed by the Calendar
''============================================================================================================================
''
'Public Function GetForm(Optional CalendarName As Variant, Optional TargetDate As Variant) As String
'    On Error GoTo ErrorTrap
'    '
'    Dim Stream As String
'    Dim MondayDate As Date
'    Dim Week As Long
'    Dim CurrentFormID As Long
'    Dim ReturnFormID As Long
'    Dim BackGroundColor As String
'    Dim Copy As String
'    Dim EventContentName As String
'    Dim CSPointer As Long
'    Dim PageList As Object
'    Dim Position As Long
'    Dim CRKeyContentID As Long
'    Dim CRKeyRecordID As Long
'    Dim ContentName As String
'    Dim TableName As String
'    Dim DisplayCalendar As Boolean
'    Dim QueryString As String
'    Dim CSPage As Long
'    Dim CSRootPage As Long
'    Dim ParentID As Long
'    Dim ChildID As Long
'    Dim Link As String
'    Dim TickTime As Long
'    '
'    TickTime = GetTickCount()
'    If Not (Main Is Nothing) Then
'        '
'        ' ----- Initialize environment
'        '
'        Call SetRequests(CalendarName, TargetDate)
'        CurrentFormID = Main.GetStreamInteger(rnTabNumber)
'        ReturnFormID = Main.GetStreamInteger(RequestNameReturnTabNumber)
'        '
'        ' ----- Look up the calendar to get the PageContentID
'        '
'        CSPointer = Main.OpenCSContent_Internal("Calendars", "name=" & KmaEncodeSQLText(GlobalCalendarName), , , , , "ID,Caption,PageContentID")
'        If Not Main.IsCSOK(CSPointer) Then
'            '
'            ' New Calendar, create it now
'            '
'            Call Main.CloseCS(CSPointer)
'            CSPointer = Main.InsertCSContent("Calendars")
'            If Main.IsCSOK(CSPointer) Then
'                GlobalCalendarID = Main.GetCSInteger(CSPointer, "ID")
'                GlobalCalendarCaption = GlobalCalendarName
'                GlobalCalendarPageContentID = Main.GetContentID("Page Content")
'                Call Main.SetCS(CSPointer, "active", True)
'                Call Main.SetCS(CSPointer, "Name", GlobalCalendarName)
'                Call Main.SetCS(CSPointer, "caption", GlobalCalendarCaption)
'                Call Main.SetCS(CSPointer, "PageContentID", GlobalCalendarPageContentID)
'            End If
'        Else
'            '
'            ' Get Calendar
'            '
'            GlobalCalendarID = Main.GetCSInteger(CSPointer, "ID")
'            GlobalCalendarCaption = Main.GetCSText(CSPointer, "caption")
'            GlobalCalendarPageContentID = Main.GetCSInteger(CSPointer, "PageContentID")
'        End If
'        Call Main.CloseCS(CSPointer)
'        GlobalCalendarPageContentName = Main.GetContentNameByID(GlobalCalendarPageContentID)
'        '
'        ' ----- If a page content definition is selected, include calendar content pages as well
'        '
'        DisplayCalendar = True
'        If (GlobalCalendarPageContentName <> "") Then
'            '
'            ' If a Page Content was selected, check for Non-Mapped page mode
'            '
'            If GlobalContentRecordKey <> "" Then
'                '
'                ' GlobalContentRecordKey reference that did not forward (no Content Watch)
'                ' Translate to BID reference and redirect - this will create the correct Contentwatch record
'                '
'                Position = InStr(1, GlobalContentRecordKey, ".")
'                If Position > 1 Then
'                    CRKeyContentID = kmaEncodeInteger(Mid(GlobalContentRecordKey, 1, Position - 1))
'                    CRKeyRecordID = kmaEncodeInteger(Mid(GlobalContentRecordKey, Position + 1))
'                    If (CRKeyContentID <> 0) And (CRKeyRecordID <> 0) Then
'                        ContentName = Main.GetContentNameByID(CRKeyContentID)
'                        If Main.IsWithinContent(ContentName, GlobalCalendarPageContentName) Then
'                            '
'                            ' All is OK with this CRKey record, redirect and exit now
'                            '
'                            Link = kmaModifyLinkQuery(Main.ServerLink, "CRKey", "", False)
'                            Link = kmaModifyLinkQuery(Link, "bid", kmaEncodeText(CRKeyRecordID), True)
'                            Call Main.Redirect(Link)
'                            DisplayCalendar = False
'                        End If
'                    End If
'                End If
'            ElseIf (GlobalPageRecordID <> 0) Then
'                '
'                ' Display the Calendar Content Page from GlobalPageRecordID
'                '
'                Stream = Stream & Main.GetContentPage(GlobalRootPageName, GlobalCalendarPageContentName, , False, False, GlobalPageRecordID)
'                Link = ModifyQueryString(Main.RefreshQueryString, "bid", "", False)
'                If Link = "" Then
'                    Stream = Stream & "<a href=""" & Main.ServerPage & """>Return to Calendar</a>"
'                Else
'                    Stream = Stream & "<a href=""" & Main.ServerPage & "?" & Link & """>Return to Calendar</a>"
'                End If
'                DisplayCalendar = False
'            End If
'        End If
'        '
'        If DisplayCalendar Then
'            '
'            ' ----- Calendar Display Mode
'            '
'            If Main.IsAuthoring(GlobalCalendarPageContentName) Then
'                '
'                ' ----- Link Authoring, display the non-mapped content page tree
'                '
'                Stream = Stream & Main.GetContentPage(GlobalRootPageName, GlobalCalendarPageContentName, , False)
'            Else
'                '
'                ' ----- non Authoring, display the content root page
'                '
'                Stream = Stream & Main.GetContentPage(GlobalRootPageName, GlobalCalendarPageContentName, , False, False, 0)
'                End If
'            Stream = Stream & GetCalendar(GlobalCalendarName, TargetDate)
'            End If
'        End If
'    GetForm = Stream & "<BR>t=" & (GetTickCount - TickTime)
'    '
'    Exit Function
'ErrorTrap:
'    Set PageList = Nothing
'    Call HandleTrapError("GetForm")
'    End Function
'
'============================================================================================================================
' The Calendar Application Developed by Kma Net http://www.contensive.com
' This Calendar Application Displays the Events Entered.
' This Standard Calendar Module was developed on June 07 2002.
' CopyRight(c) Kma Net
' Last Authored by Santosh, Kma Net
'
'   If LinkAuthoring PageContent, display the Calendars Root page (and child pages)
'       This also covers the bid case, and with Link Authoring, Content Watch is not set
'   if CRKey detected, display the non-mapped ContentPage (convert crkey to BID argument)
'   Otherwise, print the Calendar Root page followed by the Calendar
'============================================================================================================================
'
Public Function GetCalendar(Optional CalendarName As Variant, Optional TargetDate As Variant) As String
    On Error GoTo ErrorTrap
    '
    Dim stream As String
    Dim MondayDate As Date
    Dim Week As Long
    Dim CurrentFormID As Long
    Dim BackGroundColor As String
    Dim Copy As String
    Dim EventContentName As String
    Dim CSPointer As Long
    Dim PageList As Object
    Dim Position As Long
    'Dim RecordID As Long
    Dim CRKeyContentID As Long
    Dim CRKeyRecordID As Long
    Dim ContentName As String
    Dim TableName As String
    'Dim GlobalContentRecordKey As String
    Dim DisplayCalendar As Boolean
    Dim QueryString As String
    'Dim GlobalRootPageName As String
    Dim CSPage As Long
    Dim CSRootPage As Long
    Dim ParentID As Long
    Dim ChildID As Long
    Dim Link As String
    Dim calendarId As Long
    Dim CalendarEventsContentID As Long
    '
    CalendarEventsContentID = Main.GetContentID("Calendar Events")
    '
    If Not (Main Is Nothing) Then
        '
        ' ----- Initialize environment
        '
        Call SetRequests(CalendarName, TargetDate)
        calendarId = Main.GetRecordID("Calendars", GlobalCalendarName)
        CurrentFormID = Main.GetStreamInteger(rnTabNumber)
        '
        MondayDate = CDate(GlobalDate) - Weekday(CDate(GlobalDate)) + 2
        '
        ' ----- Display calendar forms
        '
        If CurrentFormID <> FormDetails Then
            stream = stream & GetForm_Tabs(CurrentFormID)
        End If
        '
        ' Select the appropriate tab (form)
        '
        Select Case CurrentFormID
            Case FormDay
                '
                ' ----- Day Events Tab
                '
                stream = stream & GetForm_Day(CalendarEventsContentID)
            Case FormWeek
                '
                ' ----- Week Events Tab
                '
                stream = stream & GetForm_Week(CalendarEventsContentID)
            Case FormYear
                '
                ' ----- Year Events Tab
                '
                stream = stream & GetForm_Year
            Case FormDetails
                '
                ' ----- Event with no link - show page
                '
                stream = stream & GetForm_Details
            Case FormEdit
                '
                ' ----- Edit/Add an event
                '
                Call loadEditForm
                Call processEditForm
                If btnValue <> "" Then
                    stream = stream & GetForm_Month(calendarId)
                Else
                    stream = stream & getFormEdit(calendarId)
                End If
            Case Else
                '
                ' ----- Month Events Tab
                '
                stream = stream & GetForm_Month(calendarId)
            End Select
        End If
    GetCalendar = getNavigation & stream
    '
    Exit Function
ErrorTrap:
    Set PageList = Nothing
    Call HandleTrapError("GetCalendar")
    End Function
'
'============================================================================================================================
'
'============================================================================================================================
'
Public Sub SetRequests(Optional CalendarName As Variant, Optional TargetDate As Variant)
    On Error GoTo ErrorTrap
    '
    Dim stream As String
    Dim MondayDate As Date
    Dim Week As Long
    Dim CurrentFormID As Long
    Dim BackGroundColor As String
    Dim Copy As String
    Dim EventContentName As String
    Dim CSPointer As Long
    Dim PageList As Object
    Dim Position As Long
    'Dim RecordID As Long
    Dim CRKeyContentID As Long
    Dim CRKeyRecordID As Long
    Dim ContentName As String
    Dim TableName As String
    'Dim GlobalContentRecordKey As String
    Dim DisplayCalendar As Boolean
    Dim QueryString As String
    'Dim GlobalRootPageName As String
    Dim CSPage As Long
    Dim CSRootPage As Long
    Dim ParentID As Long
    Dim ChildID As Long
    Dim Link As String
    '
    If Not (Main Is Nothing) Then
        '
        ' ----- Initialize environment
        '
        GlobalCalendarName = KmaEncodeMissingText(CalendarName, "Default")
        GlobalRootPageName = GlobalCalendarName & " Calendar Root Page"
        GlobalPageRecordID = Main.GetStreamInteger("bid")
        GlobalContentRecordKey = Main.GetStreamText(RequestNameCRKey)
        '
        iMonthCols = 4
        aryMeetings(0, 0) = "1"
        aryMeetings(1, 0) = Date
        '
        'GlobalValueNoIdea = Main.GetStreamInteger(RequestNameNoIdea)
        CurrentFormID = Main.GetStreamInteger(rnTabNumber)
        Call Main.AddRefreshQueryString(rnTabNumber, CurrentFormID)
        '
        ' ----- Create the Calendar Page
        '
        GlobalDate = KmaEncodeDate(Main.GetStreamDate(RequestNameTargetdate))
        If GlobalDate = CDate(0) Then
            GlobalDate = KmaEncodeDate(TargetDate)
            If GlobalDate = CDate(0) Then
                GlobalDate = Date
            End If
        End If
        Call Main.AddRefreshQueryString(RequestNameTargetdate, GlobalDate)
        '
        GlobalDateLimitNext = Date + (kmaEncodeInteger(Main.SiteProperty_CalendarYearLimit) * 365)
        If GlobalDate > GlobalDateLimitNext Then
            GlobalDate = GlobalDateLimitNext
        End If
        '
        GlobalDateLimitPrevious = Date - (kmaEncodeInteger(Main.SiteProperty_CalendarYearLimit) * 365)
        If GlobalDate < GlobalDateLimitPrevious Then
            GlobalDate = GlobalDateLimitPrevious
        End If
        '
        GlobalMonth = Month(GlobalDate)
        GlobalYear = Year(GlobalDate)
        GlobalDay = Day(GlobalDate)
        GlobalCalendarEventsContentID = Main.GetContentID("Calendar Events")
        '
        ' ----- Look up the calendar to get the PageContentID
        '
        CSPointer = Main.OpenCSContent_Internal("Calendars", "name=" & KmaEncodeSQLText(GlobalCalendarName), , , , , "ID,Caption")
        'CSPointer = Main.OpenCSContent_Internal("Calendars", "name=" & KmaEncodeSQLText(GlobalCalendarName), , , , , "ID,Caption,PageContentID")
        If Not Main.IsCSOK(CSPointer) Then
            '
            ' New Calendar, create it now
            '
            Call Main.CloseCS(CSPointer)
            CSPointer = Main.InsertCSContent("Calendars")
            If Main.IsCSOK(CSPointer) Then
                GlobalCalendarID = Main.GetCSInteger(CSPointer, "ID")
                GlobalCalendarCaption = GlobalCalendarName
                'GlobalCalendarPageContentID = Main.GetContentID("Page Content")
                Call Main.SetCS(CSPointer, "active", True)
                Call Main.SetCS(CSPointer, "Name", GlobalCalendarName)
                Call Main.SetCS(CSPointer, "caption", GlobalCalendarCaption)
                'Call Main.SetCS(CSPointer, "PageContentID", GlobalCalendarPageContentID)
            End If
        Else
            '
            ' Get Calendar
            '
            GlobalCalendarID = Main.GetCSInteger(CSPointer, "ID")
            GlobalCalendarCaption = Main.GetCSText(CSPointer, "caption")
            'GlobalCalendarPageContentID = Main.GetCSInteger(CSPointer, "PageContentID")
        End If
        Call Main.CloseCS(CSPointer)
        'GlobalCalendarPageContentName = Main.GetContentNameByID(GlobalCalendarPageContentID)
        End If
    '
    Exit Sub
ErrorTrap:
    Set PageList = Nothing
    Call HandleTrapError("SetRequests")
    End Sub
'
'============================================================================================================================
' The Calendar Application Developed by Kma Net http://www.contensive.com
' This Calendar Application Displays the Events Entered.
' This Standard Calendar Module was developed on June 07 2002.
' CopyRight(c) Kma Net
' Last Authored by Santosh, Kma Net
'============================================================================================================================
'
Public Function GetForm_Old(Optional CalendarName As Variant, Optional TargetDate As Variant) As String
    End Function
'////////////////////////////////////////////////////////////////////////////////////////
'/// Public Function TimeWithoutSeconds( as_time ) returns int
'///
'/// purpose:  This function returns the non military hour ( 1 thru 12 )
'///           The vbscript hour funciton returns the hour ( 1 thru 24 )
'///
'///////////////////////////////////////////////////////////////////////////////////////
Private Function TimeWithoutSeconds(as_time)
    On Error GoTo ErrorTrap
    '
    ' Declare Variables
    Dim la_time
    Dim la_time_mh
    ' split up the time into an arrray
    If as_time <> "" Then
        la_time = Split(as_time, ":")
        la_time_mh = Split(la_time(2), " ")
        TimeWithoutSeconds = la_time(0) & ":" & la_time(1) & " " & la_time_mh(1)
    Else
        TimeWithoutSeconds = -1
    End If
    '
    Exit Function
ErrorTrap:
    Call HandleTrapError("TimeWithoutSeconds")
    End Function
'////////////////////////////////////////////////////////////////////////////////////////
'/// Public Function WordWrap(strTextToBeWrapped, intMaxLineLength) returns string
'///
'/// purpose: Here's the function that does the work.  It takes a string
'///          to word wrap and a maximum line length and returns the
'///          string wrapped appropriately.
'///
'///////////////////////////////////////////////////////////////////////////////////////
Private Function WordWrap(strTextToBeWrapped, intMaxLineLength)
    On Error GoTo ErrorTrap
    '
             Dim strWrappedText           ' Result storage
             Dim intLengthOfInput         ' Length of original
             Dim intCurrentPosition       ' Where we're at now
             Dim intCurrentLineStart      ' Where the current line starts
             Dim intPositionOfLastSpace   ' Last space we saw
            
             ' Get this once so we don't have to keep checking
             intLengthOfInput = Len(strTextToBeWrapped)
                
             ' Start both of these at the beginning
             intCurrentPosition = 1
             intCurrentLineStart = 1
                
                
             ' Loop through until we get to the end
             Do While intCurrentPosition < intLengthOfInput
               ' If the current position is a space, make a note of
                  ' it's location for later use.
               If Mid(strTextToBeWrapped, intCurrentPosition, 1) = " " Then
                     intPositionOfLastSpace = intCurrentPosition
                  End If
                    
                  ' If we're at what should be the end of a line, we go back
               ' to the last space we saw and cut the line there.
                  If intCurrentPosition = intCurrentLineStart + intMaxLineLength Then
                    ' Some debugging lines if something's not lining up.
                     'Stream = Stream &  intCurrentLineStart & "<br />"
                       'Stream = Stream &  intPositionOfLastSpace & "<br />"
                    'Stream = Stream &  Trim(Mid(strTextToBeWrapped, intcurrentLineStart, _
                     ' intPositionOfLastSpace - intCurrentLineStart + 1)) & "<br />"

                       ' Append this latest line to our result
                    strWrappedText = strWrappedText _
                                & Trim(Mid(strTextToBeWrapped, _
                                           intCurrentLineStart, _
                                           intPositionOfLastSpace - intCurrentLineStart + 1)) _
                                & "<br>" & vbCrLf   ' added the "<br>", because it wasn't wrapping

                     ' Reset the next line's starting point to the point we
                       ' used for the last one's end + 1.
                    intCurrentLineStart = intPositionOfLastSpace + 1
                        
                       ' Remove any leading spaces that might mess up our
                    ' character count.  If you want to just pull of one,
                     ' switch this to a simple If conditional instead of
                       ' looping.
                   Do While Mid(strTextToBeWrapped, intCurrentLineStart, 1) = " "
                        intCurrentLineStart = intCurrentLineStart + 1
                    Loop
                  End If
                    
                  ' Increment our location indicator.
               intCurrentPosition = intCurrentPosition + 1
             Loop
                
             ' Since the loop ends before we add the remaining text,
             ' add remaining text as the last line.
                strWrappedText = strWrappedText & Trim(Mid(strTextToBeWrapped, _
           intCurrentLineStart)) & vbCrLf

               ' Return our result to the calling line.
             WordWrap = strWrappedText
    '
    Exit Function
ErrorTrap:
    Call HandleTrapError("WordWrap")
    End Function

'////////////////////////////////////////////////////////////////////////////////////////
'/// Public Function GetDayNumberMonthStartsOn() returns int
'///
'/// purpose: Determines the day of the week that the month starts on.
'///
'///////////////////////////////////////////////////////////////////////////////////////
Private Function GetDayNumberMonthStartsOn(dAnyDayInTheMonth)
    On Error GoTo ErrorTrap
    '
            Dim dTemp
            dTemp = DateAdd("d", -(Day(dAnyDayInTheMonth) - 1), dAnyDayInTheMonth)
            GetDayNumberMonthStartsOn = Weekday(dTemp)
            '
    Exit Function
ErrorTrap:
    Call HandleTrapError("GetDayNumberMonthStartsOn")
    End Function
 ' GetDayNumberMonthStartsOn
'
'
'
Private Function GetPreviousMonth(dDate)
    On Error GoTo ErrorTrap
    '
            GetPreviousMonth = DateAdd("m", -1, dDate)
            '
    Exit Function
ErrorTrap:
    Call HandleTrapError("GetPreviousMonth")
    End Function
 ' GetPreviousMonth
'
'
'
Private Function GetNextMonth(dDate)
    On Error GoTo ErrorTrap
    '
            GetNextMonth = DateAdd("m", 1, dDate)
            '
    Exit Function
ErrorTrap:
    Call HandleTrapError("GetNextMonth")
    End Function
 ' GetNextMonth
'
'
'
Private Function GetDaysInMonth(iMonth, iYear)
    On Error GoTo ErrorTrap
    '
     Dim dTemp
     dTemp = DateAdd("d", -1, DateSerial(iYear, iMonth + 1, 1))
     GetDaysInMonth = Day(dTemp)
    '
    Exit Function
ErrorTrap:
    Call HandleTrapError("GetDaysInMonth")
    End Function
'
Public Function GetForm_Month_Details(CellDate As Date, Events As Variant, EventCnt As Long, EditLinkTemplate As String) As String
    On Error GoTo ErrorTrap
    '
    Dim CSPointer As Long
    Dim MoreRecords As Boolean
    Dim ContentRecordKey As String
    Dim PageID As Long
    Dim EventPtr As Long
    Dim EditLink As String
    '
    ' List events on this date
    '
    If EventCnt > 0 Then
        For EventPtr = 0 To EventCnt - 1
            If (KmaEncodeDate(Events(e_StartDate, EventPtr)) <= CellDate) And (KmaEncodeDate(Events(e_EndDate, EventPtr)) >= CellDate) Then
                PageID = kmaEncodeInteger(Events(e_PageID, EventPtr))
                GetForm_Month_Details = GetForm_Month_Details _
                    & GetEditLink(kmaEncodeInteger(Events(e_EventID, EventPtr))) _
                    & GetCalendarEventLink3(Events, EventPtr, FormMonth, CellDate)
'                GetForm_Month_Details = GetForm_Month_Details _
'                    & "<BR>" & Replace(EditLinkTemplate, "xxx", kmaEncodeInteger(Events(e_EventID, EventPtr))) _
'                    & GetCalendarEventLink3(Events, EventPtr, FormMonth, CellDate)
            End If
        Next
        'If GetForm_Month_Details <> "" Then
        '    GetForm_Month_Details = Mid(GetForm_Month_Details, 5)
        'End If
    End If
    '
    ' loop thru the recordset
    '
'    CSPointer = OpenCSEvents(CellDate, CellDate)
'    MoreRecords = Main.IsCSOK(CSPointer)
'    Do While MoreRecords
'        GetForm_Month_Details = GetForm_Month_Details & Main.GetCSRecordEditLink(CSPointer)
'        PageID = Main.GetCSInteger(CSPointer, "PageID")
'        'GetForm_Month_Details = GetForm_Month_Details & GetCalendarEventLink(Main.GetCSText(CSPointer, "ContentRecordKey"), Main.GetCSText(CSPointer, "Name"), CellDate)
'        'GetForm_Month_Details = GetForm_Month_Details & GetCalendarEventLink(ContentRecordKey, Main.GetCSText(CSPointer, "Name"), CellDate)
'        GetForm_Month_Details = GetForm_Month_Details & GetCalendarEventLink2(CSPointer)
'        Call Main.NextCSRecord(CSPointer)
'        MoreRecords = Main.IsCSOK(CSPointer)
'        If MoreRecords Then
'            GetForm_Month_Details = GetForm_Month_Details & "<br>"
'            End If
'        ContentRecordKey = ""
'        Loop
'    Call Main.CloseCS(CSPointer)
    '
    Exit Function
ErrorTrap:
    Call HandleTrapError("GetForm_Month_Details")
    End Function
''
'Private Function GetCalendarEventLink(ContentRecordKey As String, Menuheadline As String, EventDate As Date) As String
'    On Error GoTo ErrorTrap
'    '
'    Dim QueryString As String
'    '
'    GetCalendarEventLink = "<a href=""" & RequestLink & rnTabNumber & "=" & FormDay & "&" & RequestNameTargetdate & "=" & EventDate & "&" & RequestNameCRKey & "=" & ContentRecordKey & """>" & Menuheadline & "</a>"
'    '
'    Exit Function
'ErrorTrap:
'    Call HandleTrapError("GetCalendarEventLink")
'    End Function
'
Private Function GetCalendarEventLink3(Events As Variant, EventPtr As Long, ReturnTab As Long, ReturnTargetDate As Date) As String
    On Error GoTo ErrorTrap
    '
    Dim QueryString As String
    Dim Caption As String
    Dim Link As String
    Dim Description As String
    Dim eventID As Long
    Dim EventDate As Date
    '
    Dim LinkFlag As Boolean
    '
    eventID = kmaEncodeInteger(Events(e_EventID, EventPtr))
    If eventID > 0 Then
        Caption = kmaEncodeText(Events(e_Name, EventPtr))
        If Caption = "" Then
            Caption = "Event " & eventID
        End If
        Link = kmaEncodeText(Events(e_Link, EventPtr))
        Description = kmaEncodeText(Events(e_Copyfilename, EventPtr))
        EventDate = KmaEncodeDate(Events(e_StartDate, EventPtr))
        If Caption = "" Then
            Caption = "Event " & eventID
        End If
        '
        LinkFlag = (Description <> "") Or (Link <> "")
        '
        'If LinkFlag Then
            GetCalendarEventLink3 = "<a class=""CalendarEventLink"" "
            If Link <> "" Then
                If (InStr(1, Link, Main.ServerHost, vbTextCompare) = 0) And (Left(Link, 1) <> "/") Then
                    GetCalendarEventLink3 = GetCalendarEventLink3 & "target=""_blank"" "
                    End If
                GetCalendarEventLink3 = GetCalendarEventLink3 & "href=""" & Link & """"
                '
                '   Page Id short lived - use link field type
                '
                'GetCalendarEventLink3 = GetCalendarEventLink3 & "&" & RequestNamePageID & "=" & PageID
            Else
                '
                '   link directly to details page
                '
                GetCalendarEventLink3 = GetCalendarEventLink3 & "href=""" & RequestLink & rnTabNumber & "=" & FormDetails & "&" & RequestNameTargetdate & "=" & ReturnTargetDate & "&" & RequestNameEventID & "=" & eventID & "&" & RequestNameReturnTabNumber & "=" & ReturnTab
                End If
            '
            GetCalendarEventLink3 = GetCalendarEventLink3 & """>" & Caption & "</a>"
            '
        'Else
        '    GetCalendarEventLink3 = Caption
        '    End If
    End If
    '
    Exit Function
ErrorTrap:
    Call HandleError("CalendarClass", "GetCalendarEventLink3", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
Private Function GetCalendarEventLink2(CSPointer As Long) As String
    On Error GoTo ErrorTrap
    '
    Dim QueryString As String
    Dim Caption As String
    Dim Link As String
    Dim Description As String
    Dim eventID As Long
    Dim EventDate As Date
    '
    Dim LinkFlag As Boolean
    '
    eventID = Main.GetCSInteger(CSPointer, "ID")
    Caption = Main.GetCSText(CSPointer, "Name")
    If Caption = "" Then
        Caption = "Event " & eventID
    End If
    Link = Main.GetCSText(CSPointer, "EventLink")
    Description = Main.GetCSTextFile(CSPointer, "CopyFileName")
    EventDate = Main.GetCSDate(CSPointer, "StartDate")
    '
    LinkFlag = (Description <> "") Or (Link <> "")
    '
    If LinkFlag Then
        GetCalendarEventLink2 = "<a class=""CalendarEventLink"" "
        If Link <> "" Then
            If (InStr(1, Link, Main.ServerHost, vbTextCompare) = 0) And (Left(Link, 1) <> "/") Then
                GetCalendarEventLink2 = GetCalendarEventLink2 & "target=""_blank"" "
                End If
            GetCalendarEventLink2 = GetCalendarEventLink2 & "href=""" & Link & """"
            '
            '   Page Id short lived - use link field type
            '
            'GetCalendarEventLink2 = GetCalendarEventLink2 & "&" & RequestNamePageID & "=" & PageID
        Else
            '
            '   link directly to details page
            '
            GetCalendarEventLink2 = GetCalendarEventLink2 & "href=""" & RequestLink & rnTabNumber & "=" & FormDetails & "&" & RequestNameTargetdate & "=" & EventDate & "&" & RequestNameEventID & "=" & eventID
            End If
        '
        GetCalendarEventLink2 = GetCalendarEventLink2 & """>" & Caption & "</a>"
        '
    Else
        GetCalendarEventLink2 = Caption
        End If
    '
    Exit Function
ErrorTrap:
    Call HandleError("CalendarClass", "GetCalendarEventLink2", Err.Number, Err.Source, Err.Description, True, False)
    End Function


'=========================================================================================================================
                ' Function to display every week starting from Current week Monday
'=========================================================================================================================

Private Function dayName(iwd)
    On Error GoTo ErrorTrap
    '
            Dim owd
            '
            If iwd = 1 Then
                owd = "Sunday"
            ElseIf iwd = 2 Then
                owd = "Monday"
            ElseIf iwd = 3 Then
                owd = "Tuesday"
            ElseIf iwd = 4 Then
                owd = "Wednesday"
            ElseIf iwd = 5 Then
                owd = "Thursday"
            ElseIf iwd = 6 Then
                owd = "Friday"
            ElseIf iwd = 7 Then
                owd = "Saturday"
            End If
                
            dayName = owd
        '
    Exit Function
ErrorTrap:
    Call HandleTrapError("dayName")
    End Function

'
'============================================================================================================================
' Function to Get a Pretty Date
'============================================================================================================================
'
Private Function PrettyDate(DateValue As Date) As String
    On Error GoTo ErrorTrap
    '
    PrettyDate = dayName(Weekday(DateValue)) & ", " & Day(DateValue) & " " & MonthName(Month(DateValue)) & " " & Year(DateValue)
    '
    Exit Function
ErrorTrap:
    Call HandleTrapError("PrettyDate")
    End Function

'
'=================================================================================================
        '---/Start Tab Content
'=================================================================================================
'
Private Function GetForm_Start(FormID, TabHeadline) As String
    On Error GoTo ErrorTrap
    '
    GetForm_Start = GetForm_Start & cr1 & "<div name=""Form" & FormID & """ id=""Form" & FormID & """>"
    GetForm_Start = GetForm_Start & cr2 & "<table Border=""0"" Width=""100%"">"
    GetForm_Start = GetForm_Start & cr3 & "<tr><td>" & TabHeadline & "</td></tr>"
    GetForm_Start = GetForm_Start & cr3 & "<tr><td><font size=""2"">"
    '
    Exit Function
ErrorTrap:
    Call HandleTrapError("GetForm_Start")
    End Function

'
'=================================================================================================
        '---/End Tab Content
'=================================================================================================
'
Private Function GetForm_End()
    On Error GoTo ErrorTrap
    '
    GetForm_End = GetForm_End _
        & cr3 & "</font></td></tr>" _
        & cr2 & "</table>" _
        & cr1 & "</div>"
    '
    Exit Function
ErrorTrap:
    Call HandleTrapError("GetForm_End")
    End Function

'
'=================================================================================================
        '---/Draw tabs
'=================================================================================================
'
Private Function GetForm_Tabs(CurrentFormID As Long) As String
    On Error GoTo ErrorTrap
    '
    Dim buttonActive As String
    Dim buttonInactive As String
    '
    GetForm_Tabs = ""
    '
    buttonActive = "<div class=""calTab calTabActive""><NOBR>x2x</NOBR></div>"
    buttonInactive = "<div class=""calTab calTabInactive""><NOBR><a href=""javascript: document.location='" & RequestLink & rnTabNumber & "=x1x&" & RequestNameTargetdate & "=" & GlobalDate & "';"" class=""CalendarHeaderInActive"">x2x</a></NOBR></div>"
    '
    'buttonActive = "<p align=""center"" class=""CalendarHeaderActive""><NOBR>x2x</NOBR></p>"
    'buttonInactive = "<p align=""center"" class=""CalendarHeaderInactive""><NOBR><a href=""" & RequestLink & rnTabNumber & "=x1x&" & RequestNameTargetdate & "=" & GlobalDate & """ class=""CalendarHeaderInActive"">x2x</a></NOBR></p>"
    'buttonActive = "<td width=""100"" class=""calTabActive"">" & Main.GetPanel(buttonActive, "CalendarHeaderActive", , , 100, 2) & "</td>"
    'buttonInactive = "<td width=""100"" class=""calTabInactive"">" & Main.GetPanel(buttonInactive, , , , 100, 2) & "</td>"
    Select Case CurrentFormID
        Case FormDay
            GetForm_Tabs = GetForm_Tabs & Replace(Replace(buttonActive, "x1x", FormDay), "x2x", "Day View")
            GetForm_Tabs = GetForm_Tabs & Replace(Replace(buttonInactive, "x1x", FormWeek), "x2x", "Week View")
            GetForm_Tabs = GetForm_Tabs & Replace(Replace(buttonInactive, "x1x", FormMonth), "x2x", "Month View")
            GetForm_Tabs = GetForm_Tabs & Replace(Replace(buttonInactive, "x1x", FormYear), "x2x", "Year View")
        Case FormWeek
            GetForm_Tabs = GetForm_Tabs & Replace(Replace(buttonInactive, "x1x", FormDay), "x2x", "Day View")
            GetForm_Tabs = GetForm_Tabs & Replace(Replace(buttonActive, "x1x", FormWeek), "x2x", "Week View")
            GetForm_Tabs = GetForm_Tabs & Replace(Replace(buttonInactive, "x1x", FormMonth), "x2x", "Month View")
            GetForm_Tabs = GetForm_Tabs & Replace(Replace(buttonInactive, "x1x", FormYear), "x2x", "Year View")
        Case FormYear
            GetForm_Tabs = GetForm_Tabs & Replace(Replace(buttonInactive, "x1x", FormDay), "x2x", "Day View")
            GetForm_Tabs = GetForm_Tabs & Replace(Replace(buttonInactive, "x1x", FormWeek), "x2x", "Week View")
            GetForm_Tabs = GetForm_Tabs & Replace(Replace(buttonInactive, "x1x", FormMonth), "x2x", "Month View")
            GetForm_Tabs = GetForm_Tabs & Replace(Replace(buttonActive, "x1x", FormYear), "x2x", "Year View")
        Case FormDetails
            GetForm_Tabs = GetForm_Tabs & Replace(Replace(buttonInactive, "x1x", FormDay), "x2x", "Day View")
            GetForm_Tabs = GetForm_Tabs & Replace(Replace(buttonInactive, "x1x", FormWeek), "x2x", "Week View")
            GetForm_Tabs = GetForm_Tabs & Replace(Replace(buttonInactive, "x1x", FormMonth), "x2x", "Month View")
            GetForm_Tabs = GetForm_Tabs & Replace(Replace(buttonInactive, "x1x", FormYear), "x2x", "Year View")
        Case FormEdit
                GetForm_Tabs = GetForm_Tabs & Replace(Replace(buttonInactive, "x1x", FormDay), "x2x", "Day View")
            GetForm_Tabs = GetForm_Tabs & Replace(Replace(buttonInactive, "x1x", FormWeek), "x2x", "Week View")
            GetForm_Tabs = GetForm_Tabs & Replace(Replace(buttonInactive, "x1x", FormMonth), "x2x", "Month View")
            GetForm_Tabs = GetForm_Tabs & Replace(Replace(buttonInactive, "x1x", FormYear), "x2x", "Year View")
        Case Else
            GetForm_Tabs = GetForm_Tabs & Replace(Replace(buttonInactive, "x1x", FormDay), "x2x", "Day View")
            GetForm_Tabs = GetForm_Tabs & Replace(Replace(buttonInactive, "x1x", FormWeek), "x2x", "Week View")
            GetForm_Tabs = GetForm_Tabs & Replace(Replace(buttonActive, "x1x", FormMonth), "x2x", "Month View")
            GetForm_Tabs = GetForm_Tabs & Replace(Replace(buttonInactive, "x1x", FormYear), "x2x", "Year View")
        End Select
    '
    GetForm_Tabs = "" _
        & cr1 & "<div class=""calTabs"">" _
        & kmaIndent(GetForm_Tabs) _
        & cr1 & "</div>" _
        & ""
'    '
'    GetForm_Tabs = "" _
'        & cr1 & "<tr>" _
'        & kmaIndent(GetForm_Tabs) _
'        & cr1 & "</tr>" _
'        & ""
'    '
'    GetForm_Tabs = "" _
'        & cr1 & "<table class=""calTabTable"" border=0 cellspacing=0 cellpadding=0>" _
'        & kmaIndent(GetForm_Tabs) _
'        & cr1 & "</table>" _
'        & ""
    '
    Exit Function
ErrorTrap:
    Call HandleTrapError("GetForm_Tabs")
    End Function

'
'===================================================================================================================================
        '---/ Begin the Content Section
'===================================================================================================================================
'
        'GetForm_Tabs = GetForm_Tabs & ( tableWidth )
        
Private Function GetStartContentSection()
    On Error GoTo ErrorTrap
    '
    GetStartContentSection = GetStartContentSection & ("<table border=""0"" cellspacing=""0"" cellpadding=""0"" bgcolor=""" & BackGroundColor & """ width=""" & TableWidth & """>" & vbCrLf)
    GetStartContentSection = GetStartContentSection & ("<tr>" & vbCrLf)
    GetStartContentSection = GetStartContentSection & ("<td width=""2"" bgcolor=""white""><img src=""" & Main.ServerAppRootPath & "images/white.gif"" border=""0"" width=""2"" height=""100%""></td>")
    GetStartContentSection = GetStartContentSection & ("<td><font face=""verdana, arial"" size=""2"">" & vbCrLf)
    Exit Function
ErrorTrap:
    Call HandleTrapError("GetStartContentSection")
    End Function
'
'===================================================================================================================================
        '---/Procedure to Print Complete Year
'===================================================================================================================================
'
Private Function GetForm_Year() As String
    On Error GoTo ErrorTrap
    '
    Dim MonthPointer
    Dim sDate
    Dim DateNext As Date
    Dim DateLast As Date
    Dim MonthHasEvents As Boolean
    '
    DateNext = CDate(GlobalMonth & "/" & GlobalDay & "/" & CStr(GlobalYear)) + 365
    DateLast = CDate(GlobalMonth & "/" & GlobalDay & "/" & CStr(GlobalYear)) - 365
    '
    GetForm_Year = GetForm_Header(CStr(GlobalYear), FormYear, DateLast, DateNext)
    '
    GetForm_Year = GetForm_Year & "<TABLE border=""0"" cellPadding=""0"" cellSpacing=""0"" width=""100%"" >"
    For MonthPointer = 1 To 12
        sDate = CStr(MonthPointer) & "/1/" & GlobalYear
        If ((MonthPointer - 1) Mod iMonthCols = 0) Then
            GetForm_Year = GetForm_Year & "<TR>"
        End If
        '
        ' blocking events puts a javascript link up that works fine, but will block spiders
        '
        MonthHasEvents = False
        GetForm_Year = GetForm_Year & "<TD>" & Main.GetPanel(GetForm_Year_Month(sDate, MonthHasEvents)) & "</TD>"
        If (MonthPointer Mod iMonthCols = 0) Then
            GetForm_Year = GetForm_Year & ("</TR>")
            End If
        Next
    GetForm_Year = GetForm_Year & "</TABLE>"
    '
    GetForm_Year = Main.GetPanel(GetForm_Year)
    '
    '
    Exit Function
ErrorTrap:
    Call HandleTrapError("GetForm_Year")
    End Function

'================================================================================================
'---/ End Content Section
'=================================================================================================
'
Private Function GetEndContentSection()
    On Error GoTo ErrorTrap
    '
    GetEndContentSection = GetEndContentSection & ("</td>" & vbCrLf)
    GetEndContentSection = GetEndContentSection & ("<td width=""2"" bgcolor=""black""><img src=""" & Main.ServerAppRootPath & "images/black.gif"" border=""0"" width=""2"" height=""100%""></td>")
    GetEndContentSection = GetEndContentSection & ("</tr>" & vbCrLf)
    GetEndContentSection = GetEndContentSection & ("<tr>" & vbCrLf)
    GetEndContentSection = GetEndContentSection & ("<td width=""2"" bgcolor=""white""><img src=""" & Main.ServerAppRootPath & "images/white.gif"" border=""0"" width=""2"" height=""100%""></td>")
    GetEndContentSection = GetEndContentSection & ("<td align=center><br><br><br></td>")
    GetEndContentSection = GetEndContentSection & ("<td width=2 bgcolor=black><img src=black.gif border=0 width=2 height=""100%""></td>")
    GetEndContentSection = GetEndContentSection & ("</tr>" & vbCrLf)
    GetEndContentSection = GetEndContentSection & ("<tr>" & vbCrLf)
    GetEndContentSection = GetEndContentSection & ("<td colspan=3 height=2 bgcolor=black></td>")
    GetEndContentSection = GetEndContentSection & ("</tr>" & vbCrLf)
    GetEndContentSection = GetEndContentSection & ("</table>" & vbCrLf)
    '
    Exit Function
ErrorTrap:
    Call HandleTrapError("GetEndContentSection")
    End Function

'
'==================================================================================================================================
' New Print Monthly Calendar Populated Events
'==================================================================================================================================
'
Private Function GetForm_Month(calendarId As Long) As String
    On Error GoTo ErrorTrap
    '
    Dim EditLinkTemplate As String
    Dim Events() As Variant
    Dim EventCnt As Long
    '
    Dim CalendarEventsContentID As Long
    Dim DayNumberMonthStartsOn As Long
    Dim DayNumberPointer   As Long
    Dim CellDate As Date
    Dim EventLabels As String
    Dim Link As String
    Dim stream As String
    '
    Dim CellOpen As String
    Dim CellClose As String
    Dim ColumnPointer As Long
    Dim DateOfFirstCell As Date
    Dim DayNumberOfFirstCell As Long
    Dim DaysInMonth As Long
    Dim CellDayNumber As Long
    Dim DateNext As Date
    Dim DateLast As Date
    Dim CellCopy As String
    Dim CornerBox As String
    Dim AllowAdd As Boolean
    Dim TickStart As Long
    Dim s As New FastStringClass
    Dim startDate As Date
    Dim endDate As Date
    Dim TickData As Long
    Dim TickLoop As Long
    Dim CS As Long
    Dim c As FastStringClass
    Dim tabs As FastStringClass
    Dim linkHref As String
    Dim dateStyle As String
    '
    TickStart = GetTickCount()
    '
    AllowAdd = Main.IsLinkAuthoring("Calendar Events")
    CalendarEventsContentID = Main.GetContentID("Calendar Events")
    '
    If True Then
        '
        ' ----- header
        '
        DateNext = DateAdd("m", 1, GlobalDate)
        DateLast = DateAdd("m", -1, GlobalDate)
        s.Add GetForm_Header(MonthName(GlobalMonth) & ", " & GlobalYear, FormMonth, DateLast, DateNext)
        '
        s.Add cr1 & "<table class=""CalendarTable"" BORDER=""1"" CELLSPACING=""0"" CELLPADDING=""0"" bgcolor=""#ffffff"" width=""100%"">"
        '
        ' ----- spacers with day names
        '
        CellOpen = cr1 & "<td class=""CalendarDay"" ALIGN=""center"" width=""14%"">"
        CellClose = "" _
            & cr2 & "<br /><img src=""images/spacer.gif"" width=""70"" height=""1"">" _
            & cr1 & "</td>"
        s.Add cr1 & "<tr class=""CalendarDayRow"">"
        s.Add cr1 & "<TD class=""CalendarDay"" width=""1%"" align=""center"">wk<BR>&nbsp;view&nbsp;</td>"
        s.Add CellOpen & "Sun" & CellClose
        s.Add CellOpen & "Mon" & CellClose
        s.Add CellOpen & "Tue" & CellClose
        s.Add CellOpen & "Wed" & CellClose
        s.Add CellOpen & "Thu" & CellClose
        s.Add CellOpen & "Fri" & CellClose
        s.Add CellOpen & "Sat" & CellClose
        s.Add cr1 & "</tr>"
        
        '---/ Get day of the week the month starts on.
        '
        DayNumberMonthStartsOn = GetDayNumberMonthStartsOn(GlobalDate)
        DayNumberOfFirstCell = 1 - DayNumberMonthStartsOn
        DateOfFirstCell = CDate(GlobalMonth & "/1/" & GlobalYear) + DayNumberOfFirstCell
        '
        '---/ Write days of month in proper day slots
        '

        DaysInMonth = GetDaysInMonth(Month(GlobalDate), Year(GlobalDate))
        startDate = CDate(Month(GlobalDate) & "/1/" & Year(GlobalDate))
        endDate = GetNextMonth(startDate) - 1
        EventCnt = 0
        TickData = GetTickCount
        CS = OpenCSEvents2(startDate, endDate)
        If Main.IsCSOK(CS) Then
            Events = Main.GetCSRows(CS)
            EventCnt = UBound(Events, 2) + 1
        End If
        Call Main.CloseCS(CS)
        VerboseMessage = VerboseMessage & "<div>GetForm_Month, Data time = " & (GetTickCount - TickData) & "</div>"
        '
        EditLinkTemplate = ""
        If Main.IsEditing("Calendar Events") Then
            EditLinkTemplate = "<a href=""" & adminUrl & "?af=4&cid=" & Main.GetContentID("Calendar Events") & "&ID=xxx""><img src=""/cclib/images/IconContentEdit.gif"" width=18 height=22 border=0 style=""vertical-align:middle;""></a>"
        End If
        '
        TickLoop = GetTickCount
        For DayNumberPointer = DayNumberOfFirstCell To DaysInMonth
            Set c = New FastStringClass
            CellDate = CDate(Month(GlobalDate) & "/1/" & Year(GlobalDate)) + DayNumberPointer
            CellDayNumber = DayNumberPointer + 1
            '
            '--- Write TR and Week View Link if at the begining of the row
            '
            If ColumnPointer = 0 Then
                Link = RequestLink & rnTabNumber & "=" & FormWeek & "&" & RequestNameTargetdate & "=" & CellDate
                s.Add "<TR class=""CalendarCellRow"">"
                's.add  "<TD class=""ccCalendarMonth"" valign=""middle"">"
                s.Add "<TD class=""CalendarRowSelect"" style=""text-align: center;vertical-align:middle;height:40px;"" valign=""middle"">"
                s.Add "<a href=""javascript: document.location='" & Link & "';""><img src=""/ccLib/images/CalendarForwardButton.gif"" width=""16"" HEIGHT=""15"" border=""0"" alt=""Week Starting " & CellDate & """ valign=""middle""></a>"
                's.Add "<br><img src=""/ccLib/images/spacer.gif"" width=""1"" HEIGHT=""40"">"
                s.Add "</TD>"
            End If
            If (CellDayNumber < 1) Or (CellDayNumber > DaysInMonth) Then
                '
                ' Not in current month, set a blank cell
                '
                c.Add "&nbsp;"
                s.Add "<TD class=""CalendarCellBlank"" valign=""top"">" & c.Text & "</TD>"
            Else
                '
                ' Create cell and add the day number in the corner
                '
                Link = RequestLink & rnTabNumber & "=" & FormDay & "&" & RequestNameTargetdate & "=" & CellDate
                CornerBox = CellDayNumber
                If CellDayNumber < 10 Then
                    CornerBox = "&nbsp;" & CornerBox
                End If
                dateStyle = "CalendarDates"
                If (CellDate = GlobalDate) Then
                    CornerBox = "<div class=""CalendarActiveDate"">" & CornerBox & "</div>"
                End If
                '
                'EventLabels = GetForm_Month_Details(CellDate, Events, EventCnt, EditLinkTemplate)
                EventLabels = GetForm_Month_Details(CellDate, Events, EventCnt, EditLinkTemplate)
                If EventLabels = "" Then
                    '
                    ' no events, attempt to stop spiders
                    '
                    EventLabels = "<br><br>"
                    linkHref = "href=""javascript:document.location='" & Link & "'"""
                Else
                    '
                    ' no events
                    '
                    linkHref = "href=""" & Link & """"
                End If
                '
                c.Add cr1 & "<table class=""CalendarCellTable"" BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""1"" width=""100%"">"
                c.Add cr2 & "<tr>"
                c.Add cr3 & "<td align=""center"" width=""10%"" class=""CalendarCellCorner " & dateStyle & """>"
                c.Add cr4 & "<a " & linkHref & """>" & CornerBox & "</a>"
                'c.Add cr4 & "<a " & linkHref & """><span class=""CalendarDates"">" & CornerBox & "</span></a>"
                c.Add cr3 & "</td>"
                c.Add cr3 & "<td width=""90%"">&nbsp;</td>"
                c.Add cr2 & "</tr>"
                c.Add cr2 & "<tr>"
                c.Add cr3 & "<td colspan=""2"" class=""CalendarMonthEventLink"">" & EventLabels & "</td>"
                c.Add cr2 & "</tr>"
                c.Add cr1 & "</table>"
                If AllowAdd Then
                    c.Add GetAddLink(adminUrl, CalendarEventsContentID, "StartDate", Int(CellDate), "CalendarHolder", CStr(calendarId), "EndDate", Int(CellDate))
                End If
                s.Add "<TD class=""CalendarCell"" valign=""top"">" & c.Text & "</TD>"
            End If
            '
            '---/ If we're at the endof a row then close it up
            '
            ColumnPointer = ColumnPointer + 1
            If ColumnPointer >= 7 Then
                s.Add "</TR>"
                ColumnPointer = 0
            End If
        Next
        VerboseMessage = VerboseMessage & "<div>GetForm_Month, Time in loop = " & (GetTickCount - TickLoop) & "</div>"
        '
        '---/ Write spacer cells at end of last row if month doesn't end on a Saturday.
        '
        If ColumnPointer <> 0 Then
            Do While ColumnPointer < 7
                s.Add "<TD>&nbsp;</TD>"
                ColumnPointer = ColumnPointer + 1
            Loop
            s.Add "</TR>"
        End If
        '
        '---/ close up both the tables
        '
        '
    End If
    GetForm_Month = Main.GetPanel(s.Text & "</table>")
    VerboseMessage = VerboseMessage & "<div>GetForm_Month = " & (GetTickCount - TickStart) & "</div>"
    '
    Exit Function
ErrorTrap:
    Call HandleTrapError("GetForm_Month")
    End Function

'============================================================================================================================
        '---/ Procedure to Get Weekly Schedule Details for a Specific Date
        '---/Important Date value should be given correctly
'============================================================================================================================

Private Function GetForm_Week_Details2(DateValue As Date, Events As Variant, EventCnt As Long) As String
    On Error GoTo ErrorTrap
    '
    Dim eventID As Long
    Dim Temp As String
    Dim OverView As String
    Dim Filename As String
    Dim DetailPanel As String
    Dim stream As String
    Dim CSEvents As Long
    Dim eventName As String
    Dim HeaderPanel As String
    Dim ContentRecordKey As String
    Dim MoreRecords As Boolean
    Dim Delimiter As String
    Dim PageID As Long
    Dim AllowAdd As Boolean
    Dim EventPtr As Long
    Dim EventHits As Long
    Dim startDate As Date
    Dim endDate As Date
    '
    AllowAdd = Main.IsLinkAuthoring("Calendar Events")
    '
    HeaderPanel = "<a href=""javascript: document.location='" & RequestLink & rnTabNumber & "=" & FormDay & "&" & RequestNameTargetdate & "=" & DateValue & "';"">" & PrettyDate(DateValue) & "</a>"
    '
    EventHits = 0
    If EventCnt > 0 Then
        For EventPtr = 0 To EventCnt - 1
            If (KmaEncodeDate(Events(e_StartDate, EventPtr)) <= DateValue) And (KmaEncodeDate(Events(e_EndDate, EventPtr)) >= DateValue) Then
                eventName = kmaEncodeText(Events(e_Name, EventPtr))
                If (eventName <> "") Then
                    eventID = kmaEncodeInteger(Events(e_EventID, EventPtr))
                    PageID = kmaEncodeInteger(Events(e_PageID, EventPtr))
                    startDate = KmaEncodeDate(Events(e_StartDate, EventPtr))
                    endDate = KmaEncodeDate(Events(e_EndDate, EventPtr))
                    Filename = kmaEncodeText(Events(e_BriefFilename, EventPtr))
                    If Filename <> "" Then
                        OverView = Main.ReadVirtualFile(Filename)
                    End If
                    stream = stream & Delimiter
                    'Temp = "<a TabIndex=-1 href=""/admin/index.asp?cid=90&amp;id=11&amp;af=4&amp;aa=2&amp;ad=1""><img src=""/cclib/images/IconContentEdit.gif"" border=""0"" alt=""Edit Calendar Events [Event A]"" align=""absmiddle""></a>"
                    'stream = stream & Main.GetRecordEditLink("Calendar Events", eventID)
                    stream = stream & GetEditLink(eventID)
If False Then
                    stream = stream & GetCalendarEventLink3(Events, EventPtr, FormWeek, DateValue)
                    Delimiter = "<BR>"
Else
                    DetailPanel = ""
                    DetailPanel = DetailPanel & ("<table border=""0"" cellpadding=""3"" width=""100%"" bordercolordark=""#003399"">")
                    '
                    DetailPanel = DetailPanel & ("<tr><td align=""right"" class=""detailCaption""><b>Event Name:</b></td>")
                    DetailPanel = DetailPanel & ("<td align=""left"" class=""detailData"">" & GetCalendarEventLink3(Events, EventPtr, FormWeek, DateValue) & "</td></tr>")
                    DetailPanel = DetailPanel & ("<tr><td align=""right"" class=""detailCaption""><b>Start Date:</b></td><td align=""left"" class=""detailData"">" & startDate & "</td></tr>")
                    DetailPanel = DetailPanel & ("<tr><td align=""right"" class=""detailCaption""><b>End Date:</b></td><td align=""left"" class=""detailData"">" & endDate & "</td></tr>")
                    DetailPanel = DetailPanel & ("<tr><td align=""right"" class=""detailData""></td><td align=""left"" class=""detailData"">" & OverView & "</td></tr>")
                    DetailPanel = DetailPanel & ("<tr><td><img src=""/cclib/images/spacer.gif"" width=""100"" height=""2""></td><td width=""100%""><img src=""/cclib/images/spacer.gif"" width=1 height=2></td></tr>")
                    DetailPanel = DetailPanel & ("</TABLE>")
                    stream = stream & DetailPanel
End If
                    EventHits = EventHits + 1
                End If
            End If
        Next
    End If
    If EventHits = 0 Then
        stream = stream & "<div class=""detailCon"">There are no events recorded on this date</div>"
    End If
'    CSEvents = OpenCSEvents(DateValue, DateValue)
'    MoreRecords = Main.IsCSOK(CSEvents)
'    If Not MoreRecords Then
'        Stream = Stream & "There are no events recorded on this date"
'    Else
'        Do While MoreRecords
'            EventName = Main.GetCSText(CSEvents, "Name")
'            If (EventName <> "") Then
'                'ContentRecordKey = Main.GetCSText(CSEvents, "ContentRecordKey")
'                PageID = Main.GetCSInteger(CSEvents, "PageID")
'                Stream = Stream & Delimiter
'                Stream = Stream & Main.GetCSRecordEditLink(CSEvents)
'                'Stream = Stream & GetCalendarEventLink(ContentRecordKey, EventName, Main.GetCSDate(CSEvents, "StartDate"))
'                Stream = Stream & GetCalendarEventLink2(CSEvents)
'                Delimiter = "<BR>"
'            End If
'            '
'            Main.NextCSRecord (CSEvents)
'            MoreRecords = Main.IsCSOK(CSEvents)
'        Loop
'    End If
'    Call Main.CloseCS(CSEvents)
    '
    If AllowAdd Then
Dim ContentID As Long
        ContentID = Main.GetContentID("Calendar Events")
        stream = stream & "<br />" & GetAddLink(adminUrl, ContentID, "StartDate", Int(DateValue), "CalendarHolder", Main.GetRecordID("Calendars", GlobalCalendarName), "EndDate", Int(DateValue))
        End If
    '
    GetForm_Week_Details2 = HeaderPanel & Main.GetPanel(stream, "ccPanelHilite", "ccPanelShadow", "ccPanelHilite")
    '
    Exit Function
ErrorTrap:
    Call HandleTrapError("GetForm_Week_Details2")
    End Function
'
''============================================================================================================================
'        '---/ Procedure to Get Weekly Schedule Details for a Specific Date
'        '---/Important Date value should be given correctly
''============================================================================================================================
'
'Private Function GetForm_Week_Details(AdminURL As String, DateValue As Date, ContentName As String) As String
'    On Error GoTo ErrorTrap
'    '
'    Dim Stream As String
'    Dim CSEvents As Long
'    Dim EventName As String
'    Dim HeaderPanel As String
'    Dim ContentRecordKey As String
'    Dim MoreRecords As Boolean
'    Dim Delimiter As String
'    Dim PageID As Long
'    Dim AllowAdd As Boolean
'    '
'    AllowAdd = Main.IsLinkAuthoring("Calendar Events")
'    '
'    HeaderPanel = "<a href=""" & RequestLink & rnTabNumber & "=" & FormDay & "&" & RequestNameTargetdate & "=" & DateValue & """>" & PrettyDate(DateValue) & "</a>"
'    '
'    CSEvents = OpenCSEvents(DateValue, DateValue)
'    MoreRecords = Main.IsCSOK(CSEvents)
'    If Not MoreRecords Then
'        Stream = Stream & "There are no events recorded on this date"
'    Else
'        Do While MoreRecords
'            EventName = Main.GetCSText(CSEvents, "Name")
'            If (EventName <> "") Then
'                'ContentRecordKey = Main.GetCSText(CSEvents, "ContentRecordKey")
'                PageID = Main.GetCSInteger(CSEvents, "PageID")
'                Stream = Stream & Delimiter
'                Stream = Stream & Main.GetCSRecordEditLink(CSEvents)
'                'Stream = Stream & GetCalendarEventLink(ContentRecordKey, EventName, Main.GetCSDate(CSEvents, "StartDate"))
'                Stream = Stream & GetCalendarEventLink2(CSEvents)
'                Delimiter = "<BR>"
'            End If
'            '
'            Main.NextCSRecord (CSEvents)
'            MoreRecords = Main.IsCSOK(CSEvents)
'        Loop
'    End If
'    Call Main.CloseCS(CSEvents)
'    '
'    If AllowAdd Then
'        Stream = Stream & "<br />" & GetAddLink(AdminURL, "Calendar Events", "StartDate", Int(DateValue), "CalendarHolder", Main.GetRecordID("Calendars", GlobalCalendarName), "EndDate", Int(DateValue))
'        End If
'    '
'    GetForm_Week_Details = HeaderPanel & Main.GetPanel(Stream, "ccPanelHilite", "ccPanelShadow", "ccPanelHilite")
'    '
'    Exit Function
'ErrorTrap:
'    Call HandleTrapError("GetForm_Week_Details")
'    End Function


'============================================================================================================================
        '---/Display the Week Details
'============================================================================================================================

Private Function GetForm_Week(CalendarEventsContentID As Long) As String
    On Error GoTo ErrorTrap
    '
    Dim stream As String
    Dim DateValue As Date
    Dim DatePointer As Long
    Dim CS As Long
    Dim startDate As Date
    Dim endDate As Date
    Dim Events() As Variant
    Dim EventCnt As Long
    '
    DateValue = GlobalDate - Weekday(GlobalDate) + 1
    startDate = DateValue
    endDate = DateAdd("d", 7, startDate)
    EventCnt = 0
    CS = OpenCSEvents2(startDate, endDate)
    If Main.IsCSOK(CS) Then
        Events = Main.GetCSRows(CS)
        EventCnt = UBound(Events, 2) + 1
    End If
    Call Main.CloseCS(CS)
    '
    stream = stream & GetForm_Header("Week of " & GlobalDate, FormWeek, GlobalDate - 7, GlobalDate + 7)
    For DatePointer = 0 To 6
        stream = stream & Main.GetPanel(GetForm_Week_Details2(DateValue, Events, EventCnt))
        DateValue = DateAdd("d", 1, DateValue)
    Next
    GetForm_Week = Main.GetPanel(stream)
    '
    Exit Function
ErrorTrap:
    Call HandleTrapError("GetForm_Week")
    End Function

'=========================================================================================================================
        '---/Function to Display the Daily Events
'=========================================================================================================================

Private Function OpenCSEvents(DateStart As Date, DateEnd As Date)
    OpenCSEvents = OpenCSEvents2(DateStart, DateEnd)
    Exit Function
    On Error GoTo ErrorTrap
    '
    Dim SQL As String
    Dim calCriteria As String
    Dim locArray() As String
    Dim locArrayCount As Long
    Dim locPointer As Long
    
    If calSelected <> "" Then
        locArray = Split(calSelected, ",")
        locArrayCount = UBound(locArray())
    End If
    
    If locArrayCount > 0 Then
        For locPointer = 0 To locArrayCount
            If locArray(locPointer) <> "" Then
                If calCriteria <> "" Then
                    calCriteria = calCriteria & " OR  "
                End If
                calCriteria = calCriteria & "(ccCalendarEventRules.CalendarID=" & Main.GetRecordID("Calendars", locArray(locPointer)) & ")"
            End If
        Next
        
        If calCriteria = "" Then
            calCriteria = "(ccCalendarEventRules.CalendarID=" & KmaEncodeSQLNumber(GlobalCalendarID) & ")"
        End If
        
    Else
        calCriteria = "(ccCalendarEventRules.CalendarID=" & KmaEncodeSQLNumber(GlobalCalendarID) & ")"
    End If
    
    SQL = "SELECT ccCalendarEventRules.CalendarID, ccCalendarEvents.ID, ccCalendarEvents.StartDate, ccCalendarEvents.EndDate, ccCalendarEvents.ContentRecordKey, ccCalendarEvents.Name, ccCalendarEvents.PageID, ccCalendarEvents.ContentControlID, ccCalendarEvents.CopyFilename, ccCalendarEvents.EventLink, ccCalendarEvents.BriefFilename" _
        & " FROM ccCalendarEventRules LEFT JOIN ccCalendarEvents ON ccCalendarEventRules.CalendarEventID = ccCalendarEvents.ID" _
        & " WHERE (" _
        & "(" & calCriteria & ")" _
        & "AND(ccCalendarEvents.StartDate<" & KmaEncodeSQLDate(Int(DateStart + 1)) & ")" _
        & "AND(ccCalendarEvents.EndDate>=" & KmaEncodeSQLDate(Int(DateEnd)) & ")" _
        & "AND(ccCalendarEventRules.Active<>0)" _
        & "AND(ccCalendarEvents.Active<>0));"
    OpenCSEvents = Main.OpenCSSQL("default", SQL)
Call Main.TestPoint("SQL [" & DateStart & "] - [" & DateEnd & "]: " & SQL)
    '
    Exit Function
ErrorTrap:
    Call HandleTrapError("OpenCSEvents")
    End Function

'=========================================================================================================================
'   Select all events that have days in the period
'=========================================================================================================================

Private Function OpenCSEvents2(DateStart As Date, DateEnd As Date)
    On Error GoTo ErrorTrap
    '
    Dim SQL As String
    Dim calCriteria As String
    Dim locArray() As String
    Dim locArrayCount As Long
    Dim locPointer As Long
    
    If calSelected <> "" Then
        locArray = Split(calSelected, ",")
        locArrayCount = UBound(locArray())
    End If
    
    If locArrayCount > 0 Then
        For locPointer = 0 To locArrayCount
            If locArray(locPointer) <> "" Then
                If calCriteria <> "" Then
                    calCriteria = calCriteria & " OR  "
                End If
                calCriteria = calCriteria & "(ccCalendarEventRules.CalendarID=" & Main.GetRecordID("Calendars", locArray(locPointer)) & ")"
            End If
        Next
        
        If calCriteria = "" Then
            calCriteria = "(ccCalendarEventRules.CalendarID=" & KmaEncodeSQLNumber(GlobalCalendarID) & ")"
        End If
        
    Else
        calCriteria = "(ccCalendarEventRules.CalendarID=" & KmaEncodeSQLNumber(GlobalCalendarID) & ")"
    End If
    
    SQL = "SELECT ccCalendarEvents.StartDate, ccCalendarEvents.EndDate, ccCalendarEvents.Name, ccCalendarEvents.PageID, ccCalendarEvents.CopyFilename, ccCalendarEvents.EventLink, ccCalendarEvents.ID, ccCalendarEvents.BriefFilename" _
        & " FROM ccCalendarEvents" _
        & " LEFT JOIN ccCalendarEventRules ON ccCalendarEventRules.CalendarEventID = ccCalendarEvents.ID" _
        & " WHERE " _
        & "((" _
            & "(ccCalendarEventRules.id is null)" _
        & ")or(" _
            & "(" & calCriteria & ")" _
            & "AND(ccCalendarEventRules.Active<>0)" _
        & "))" _
        & "AND(ccCalendarEvents.StartDate<" & KmaEncodeSQLDate(Int(DateEnd + 1)) & ")" _
        & "AND(ccCalendarEvents.EndDate>=" & KmaEncodeSQLDate(Int(DateStart)) & ")" _
        & "AND(ccCalendarEvents.Active<>0)" _
        & ";"
    OpenCSEvents2 = Main.OpenCSSQL("default", SQL)
    '
    Exit Function
ErrorTrap:
    Call HandleTrapError("OpenCSEvents")
    End Function
''
''=========================================================================================================================
''---/Important Function
''=========================================================================================================================
''
'Private Function Work_Days(BegDate, EndDate)
'    On Error GoTo ErrorTrap
'    '
'    Dim WholeWeeks
'    Dim DateCnt
'    Dim EndDays
'    '
'    BegDate = DateValue(BegDate)
'    EndDate = DateValue(EndDate)
'    WholeWeeks = DateDiff("w", BegDate, EndDate)
'    DateCnt = DateAdd("ww", WholeWeeks, BegDate)
'    EndDays = 0
'    Do While DateCnt <= EndDate
'        If Weekday(DateCnt) <> "1" And _
'            Weekday(DateCnt) <> "7" Then
'            EndDays = EndDays + 1
'        End If
'        DateCnt = DateAdd("d", 1, DateCnt)
'    Loop
'    Work_Days = WholeWeeks * 7 + EndDays
'    '
'    Exit Function
'ErrorTrap:
'    Call HandleTrapError("Work_Days")
'    End Function
'
'
''=========================================================================================================================
'        '---/ To Display all the Events for that Day (More then 1 Occurence Events)
''=========================================================================================================================
'Private Function GetAllDayEventRule(ContentName As String) As String
'    On Error GoTo ErrorTrap
'    '
'    Dim DateStart As Date
'    Dim CSCalendarEventRules As Long
'    '
'    DateStart = CDate(GlobalMonth & "/" & GlobalDay & "/" & GlobalYear)
'    '
'    CSCalendarEventRules = Main.OpenCSContent_Internal(ContentName, "(StartDate<=" & KmaEncodeSQLDate(DateStart) & ") AND ( EndDate>=" & KmaEncodeSQLDate(DateStart) & " )", "StartDate")
'    Do While Main.CSOK(CSCalendarEventRules)
'        GetAllDayEventRule = GetAllDayEventRule _
'            & "<a href=""" & RequestLink & "id=" & Main.GetCSField(CSCalendarEventRules, "ID") & "&" & RequestNameTargetdate & "=" & GlobalDate & """>" _
'            & Main.GetCSField(CSCalendarEventRules, "Name") _
'            & "</a>" _
'            & "<br>"
'        Main.NextCSRecord (CSCalendarEventRules)
'    Loop
'    Call Main.CloseCS(CSCalendarEventRules)
'    '
'    Exit Function
'ErrorTrap:
'    Call HandleTrapError("GetAllDayEventRule")
'    End Function
'
'===============================================================================================================================================
'---/Procedure to print the Event in Calendar.asp Page
'================================================================================================================================================
'
Private Function GetForm_Day(CalendarEventsContentID As Long)
    On Error GoTo ErrorTrap
    '
    Dim calendarId As Long
    Dim CSEvents As Long
    Dim stream As String
    Dim DetailPanel As String
    Dim HeaderPanel As String
    Dim PageID As Long
    Dim ContentRecordKey As String
    Dim AllowAdd As Boolean
    '
    AllowAdd = Main.IsLinkAuthoring("Calendar Events")
    calendarId = Main.GetRecordID("Calendars", GlobalCalendarName)
    '
    ' ----- Print Table Header
    '
    HeaderPanel = HeaderPanel & GetForm_Header(FormatDateTime(GlobalDate, vbLongDate), FormDay, GlobalDate - 1, GlobalDate + 1)
    DetailPanel = DetailPanel & ("<table border=""1""  cellspacing=""0"" cellpadding=""0"" width=""100%"" bordercolordark=""#ffffff"">")
    '
    ' ----- display daily events
    '
    CSEvents = OpenCSEvents(Int(GlobalDate), Int(GlobalDate))
    If Not Main.CSOK(CSEvents) Then
        DetailPanel = DetailPanel & ("<tr><td width=""100%"" class=""detailCon"">There are no events recorded on this date</td></tr>")
    Else
        Do While Main.CSOK(CSEvents)
            DetailPanel = DetailPanel & ("<tr><td width=""100%"" class=""detailCon"">")
            DetailPanel = DetailPanel & ("<table border=""0"" cellpadding=""3"" width=""100%"" bordercolordark=""#003399"">")
            DetailPanel = DetailPanel & ("<tr><td align=""right"" class=""detailCaption""><b>Event Name:</b></td>")
            DetailPanel = DetailPanel & ("<td align=""left"" class=""detailData"">")
            '
            'DetailPanel = DetailPanel & Main.GetCSRecordEditLink(CSEvents)
            DetailPanel = DetailPanel & GetEditLink(Main.GetCSInteger(CSEvents, "ID"))
            PageID = Main.GetCSInteger(CSEvents, "PageID")
            '
            'DetailPanel = DetailPanel & GetCalendarEventLink(Main.GetCSText(CSEvents, "ContentRecordKey"), Main.GetCSText(CSEvents, "Name"), GlobalDate)
            'DetailPanel = DetailPanel & GetCalendarEventLink(ContentRecordKey, Main.GetCSText(CSEvents, "Name"), GlobalDate)
            DetailPanel = DetailPanel & GetCalendarEventLink2(CSEvents)
            DetailPanel = DetailPanel & ("</td></tr>")
            DetailPanel = DetailPanel & ("<tr><td align=""right"" class=""detailCaption""><b>Start Date:</b></td>")
            DetailPanel = DetailPanel & ("<td align=""left"" class=""detailData"">" & Main.GetCSField(CSEvents, "StartDate") & "</td></tr>")
            DetailPanel = DetailPanel & ("<tr><td align=""right"" class=""detailCaption""><b>End Date:</b></td>")
            DetailPanel = DetailPanel & ("<td align=""left"" class=""detailData"">" & Main.GetCSField(CSEvents, "EndDate") & "</tr>")
            DetailPanel = DetailPanel & ("<tr><td align=""right"" class=""detailCaption""></td>")
            DetailPanel = DetailPanel & ("<td align=""left"" class=""detailData"">" & Main.GetCSTextFile(CSEvents, "BriefFileName") & "</tr>")
            'DetailPanel = DetailPanel & ("<tr><td><img src=""images/spacer.gif"" width=""100"" height=""2""></td>")
            'DetailPanel = DetailPanel & ("<td width=""100%""></td></tr>")
            DetailPanel = DetailPanel & ("</TABLE>")
            DetailPanel = DetailPanel & ("</td></tr>")
            Main.NextCSRecord (CSEvents)
        Loop
    End If
    Call Main.CloseCS(CSEvents)
    DetailPanel = DetailPanel & "</table>"
    If AllowAdd Then
        DetailPanel = DetailPanel & GetAddLink(adminUrl, CalendarEventsContentID, "StartDate", Int(GlobalDate), "CalendarHolder", CStr(calendarId), "EndDate", Int(GlobalDate))
        End If
    '
    stream = Main.GetPanel(DetailPanel, "ccPanelHilite", "ccPanelShadow", "ccPanelHilite")
    GetForm_Day = Main.GetPanel(HeaderPanel & stream)
    '
    Exit Function
ErrorTrap:
    Call HandleTrapError("GetForm_Day")
    End Function
'
'Check the array for a meeting on given date
'If a meeting exists, return the position in the array
'
Private Function HasMeeting(dDate)
    On Error GoTo ErrorTrap
    '
        If IsEmpty(aryMeetings) Then
            HasMeeting = -1
        Else
            For i = 0 To UBound(aryMeetings, 2)
                If CDate(aryMeetings(1, i)) = dDate Then
                    HasMeeting = i
                    Exit Function
                End If
            Next
            HasMeeting = -1
        End If
    '
    Exit Function
ErrorTrap:
    Call HandleTrapError("HasMeeting")
    End Function
'
'Write the cell with the correct colors and applicable link
'
Private Function GetForm_Year_Month_Day(iCellPos, dDate, HasEvents As Boolean)
    On Error GoTo ErrorTrap
    '
    Dim sHTMLOut
    Dim lMeetingId
    Dim DayOfWeek As Long
    Dim DayOfMonth As Long
    Dim Link As String
        
    sHTMLOut = ""
    '
    'If this is Sunday, start a new <tr>
    '
    If (iCellPos Mod 7 = 1) Then
        sHTMLOut = sHTMLOut & "<tr>"
    End If
    '
    If IsNull(dDate) Then
        '
        'Write an empty cell
        '
        sHTMLOut = sHTMLOut & "<TD >&nbsp;</TD>"
    Else
        '
        'Check to see if date is a weekend
        '
        DayOfMonth = DatePart("d", dDate)
        sHTMLOut = sHTMLOut & "<TD align=""center"">"
        If DayOfMonth < 10 Then
            sHTMLOut = sHTMLOut & "&nbsp;&nbsp;"
        End If
        Link = RequestLink & rnTabNumber & "=" & FormDay & "&" & RequestNameTargetdate & "=" & dDate
        If Not ((dDate > GlobalDateLimitPrevious) And (dDate < GlobalDateLimitNext)) Then
            sHTMLOut = sHTMLOut & DayOfMonth
        ElseIf HasEvents Then
            sHTMLOut = sHTMLOut & "<a href=""" & Link & """>" & DayOfMonth & "</a>"
        Else
            sHTMLOut = sHTMLOut & "<a href=""javascript: document.location='" & Link & "';"">" & DayOfMonth & "</a>"
        End If
        sHTMLOut = sHTMLOut & "</TD>"
    End If
    '
    'If this is Saturday, end the <tr>
    '
    If iCellPos Mod 7 = 0 Then
        sHTMLOut = sHTMLOut & "</tr>"
    End If
    GetForm_Year_Month_Day = sHTMLOut
    '
    Exit Function
ErrorTrap:
    Call HandleTrapError("GetForm_Year_Month_Day")
    End Function
'
'Write the month table
'
Private Function GetForm_Year_Month(sDate, MonthHasEvents As Boolean) As String
    On Error GoTo ErrorTrap
    '
    Dim sHTMLOut
    Dim dValue
    Dim iStartDay
    Dim iMonth
    Dim iDateIncrement
    Dim TargetDate As Date
    '
    sHTMLOut = ""
    dValue = CDate(sDate)
    iStartDay = DatePart("w", dValue)
    iMonth = DatePart("m", dValue)
    iDateIncrement = 0
    '
    'Write the month table header
    '
    sHTMLOut = sHTMLOut & "<TABLE border=0 cellPadding=0 cellSpacing=1 bgcolor=white width=""100%"">"
    sHTMLOut = sHTMLOut & "<TR>"
    sHTMLOut = sHTMLOut & "<TD class=""CalendarMonthHeader"" align=middle colSpan=7>"
    TargetDate = CDate(Month(sDate) & "/" & Day(sDate) & "/" & Year(sDate))
    If (TargetDate < GlobalDateLimitNext) And (TargetDate > GlobalDateLimitPrevious) Then
        sHTMLOut = sHTMLOut & "<a class=""CalendarMonthName"" href=""javascript:document.location='" & RequestLink & rnTabNumber & "=" & FormMonth & "&" & RequestNameTargetdate & "=" & CStr(TargetDate) & "';"">" & MonthName(iMonth) & "</A>"
    Else
        sHTMLOut = sHTMLOut & MonthName(iMonth)
    End If
    sHTMLOut = sHTMLOut & "</TD>"
    '
    sHTMLOut = sHTMLOut & "<TR>"
    sHTMLOut = sHTMLOut & "<TD align=""center"" class=""CalendarDayHeader"">S</TD>"
    sHTMLOut = sHTMLOut & "<TD align=""center"" class=""CalendarDayHeader"">M</TD>"
    sHTMLOut = sHTMLOut & "<TD align=""center"" class=""CalendarDayHeader"">T</TD>"
    sHTMLOut = sHTMLOut & "<TD align=""center"" class=""CalendarDayHeader"">W</TD>"
    sHTMLOut = sHTMLOut & "<TD align=""center"" class=""CalendarDayHeader"">T</TD>"
    sHTMLOut = sHTMLOut & "<TD align=""center"" class=""CalendarDayHeader"">F</TD>"
    sHTMLOut = sHTMLOut & "<TD align=""center"" class=""CalendarDayHeader"">S</TD></TR>"
    '
    'Step through each cell in the month table
    '
    For iCell = 1 To 42
        '
        'Check to see if we can start writing days
        '
        If bValidCell Or (iCell = iStartDay) Then
            bValidCell = True
            '
            'Bump up the date
            '
            dValue = CDate(sDate) + iDateIncrement
            iDateIncrement = iDateIncrement + 1
            '
            'If we're not in the current month, we've gone too far
            '
            If DatePart("m", dValue) <> iMonth Then
                '
                'Invalid cell, write a blank cell
                '
                bValidCell = False
                sHTMLOut = sHTMLOut & GetForm_Year_Month_Day(iCell, Null, MonthHasEvents)
            Else
                '
                'Still in this month, write the day
                '
                sHTMLOut = sHTMLOut & GetForm_Year_Month_Day(iCell, dValue, MonthHasEvents)
            End If
        Else
            '
            'We're either before or after the month day range,
            'just write a blank cell
            '
            sHTMLOut = sHTMLOut & GetForm_Year_Month_Day(iCell, Null, MonthHasEvents)
        End If
    Next
    
    'Finish up the month table
    sHTMLOut = sHTMLOut & "</TABLE>"
    GetForm_Year_Month = sHTMLOut
    '
    Exit Function
ErrorTrap:
    Call HandleTrapError("GetForm_Year_Month")
    End Function

        

Private Sub Class_Terminate()
    Set Main = Nothing
    End Sub
'
'
'
Private Function GetForm_Header(Label As String, FormID As Long, DateBack As Date, DateNext As Date) As String
    On Error GoTo ErrorTrap
    '
    GetForm_Header = GetForm_Header & "<p align=""center"">"
    If (DateBack < GlobalDateLimitPrevious) Then
        GetForm_Header = GetForm_Header & ("<img src=""/ccLib/images/CalendarBackButton.gif"" width=""16"" HEIGHT=""15"" border=""0"" alt=""Previous"">")
    Else
        GetForm_Header = GetForm_Header & "<a href=""javascript: document.location='" & RequestLink & rnTabNumber & "=" & FormID & "&" & RequestNameTargetdate & "=" & DateBack & "';"">"
        GetForm_Header = GetForm_Header & ("<img src=""/ccLib/images/CalendarBackButton.gif"" width=""16"" HEIGHT=""15"" border=""0"" alt=""Previous"">")
        GetForm_Header = GetForm_Header & ("</a>")
    End If
    GetForm_Header = GetForm_Header & ("&nbsp;&nbsp;<B>" & Label & "</B>&nbsp;&nbsp;")
    If (DateNext > GlobalDateLimitNext) Then
        GetForm_Header = GetForm_Header & ("<img src=""/ccLib/images/CalendarForwardButton.gif"" width=""16"" HEIGHT=""15"" border=""0"" alt=""Next"">")
    Else
        GetForm_Header = GetForm_Header & ("<a href=""javascript:document.location='" & RequestLink & rnTabNumber & "=" & FormID & "&" & RequestNameTargetdate & "=" & DateNext & "';"">")
        GetForm_Header = GetForm_Header & ("<img src=""/ccLib/images/CalendarForwardButton.gif"" width=""16"" HEIGHT=""15"" border=""0"" alt=""Next"">")
        GetForm_Header = GetForm_Header & ("</a>")
    End If
    '
    Exit Function
ErrorTrap:
    Call HandleTrapError("GetForm_Header")
    End Function
'
'
'
Private Function GetForm_PageContent()
    '
    Exit Function
ErrorTrap:
    Call HandleTrapError("GetForm_PageContent")
    End Function
'
'
'
Private Sub HandleTrapError(MethodName As String)
    Set Main = Nothing
    Call HandleError("Calendar", MethodName, Err.Number, Err.Source, Err.Description, True, False)
    End Sub
'
'
'
Private Sub HandleInternalError(MethodName As String, ErrDescription As String)
    '
    Call Err.Raise(KmaErrorInternal, App.EXEName, ErrDescription)
    End Sub
'
Private Sub UpgradeCalendarEvents()
    On Error GoTo ErrorTrap
    '
    Dim CS As Long
    Dim cs2 As Long
    Dim calendarEventId As Long
    Dim calendarId As Long
    Dim CurrentVersion As String
    Dim SQL As String
    Dim CSSQL As Long
    Dim CSPointer As Long
    Dim SepPos As Long
    Dim CRKey As String
    Dim PageID As Long
    '
    CurrentVersion = Main.GetSiteProperty("aoCalendarVersion", "0.0.0")
    '
    If CurrentVersion <> AddOnVersion Then
        '
        If CurrentVersion < "1.0.0" Then
            '
            SQL = "SELECT CE.ID " _
                & "FROM ccCalendarEvents CE, ccCalendarEventRules CR " _
                & "Where (CE.ID = CR.CalendarEventID) " _
                & "AND (CE.ContentRecordKey is not null) " _
                & "AND (CE.ContentControlID=" & Main.GetContentID("Calendar Events") & ") " _
                & "AND (CR.ContentControlID=" & Main.GetContentID("Calendar Event Rules") & ") " _
            '
            CSSQL = Main.OpenCSSQL("Default", SQL)
            Do While Main.CSOK(CSSQL)
                CSPointer = Main.OpenCSContent("Calendar Events", "ID=" & Main.GetCSInteger(CSSQL, "ID"))
                If Main.CSOK(CSPointer) Then
                    CRKey = Main.GetCSText(CSPointer, "ContentRecordKey")
                    SepPos = InStr(1, CRKey, ".", vbTextCompare)
                    If SepPos <> 0 Then
                        PageID = Right(CRKey, Len(CRKey) - SepPos)
                        Call Main.SetCS(CSPointer, "PageID", PageID)
                        If PageID <> 0 Then
                            Call Main.SetCS(CSPointer, "EventLink", Main.GetPageLink(PageID))
                        End If
                        Call Main.SaveCSRecord(CSPointer)
                    End If
                End If
                Call Main.CloseCS(CSPointer)
                Call Main.NextCSRecord(CSSQL)
            Loop
            Call Main.CloseCS(CSSQL)
            '
        End If
        '
        If CurrentVersion < "1.0.48" Then
                CS = Main.OpenCSSQL("default", "select e.id from ccCalendarEvents e left join ccCalendarEventRules r on r.calendareventId=e.id where r.id is null")
                If Main.IsCSOK(CS) Then
                    cs2 = Main.InsertCSRecord("calendars")
                    If Main.IsCSOK(cs2) Then
                        Call Main.SetCS(cs2, "name", "Hidden Calendar Events")
                        calendarId = Main.GetCSInteger(cs2, "id")
                    End If
                    Call Main.CloseCS(cs2)
            If calendarId <> 0 Then
                    Do While Main.IsCSOK(CS)
                        calendarEventId = Main.GetCSInteger(CS, "id")
                        cs2 = Main.InsertCSRecord("Calendar Event Rules")
                        If Main.IsCSOK(cs2) Then
                            Call Main.SetCS(cs2, "calendarId", calendarId)
                            Call Main.SetCS(cs2, "calendarEventId", calendarEventId)
                        End If
                        Call Main.CloseCS(cs2)
                        Call Main.NextCSRecord(CS)
                    Loop
            End If
                End If
            '
        End If
        '
        Call Main.SetSiteProperty("aoCalendarVersion", AddOnVersion)
        '
    End If
    '
    Exit Sub
ErrorTrap:
    Call HandleTrapError("UpgradeCalendarEvents")
    End Sub
'
Private Function GetAddLink(adminUrl As String, ContentID As Long, Ignore0 As String, startDate As String, Ignore1 As String, calendarId As String, Ignore2 As String, endDate As String) As String
    
    On Error GoTo ErrorTrap
    '
    Dim stream As String
    Dim Link As String
    '
    'Link = AdminURL & "?cid=" & ContentID
    'Link = Link & "&af=4&aa=2&ad=1"
    
    Link = RequestLink
    Link = Link & "&" & rnTabNumber & "=" & FormEdit
    Link = Link & "&" & rnDefaultEventCalendarID & "=" & calendarId
    Link = Link & "&" & rnCalStart & "=" & startDate
    '
    Link = Main.EncodeURL(Link)
    '
    GetAddLink = "<a href=""" & Link & """><img border=""0"" src=""/cclib/images/IconContentAdd.gif""></a>"
    '
    Exit Function
ErrorTrap:
    Call HandleError("CalendarClass", "GetAddLink", Err.Number, Err.Source, Err.Description, True, False)
    End Function
    
Private Function GetEditLink(eventID) As String
    On Error GoTo ErrorTrap
    '
    Dim stream As String
    Dim Link As String
    
    If AllowEdit Then
        Link = RequestLink
        Link = Link & "&" & rnTabNumber & "=" & FormEdit
        Link = Link & "&" & rnEventID & "=" & eventID
        
        Link = Main.EncodeURL(Link)
        
        GetEditLink = "<a href=""" & Link & """><img border=""0"" src=""/cclib/images/IconContentEdit.gif""></a>"
    End If
    
    Exit Function
ErrorTrap:
    Call HandleError("CalendarClass", "GetEditLink", Err.Number, Err.Source, Err.Description, True, False)
    End Function

Private Function GetForm_Details()
    On Error GoTo ErrorTrap
    '
    Dim CSEvents As Long
    Dim stream As String
    Dim PageID As Long
    Dim eventID As Long
    Dim eventName As String
    '
    eventID = Main.GetStreamInteger(RequestNameEventID)
    '
    CSEvents = Main.OpenCSContent("Calendar Events", "ID=" & Main.EncodeSQLNumber(eventID))
    If Not Main.CSOK(CSEvents) Then
        stream = stream & ("<span class=""ccError"">The selected event is currently unavailable</span>")
    Else
        eventName = Main.GetCSText(CSEvents, "Name")
        If eventName = "" Then
            eventName = "Event " & eventID
        End If
        Call Main.TrackContent("Calendar Events", eventID)
        stream = stream & Main.GetSiteProperty("Calendar Breadcrumb Prefix", "return to ") & " <a href=""" & RequestLink & "&" & RequestNameTargetdate & "=" & kmaEncodeRequestVariable(CStr(GlobalDate)) & """>Calendar</a><br /><br />"
        'stream = stream & Main.GetCSRecordEditLink(CSEvents)
        stream = stream & GetEditLink(eventID)
        stream = stream & "<b>Event Name:</b> " & eventName & "<br />"
        stream = stream & "<b>Start Date:</b> " & Main.GetCSField(CSEvents, "StartDate") & "<br />"
        stream = stream & "<b>End Date:</b> " & Main.GetCSField(CSEvents, "EndDate") & "<br /><br />"
        stream = stream & Main.GetCSTextFile(CSEvents, "copyFileName")
    End If
    Call Main.CloseCS(CSEvents)
    '
    GetForm_Details = stream & "<!-- noWrap -->"
    '
    Exit Function
ErrorTrap:
    Call HandleTrapError("GetForm_Details")
    End Function

Private Function getNavigation()
    On Error GoTo ErrorTrap

    Dim stream As String
    Dim locPointer As Long
    
    If calArrayCount > 0 Then
        stream = "<div id=""calendarNavigation"">"
        stream = stream & "<div id=""calendarInstructions"">" & Main.GetContentCopy2("Calendar Filter Copy", , "Use the fields bellow to filter calendar events.") & "</div>"
        stream = stream & Main.GetFormStart()
        For locPointer = 0 To calArrayCount
            stream = stream & "<div class=""calSelectContainer"">"
            stream = stream & "<div class=""selectLeft"">" & Main.GetFormInputCheckBox(rnCalendar & "_" & locPointer, InStr(1, calSelected, "," & calArray(locPointer) & ",", vbTextCompare)) & "</div>"
            stream = stream & "<div class=""selectRight""><img src=""/ccLib/images/spacer.gif"" width=""10"" height=""2""><br />" & calArray(locPointer) & "</div>"
            stream = stream & "</div>"
        Next
        stream = stream & "<div class=""calButtonContainer"">" & Main.GetFormButton(buttonUpdate) & "</div>"
        stream = stream & Main.GetFormEnd
        stream = stream & "</div>"
    End If
    
    getNavigation = stream

    Exit Function
ErrorTrap:
    Call HandleTrapError("getNavigation")
    End Function

Private Sub setSelectedCalendars()
    On Error GoTo ErrorTrap

    Dim locPointer As Long
    Dim locSelcted As String
    
    If Main.GetStreamText("button") = buttonUpdate Then
        If calArrayCount > 0 Then
            For locPointer = 0 To calArrayCount
                If Main.GetStreamBoolean(rnCalendar & "_" & locPointer) Then
                    locSelcted = locSelcted & "," & calArray(locPointer)
                End If
            Next
        End If
        
        Call Main.SetMemberProperty("Calendars Selected", locSelcted & ",")
    End If

    Exit Sub
ErrorTrap:
    Call HandleTrapError("getNavigation")
    End Sub

Private Function getFormEdit(calendarId As Long)
    On Error GoTo ErrorTrap

    Dim CS As Long
    Dim stream As String
    Dim hdnList As String
    Dim btnList As String
    Dim cnt
    Dim calSelect As String
    '
    CS = Main.OpenCSSQL("", "select count(*) as cnt from ccCalendars where active<>0")
    If Main.IsCSOK(CS) Then
        cnt = Main.GetCSInteger(CS, "cnt")
    End If
    Call Main.CloseCS(CS)
    '
    If cnt = 0 Then
        '
        ' no calendars, can not add an event
        '
        getFormEdit = "<p>There are currently no calendars available to select.</p>"
        Exit Function
    ElseIf cnt = 1 Then
        '
        CS = Main.OpenCSSQL("", "select name,id from ccCalendars where active<>0")
        If Main.IsCSOK(CS) Then
            calSelect = "" _
                & "<br>" _
                & "<input type=""hidden"" value=""" & Main.GetCSInteger(CS, "id") & """ name=""cals.0.ID"">" _
                & "<input type=""checkbox"" name=""cals.0"" checked=""checked"" disabled=""disabled"">" _
                & "<span class=""ccAdminNormal"">" & Main.GetCSText(CS, "name") & "</span>" _
                & ""
        End If
        Call Main.CloseCS(CS)
    Else
        calSelect = Main.GetFormInputCheckList(rnCals, "Calendar Events", eventID, "Calendars", "Calendar Event Rules", "CalendarEventID", "CalendarID")
        '
        If eventID = 0 Then
            '
            ' try to default the last calendar this person used on, else use default
            '
            calSelect = Replace(calSelect, "name=""cals." & calendarId & """", "checked=""checked"" name=""" & calendarId & """", , , vbTextCompare)
'            calendarId = Main.GetMemberProperty("CalendarId last selected", "0")
'            If calendarId = 0 Then
'                calendarId = Main.GetRecordID("calendars", "default")
'            End If
'            If calendarId <> 0 Then
'                calendarId = calendarId
'            End If
        End If
        calSelect = "<br>" & Replace(calSelect, "<br />", "", , , vbTextCompare)
    End If
    '
    btnList = Main.GetFormButton(btnSave) & " " & Main.GetFormButton(btnCancel) & " " & Main.GetFormButton(btnDelete)
    If ((eventID <> 0) And (calRepeats <> 0)) Or (calPrntID <> 0) Then
        btnList = btnList & "<br /><img src=""/ccLib/images/spacer.gif"" height=""5"" width=""10""><br />" & Main.GetFormButton(btnSaveSeries) & " " & Main.GetFormButton(btnDeleteSeries)
    End If
    
    hdnList = Main.GetFormInputHidden(rnEventID, eventID, rnEventID)
    hdnList = hdnList & Main.GetFormInputHidden(rnCalORName, calORName, rnCalORName)
    hdnList = hdnList & Main.GetFormInputHidden(rnCalORStart, calORStart, rnCalORStart)
    hdnList = hdnList & Main.GetFormInputHidden(rnCalOREnd, calOREnd, rnCalOREnd)
    hdnList = hdnList & Main.GetFormInputHidden(rnCalORURL, calORURL, rnCalORURL)
    hdnList = hdnList & Main.GetFormInputHidden(rnCalORDayCopy, calORDayCopy, rnCalORDayCopy)
    hdnList = hdnList & Main.GetFormInputHidden(rnCalORDetailCopy, calORDetailCopy, rnCalORDetailCopy)
    'hdnList = hdnList & Main.GetFormInputHidden(rnCals, "")
    hdnList = hdnList & Main.GetFormInputHidden(rnCalORRepeats, calORRepeats, rnCalORRepeats)
    hdnList = hdnList & Main.GetFormInputHidden(rnCalORRptStart, calORRptStart, rnCalORRptStart)
    hdnList = hdnList & Main.GetFormInputHidden(rnCalORRptEnd, calORRptEnd, rnCalORRptEnd)
    hdnList = hdnList & Main.GetFormInputHidden(rnCalPrntID, calPrntID, rnCalPrntID)
    hdnList = hdnList & Main.GetFormInputHidden(rnDefaultEventCalendarID, calDefaultEventCalendarID)
    
    stream = Main.GetFormStart
    stream = stream & "<div id=""calEditContainer"">"
    stream = stream & "<div class=""calFormSection"">Event Information</div>"
    stream = stream & "<div class=""calFormCaption"">Name:</div>"
    stream = stream & "<div class=""calFormValue"">" & Main.GetFormInputText(rnCalName, calName, , 45, rnCalName) & "</div>"
    stream = stream & "<div class=""calFormCaption"">Start Date:</div>"
    stream = stream & "<div class=""calFormValue"">" & Main.GetFormInputDate(rnCalStart, calStart, 45, rnCalStart) & "</div>"
    stream = stream & "<div class=""calFormCaption"">End Date:</div>"
    stream = stream & "<div class=""calFormValue"">" & Main.GetFormInputDate(rnCalEnd, calEnd, 45, rnCalEnd) & "</div>"
    stream = stream & "<div class=""calFormCaption"">Event URL:</div>"
    stream = stream & "<div class=""calFormValue"">" & Main.GetFormInputText(rnCalURL, calURL, , 45, rnCalURL) & "</div>"
    stream = stream & "<div class=""calFormCaption"">Day View Copy:</div>"
    stream = stream & "<div class=""calFormEdit"">" & Main.GetFormInputHTML(rnCalDayCopy, calDayCopy) & "</div>"
    stream = stream & "<div class=""calFormCaption"">Details View Copy:</div>"
    stream = stream & "<div class=""calFormEdit"">" & Main.GetFormInputHTML(rnCalDetailCopy, calDetailCopy) & "</div>"
    stream = stream & "<div class=""calFormSection"">Select a Calendar(s)</div>"
    stream = stream & "<div class=""calFormList"">" & calSelect & "</div>"
    stream = stream & "<div class=""calFormSection"">Recurring Event</div>"
    stream = stream & "<div class=""calFormCaption"">Repeats:</div>"
    stream = stream & "<div class=""calFormValue"">" & Main.GetFormInputSelectList(rnCalRepeats, calRepeats, "Daily,Weekly,Monthly,Yearly", "Select One", rnCalRepeats) & "</div>"
    stream = stream & "<div class=""calFormCaption"">Start Repeat:</div>"
    stream = stream & "<div class=""calFormValue"">" & Main.GetFormInputDate(rnCalRptStart, calRptStart, 45, rnCalRptStart) & "</div>"
    stream = stream & "<div class=""calFormCaption"">End Repeat:</div>"
    stream = stream & "<div class=""calFormValue"">" & Main.GetFormInputDate(rnCalRptEnd, calRptEnd, 45, rnCalRptEnd) & "</div>"
    stream = stream & "<div class=""calFormValue""></div>"
    stream = stream & "<div id=""calFormButtons"">" & hdnList & btnList & "</div>"
    stream = stream & "</div>"
    stream = stream & Main.GetFormEnd
    
    getFormEdit = stream

    Exit Function
ErrorTrap:
    Call HandleError("getNavigation", "getFormEdit", Err.Number, Err.Source, Err.Description, True, False)
    End Function

Private Sub loadEditForm()
    On Error GoTo ErrorTrap
    
    Dim CS As Long
    '
    ' this should load with or without a button press
    ' load the default calendar event id, set from an "add link", or just 0
    '
    calDefaultEventCalendarID = Main.GetStreamInteger(rnDefaultEventCalendarID)
    
    btnValue = Main.GetStreamText("button")
    
    If btnValue <> "" Then
        eventID = Main.GetStreamInteger(rnEventID)
        
        calName = Main.GetStreamText(rnCalName)
        calStart = Main.GetStreamText(rnCalStart)
        calEnd = Main.GetStreamText(rnCalEnd)
        calURL = Main.GetStreamText(rnCalURL)
        calDayCopy = Main.GetStreamText(rnCalDayCopy)
        calDetailCopy = Main.GetStreamText(rnCalDetailCopy)
        calRepeats = Main.GetStreamInteger(rnCalRepeats)
        calRptStart = Main.GetStreamText(rnCalRptStart)
        calRptEnd = Main.GetStreamText(rnCalRptEnd)
        calPrntID = Main.GetStreamInteger(rnCalPrntID)
        
        calORName = Main.GetStreamText(rnCalORName)
        calORStart = Main.GetStreamText(rnCalORStart)
        calOREnd = Main.GetStreamText(rnCalOREnd)
        calORURL = Main.GetStreamText(rnCalORURL)
        calORDayCopy = Main.GetStreamText(rnCalORDayCopy)
        calORDetailCopy = Main.GetStreamText(rnCalORDetailCopy)
        calORRepeats = Main.GetStreamInteger(rnCalORRepeats)
        calORRptStart = Main.GetStreamText(rnCalORRptStart)
        calORRptEnd = Main.GetStreamText(rnCalORRptEnd)
    Else
        eventID = Main.GetStreamInteger(rnEventID)
        calStart = Main.GetStreamText(rnCalStart)
        
        CS = Main.OpenCSContent("Calendar Events", "ID=" & eventID)
        If Main.CSOK(CS) Then
            calName = Main.GetCSText(CS, "Name")
            calStart = Main.GetCSText(CS, "StartDate")
            calEnd = Main.GetCSText(CS, "EndDate")
            calURL = Main.GetCSText(CS, "EventLink")
            calDayCopy = Main.GetCSTextFile(CS, "BriefFileName")
            calDetailCopy = Main.GetCSTextFile(CS, "CopyFileName")
            calRepeats = Main.GetCSInteger(CS, "Repeats")
            calRptStart = Main.GetCSText(CS, "RepeatStart")
            calRptEnd = Main.GetCSText(CS, "RepeatStart")
            calPrntID = Main.GetCSInteger(CS, "ParentCalendarEventID")
        End If
        Call Main.CloseCS(CS)
        
        If calPrntID <> 0 Then
            CS = Main.OpenCSContent("Calendar Events", "ID=" & calPrntID)
            If Main.CSOK(CS) Then
                calRepeats = Main.GetCSInteger(CS, "Repeats")
                calRptStart = Main.GetCSText(CS, "RepeatStart")
                calRptEnd = Main.GetCSText(CS, "RepeatStart")
            End If
            Call Main.CloseCS(CS)
        End If
        
        calORName = calName
        calORStart = calStart
        calOREnd = calEnd
        calORURL = calURL
        calORDayCopy = calDayCopy
        calORDetailCopy = calDetailCopy
        calORRepeats = calRepeats
        calORRptStart = calRptStart
        calORRptEnd = calRptEnd
        
    End If
        
    Exit Sub
ErrorTrap:
    Call HandleError("getNavigation", "loadEditForm", Err.Number, Err.Source, Err.Description, True, False)
    End Sub

Private Sub processEditForm()
    On Error GoTo ErrorTrap
    
    Dim CS As Long
    
    If btnValue <> "" Then
        
        Select Case btnValue
            Case btnSave
                Call saveEvent
            Case btnSaveSeries
                Call saveEvent
            Case btnCancel
                
            Case btnDelete
                Call Main.DeleteContentRecords("Calendar Event Rules", "CalendarEventID=" & eventID)
            Case btnDeleteSeries
                Call Main.DeleteContentRecords("Calendar Events", "(ID=" & eventID & ") OR (ParentCalendarEventID=" & calPrntID & ")")
        End Select
        
    End If
        
    Exit Sub
ErrorTrap:
    Call HandleError("getNavigation", "processEditForm", Err.Number, Err.Source, Err.Description, True, False)
    End Sub

Private Sub saveEvent()
    On Error GoTo ErrorTrap
    
    Dim CS As Long
    Dim newEvent As Boolean
    Dim recordID As Long
    Dim srsUpd As Boolean
    Dim Response As String
    
    If eventID = 0 Then
        CS = Main.InsertCSContent("Calendar Events")
        newEvent = True
    Else
        If btnValue = btnSave Then
            CS = Main.OpenCSContent("Calendar Events", "ID=" & eventID)
        ElseIf btnValue = btnSaveSeries Then
            If calPrntID = 0 Then
                calPrntID = eventID
            End If
            CS = Main.OpenCSContent("Calendar Events", "ID=" & calPrntID)
            srsUpd = True
        End If
    End If
    If Main.CSOK(CS) Then
        If eventID = 0 Then
            Call Main.SetCS(CS, "calendarDate", calStart)
        End If
        Call Main.SetCS(CS, "Name", calName)
        If Not srsUpd Then
            Call Main.SetCS(CS, "StartDate", calStart)
            If Trim(calEnd) = "" Then
                Call Main.SetCS(CS, "EndDate", calStart)
            Else
                Call Main.SetCS(CS, "EndDate", calEnd)
            End If
        Else
            Call Main.SetCS(CS, "StartCheckOn", "")
        End If
        Call Main.SetCS(CS, "EventLink", calURL)
        Call Main.SetCS(CS, "BriefFileName", calDayCopy)
        Call Main.SetCS(CS, "CopyFileName", calDetailCopy)
        eventID = Main.GetCSInteger(CS, "ID")
        Call Main.SaveCSRecord(CS)
    End If
    Call Main.CloseCS(CS)
    
    Call Main.ProcessCheckList(rnCals, "Calendar Events", CStr(eventID), "Calendars", "Calendar Event Rules", "CalendarEventID", "CalendarID")
    '
    ' if there is a default calendar id set, make a Calendar Event Rule for this calendar
    '
Dim EventRuleExists As Boolean

    If calDefaultEventCalendarID <> 0 Then
        '
        ' if they checked the box manually, the rule will already be set. Check for it before creating a new one
        '
        CS = Main.OpenCSContent("Calendar event rules", "(calendarid=" & calDefaultEventCalendarID & ")and(CalendarEventID=" & eventID & ")", , , , , "id")
        EventRuleExists = Main.IsCSOK(CS)
        Call Main.CloseCS(CS)
        If Not EventRuleExists Then
            '
            ' new event and no checkbox is checked, force it
            '
            CS = Main.InsertCSContent("Calendar Event Rules")
            If Main.IsCSOK(CS) Then
                Call Main.SetCS(CS, "CalendarID", calDefaultEventCalendarID)
                Call Main.SetCS(CS, "CalendarEventID", eventID)
            End If
            Call Main.CloseCS(CS)
        End If
    End If
    '
    If (calPrntID <> 0) Or (newEvent) Then
        If newEvent Then
            recordID = eventID
        Else
            recordID = calPrntID
        End If
        CS = Main.OpenCSContent("Calendar Events", "ID=" & recordID)
        If Main.CSOK(CS) Then
            Call Main.SetCS(CS, "Repeats", calRepeats)
            Call Main.SetCS(CS, "RepeatStart", calRptStart)
            Call Main.SetCS(CS, "RepeatEnd", calRptEnd)
            Call Main.SaveCSRecord(CS)
        End If
        Call Main.CloseCS(CS)
    End If
    
    If srsUpd Or newEvent Then
        CS = Main.OpenCSContent("Add-Ons", "ccGUID='{108F995C-2123-4EAE-84DF-D42A931F073B}'", , , , , "ProcessRunOnce")
        If Main.CSOK(CS) Then
            'Response = Main.ExecuteAddon(Main.GetCSInteger(CS, "ID"), Main.GetCSText(CS, "Name"), "", ContextAdmin, "", 0, "", Main.GetCSText(CS, "ccGUID"), 0)
            Call Main.SetCS(CS, "ProcessRunOnce", 1)
            Call Main.SaveCSRecord(CS)
        End If
        Call Main.CloseCS(CS)
    End If
        
    Exit Sub
ErrorTrap:
    Call HandleError("getNavigation", "saveEvent", Err.Number, Err.Source, Err.Description, True, False)
    End Sub
'
'
'
Private Property Get adminUrl()
    If adminUrl_local = "" Then
        adminUrl_local = Main.GetSiteProperty("adminUrl")
    End If
    adminUrl = adminUrl_local
End Property
