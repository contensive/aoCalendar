VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "conditionerClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Main As Object
Private CSV As Object

Public Function execute(CsvObject As Object, MainObject As Object, OptionString As String, FilterInput As String) As String
    Set CSV = CsvObject
    
    Dim CS As Long
    Dim eventID As Long
    Dim calString As String
    Dim startDate As Date
    Dim endDate As Date
    Dim calEndDate As Date
    Dim curDate As Date
    Dim stream As String
    Dim CSI As Long
    Dim CSIR As Long
    Dim CSRules As Long
    Dim newEventID As Long
    Dim interval As Integer
    
    Dim calStrArr() As String
    Dim calStrPointer As Long
    Dim calStrCount As Long
    
    Dim eventStart As Date
    Dim eventEnd As Date
    Dim duration As Double
    
    Dim lastCheck As Date
    Dim calYrLimit As Double
    
    calYrLimit = Round(KmaEncodeNumber(CSV.GetSiteProperty("CALENDARYEARLIMIT", "1", 0)), 0)
    calEndDate = CDate(CDbl(Date) + CDbl((365 * calYrLimit)))
    curDate = Date
    
    CS = CSV.OpenCSContent("Calendar Events", "(Repeats<>0)")
    Do While CSV.IsCSOK(CS)
        eventID = CSV.GetCSInteger(CS, "ID")
        startDate = CSV.GetCSDate(CS, "RepeatStart")
        endDate = CSV.GetCSDate(CS, "RepeatEnd")
        lastCheck = CSV.GetCSDate(CS, "startCheckOn")
        
        eventStart = CSV.GetCSDate(CS, "StartDate")
        eventEnd = CSV.GetCSDate(CS, "EndDate")
        
        duration = CDbl(eventEnd) - CDbl(eventStart)
        
        Select Case CSV.GetCSInteger(CS, "Repeats")
            Case 1
                interval = 1
            Case 2
                interval = 7
            Case 3
                interval = 30
            Case 4
                interval = 365
            Case Else
                interval = 365
        End Select
        
        CSRules = CSV.OpenCSContent("Calendar Event Rules", "CalendareventID=" & eventID, , , , , , "CalendarID")
        Do While CSV.IsCSOK(CSRules)
            calString = calString & CSV.GetCSInteger(CSRules, "CalendarID")
            calString = calString & ","
            Call CSV.NextCSRecord(CSRules)
        Loop
        Call CSV.CloseCS(CSRules)
        
        If startDate = CDate(0) Then
            startDate = eventStart
        End If
        If (endDate > calEndDate) Or (endDate = CDate(0)) Then
            endDate = calEndDate
        End If
        
        If lastCheck <> CDate(0) Then
            startDate = lastCheck
        End If
        
        curDate = startDate
        
        Do While curDate <= endDate
            CSI = CSV.OpenCSContent("Calendar Events", "((ID=" & eventID & ") and (calendarDate=" & KmaEncodeSQLDate(curDate) & ")) OR ((ParentCalendarEventID=" & eventID & ") and (calendarDate=" & KmaEncodeSQLDate(curDate) & "))")
            If Not CSV.IsCSOK(CSI) Then
                CSI = CSV.InsertCSRecord("Calendar Events", 0)
            End If
            If CSV.IsCSOK(CSI) Then
            
                newEventID = CSV.GetCS(CSI, "ID")
                Call CSV.SetCS(CSI, "Name", CSV.GetCS(CS, "Name"))
                Call CSV.SetCS(CSI, "StartDate", curDate)
                Call CSV.SetCS(CSI, "EndDate", CDate(CDbl(curDate) + duration))
                Call CSV.SetCS(CSI, "StartTime", CSV.GetCS(CS, "StartTime"))
                Call CSV.SetCS(CSI, "EndTime", CSV.GetCS(CS, "EndTime"))
                Call CSV.SetCS(CSI, "ParentCalendarEventID", eventID)
                Call CSV.SetCS(CSI, "calendarDate", curDate)
                Call CSV.SetCS(CSI, "EventLink", CSV.GetCS(CS, "EventLink"))
                Call CSV.SetCS(CSI, "CopyFileName", CSV.GetCS(CS, "CopyFileName"))
                Call CSV.SetCS(CSI, "BriefFileName", CSV.GetCS(CS, "BriefFileName"))
                Call CSV.SaveCSRecord(CSI)
                
                Call CSV.DeleteContentRecords("Calendar Event Rules", "CalendarEventID=" & newEventID)
                
                If InStr(1, calString, ",", vbTextCompare) <> 0 Then
                    calStrArr = Split(calString, ",")
                    calStrCount = UBound(calStrArr)
                    For calStrPointer = 0 To calStrCount
                        If calStrArr(calStrPointer) <> "" Then
                            CSIR = CSV.InsertCSRecord("Calendar Event Rules", 0)
                            If CSV.IsCSOK(CSIR) Then
                                Call CSV.SetCS(CSIR, "CalendarID", calStrArr(calStrPointer))
                                Call CSV.SetCS(CSIR, "CalendarEventID", newEventID)
                            End If
                            Call CSV.CloseCS(CSIR)
                        End If
                    Next
                End If
            End If
            Call CSV.CloseCS(CSI)
            
            curDate = CDate(CDbl(curDate) + interval)
        Loop
        
        calString = ""
        Call CSV.SetCS(CS, "startCheckOn", curDate)
        Call CSV.NextCSRecord(CS)
    Loop
    Call CSV.CloseCS(CS)
    
    Exit Function
ErrorTrap:
    Call HandleError("conditionerClass", "execute", Err.Number, Err.Source, Err.Description, True, False)
    End Function
