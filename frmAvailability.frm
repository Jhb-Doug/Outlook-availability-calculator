VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAvailability 
   Caption         =   "Calculate Availability"
   ClientHeight    =   5700
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   5040
   OleObjectBlob   =   "frmAvailability.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmAvailability"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'To use this calculator, go to Outlook File-->options-->customize ribbon (Suggest doing this under the Calendar category)
'Add a new group to the list on the right side of the screen
'Select macros from the left dropdown combo box, select the OpenAvailabilityCalculator macro and add a new command to the group you created.
'Name it appropriately.

'Once the button is loaded, select the portion of the calendar you want to generate availability for and click the button you created
'This should load the default dates and times in the calculator for the period which you selected. Edit any parameters you want to and
'click the calculate button. Copy the text and paste it where you need it.



Private Sub ComboBox1_Change()
    On Error GoTo ender
    Dim strStart As String
    Dim strEnd As String
    If ComboBox2.Text <> "" And ComboBox1.Text <> "" Then
        strStart = Split(ComboBox1.Text, " ")(1) & " " & Split(ComboBox1.Text, " ")(2) & " " & Split(ComboBox1.Text, " ")(3)
        strEnd = Split(ComboBox2.Text, " ")(1) & " " & Split(ComboBox2.Text, " ")(2) & " " & Split(ComboBox2.Text, " ")(3)
        If CDate(strEnd) < CDate(strStart) Then ComboBox2.Text = ComboBox1.Text
    End If
ender:
    On Error GoTo 0
End Sub

Private Sub CommandButton1_Click()


    Dim myStart As Date
    Dim myEnd As Date
    Dim oCalendar As Outlook.Folder
    Dim oItems As Outlook.Items
    Dim oAppt As Outlook.AppointmentItem
    Dim varAppts(1000) As Variant
    Dim tAvl(1000, 2) As Variant
    Dim tMeetings(10000, 2) As Variant
    Dim i As Integer
    Dim strAvailability As String

    frmAvailability.MousePointer = fmMousePointerHourGlass

    myStart = CDate(Mid(ComboBox1.Text, 5) & " " & ComboBox3.Text)
    myEnd = CDate(Mid(ComboBox2.Text, 5) & " " & ComboBox4.Text)

    Debug.Print "Start:", myStart
    Debug.Print "End:", myEnd

    Me.MousePointer = fmMousePointerHourGlass
    
    Set oCalendar = Application.Session.GetDefaultFolder(olFolderCalendar)
    Set oItems = oCalendar.Items
    oItems.IncludeRecurrences = True
    oItems.Sort "[Start]"

    'Note you would need to change the order of month and days in the string format below if you have mm/dd on your system settings
    strRestriction = "[Start] >= '" & _
        Format$(DateAdd("d", -1, myStart), "dd/mm/yyyy hh:mm AMPM") _
        & "' AND [End] <= '" & _
        Format$(DateAdd("d", 1, myEnd), "dd/mm/yyyy hh:mm AMPM") & "'"

    Set oItems = oItems.Restrict(strRestriction)
    
    strAvailability = "I am available at the following times: " & vbCrLf
    t = myStart
    i = 0
    For Each oAppt In oItems
        Debug.Print oAppt.Start & "," & oAppt.End
        If (oAppt.Start <= myStart And oAppt.End >= myStart) Or _
            (oAppt.Start >= myStart And oAppt.End <= myEnd) Or _
            (oAppt.Start <= myEnd And oAppt.End >= myEnd) Then

            
            If t < oAppt.Start Then
                tAvl(i, 1) = t
                tAvl(i, 2) = oAppt.Start
                i = i + 1

            End If
            If t < oAppt.End Then t = oAppt.End
        End If
        If oAppt.Start > myEnd Then Exit For
    Next
    If CDate(t) < myEnd Then
        tAvl(i, 1) = t
        tAvl(i, 2) = myEnd
    End If

    For i = 0 To 999
        If tAvl(i, 1) <> "" And Not IsNull(tAvl(i, 1)) Then
            If CDate(Format(tAvl(i, 1), "hh:mm AM/PM")) < CDate(Format(ComboBox6.Text, "hh:mm AM/PM")) Then
                tAvl(i, 1) = Format(tAvl(i, 1), "dd mmm yyyy") & " " & ComboBox6.Text
            End If
            If CDate(Format(tAvl(i, 2), "hh:mm AM/PM")) < CDate(Format(ComboBox6.Text, "hh:mm AM/PM")) Then
                tAvl(i, 2) = Format(tAvl(i, 2), "dd mmm yyyy") & " " & ComboBox6.Text
            End If
            
            If CDate(Format(tAvl(i, 1), "hh:mm AM/PM")) > CDate(Format(ComboBox5.Text, "hh:mm AM/PM")) Then
                tAvl(i, 1) = Format(tAvl(i, 1), "dd mmm yyyy") & " " & ComboBox5.Text
            End If
            
            If CDate(Format(tAvl(i, 2), "hh:mm AM/PM")) > CDate(Format(ComboBox5.Text, "hh:mm AM/PM")) Then
                tAvl(i, 2) = Format(tAvl(i, 2), "dd mmm yyyy") & " " & ComboBox5.Text
            End If
        Else
            Exit For
        End If
    Next
    For i = 0 To 999
        If tAvl(i, 1) <> "" Then
            Debug.Print "***", Format(tAvl(i, 1), "dd mmm yyyy hh:mm AM/PM"), Format(tAvl(i, 2), "dd mmm yyyy hh:mm AM/PM")
            If Format(tAvl(i, 1), "dd mmm yyyy") = Format(tAvl(i, 2), "dd mmm yyyy") Then 'Same day
                strAvailability = AddAvailability(strAvailability, Format(tAvl(i, 1), "dd mmm yyyy hh:mm AM/PM"), Format(tAvl(i, 2), "dd mmm yyyy hh:mm AM/PM"))
            Else 'Moved to next day in this availability period
                strAvailability = AddAvailability(strAvailability, Format(tAvl(i, 1), "dd mmm yyyy hh:mm AM/PM"), Format(tAvl(i, 1), "dd mmm yyyy") & " " & ComboBox5.Text) 'Add from start of availability to end of day
                For j = 1 To 100
                    If Format(DateAdd("d", j, Format(tAvl(i, 1), "dd mmm yyyy")), "dd mmm yyyy") = Format(tAvl(i, 2), "dd mmm yyyy") Then
                        strAvailability = AddAvailability(strAvailability, Format(tAvl(i, 2), "dd mmm yyyy") & " " & ComboBox6.Text, Format(tAvl(i, 2), "dd mmm yyyy hh:mm AM/PM"))
                        Exit For
                    Else
                        strAvailability = AddAvailability(strAvailability, Format(DateAdd("d", j, Format(tAvl(i, 1), "dd mmm yyyy")), "dd mmm yyyy") & " " & ComboBox6.Text, Format(DateAdd("d", j, Format(tAvl(i, 1), "dd mmm yyyy")) & " " & ComboBox6.Text, "dd mmm yyyy hh:mm AM/PM"))
                    End If
                Next
            End If
        Else
            Exit For
        End If
    Next
    TextBox1.Text = strAvailability
    TextBox1.SelStart = 0
    TextBox1.SelLength = Len(TextBox1.Text)
    TextBox1.SetFocus
    CommandButton1.MousePointer = fmMousePointerDefault
    Me.MousePointer = fmMousePointerDefault
End Sub

Private Function AddAvailability(strMain As String, strDt1 As Variant, strDt2 As Variant)
    If Format(strDt1, "dd mmm yyyy hh:mm AM/PM") <> Format(strDt2, "dd mmm yyyy hh:mm AM/PM") Then
        strDate = Left(Format(strDt1, "dd mmm yyyy hh:mm AM/PM"), 11)
        If Not (CheckBox1.Value = True And (WeekdayName(Weekday(CDate(strDt1), vbMonday), True) = "Sun" Or WeekdayName(Weekday(CDate(strDt1), vbMonday), True) = "Sat")) Then
            If InStr(1, strMain, strDate) = 0 Then
                strMain = strMain & " - " & WeekdayName(Weekday(CDate(strDt1), vbMonday), True) & " " & Format(CDate(strDt1), "dd mmm yyyy") & ":" & vbCrLf
            End If
            Debug.Print Format(strDt1, "hh:mm AM/PM") & " - " & Format(strDt2, "hh:mm AM/PM") & vbCrLf
            strMain = strMain & "   - " & Format(strDt1, "hh:mm AM/PM") & " - " & Format(strDt2, "hh:mm AM/PM") & vbCrLf
        End If
    End If
    AddAvailability = strMain
    
    
End Function




Private Sub UserForm_Initialize()
    Dim i As Integer
    Dim oView As Outlook.View
    Dim oCalView As Outlook.CalendarView
    Dim oExpl As Outlook.Explorer
    Dim dtStart As Date
    Dim dtEnd As Date
    Dim blnUseSelection As Boolean
 
    Set oExpl = Application.ActiveExplorer
    Set oFolder = Application.ActiveExplorer.CurrentFolder
    Set oView = oExpl.CurrentView
    
    ' Check whether the active explorer is displaying a calendar view.
    If oView.ViewType = olCalendarView Then
        Set oCalView = oExpl.CurrentView
        dtStart = oCalView.SelectedStartTime
        dtEnd = oCalView.SelectedEndTime
        If dtStart < dtEnd Then blnUseSelection = True
    End If
 
    For i = 0 To 30
        ComboBox1.AddItem WeekdayName(Weekday(DateAdd("d", i, Now()), vbMonday), True) & " " & Format(DateAdd("d", i, Now()), "dd mmm yyyy")
        ComboBox2.AddItem WeekdayName(Weekday(DateAdd("d", i, Now()), vbMonday), True) & " " & Format(DateAdd("d", i, Now()), "dd mmm yyyy")
    Next
    
    If blnUseSelection Then
        ComboBox1.Text = WeekdayName(Weekday(DateAdd("d", 0, dtStart), vbMonday), True) & " " & Format(DateAdd("d", 0, dtStart), "dd mmm yyyy")
        ComboBox2.Text = WeekdayName(Weekday(DateAdd("d", 0, dtEnd), vbMonday), True) & " " & Format(DateAdd("d", 0, dtEnd), "dd mmm yyyy")
    Else
        ComboBox1.Text = WeekdayName(Weekday(DateAdd("d", 0, Now()), vbMonday), True) & " " & Format(DateAdd("d", 0, Now()), "dd mmm yyyy")
        ComboBox2.Text = WeekdayName(Weekday(DateAdd("d", 0, Now()), vbMonday), True) & " " & Format(DateAdd("d", 0, Now()), "dd mmm yyyy")
    End If
    
    t = CDate("00:00")
    For i = 0 To 23
        ComboBox3.AddItem t
        ComboBox4.AddItem t
        ComboBox5.AddItem t
        ComboBox6.AddItem t
        t = DateAdd("h", 1, t)
    Next
    
    If blnUseSelection Then
        ComboBox3.Text = Format(DateAdd("d", 0, dtStart), "HH:nn AM/PM")
        ComboBox4.Text = Format(DateAdd("d", 0, dtEnd), "HH:nn AM/PM")
    Else
        ComboBox3.Text = "10:00 AM"
        ComboBox4.Text = "5:00 PM"
    End If
    
    'After hours
    ComboBox6.Text = "8:00 AM"
    ComboBox5.Text = "5:00 PM"
    
    
End Sub
