Public Sub DoubleMeeting()

    Dim myMeeting, myPTO AS Object

    Set myMeeting = Application.CreateItem(olAppointmentItem)
    myMeeting.MeetingStatus = olMeeting
    myMeeting.AllDayEvent = True
    myMeeting.BusyStatus = olFree ' Must be olFree so other persons' calendars don't show as olBusy
    myMeeting.ReminderSet = False

    Set myPTO = Application.CreateItem(olAppointmentItem)
    ' Not a metting, this is for own calendar only, don't need MeetingStatus set
    myPTO.AllDayEvent = True
    myPTO.BusyStatus = olOutOfOffice
    myPTO.ReminderSet = False

    myMeeting.Display
    myPTO.Display

End Sub
