Const HKEY_LOCAL_MACHINE = &H80000002

strComputer = "."
 
Set objRegistry = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
 
strKeyPath = "Location\where you keep scripts"

strValueName = "Call it what you like"

objRegistry.GetStringValue HKEY_LOCAL_MACHINE, strKeyPath, strValueName, strValue

If strValue = "Yes" Then

    Wscript.Quit

Else

    strValue = "Yes"

    objRegistry.SetStringValue HKEY_LOCAL_MACHINE, strKeyPath, strValueName, strValue

  Const olFolderCalendar = 9
          Const olAppointmentItem = 1
          Const olFree = 0

        Set objOutlook = CreateObject("Outlook.Application")
        Set objNamespace = objOutlook.GetNamespace("MAPI")
        Set objCalendar = objNamespace.GetDefaultFolder(olFolderCalendar)

        Set objDictionary = CreateObject("Scripting.Dictionary")

        '2020 Payrol period end dates

        objDictionary.Add "January 8, 2020", "Time to approve your employees ezlabor at the end of their shifts"
        objDictionary.Add "January 23, 2020", "Time to approve your employees ezlabor at the end of their shifts"
        objDictionary.Add "February 6, 2020", "Time to approve your employees ezlabor at the end of their shifts"
        objDictionary.Add "February 20, 2020", "Time to approve your employees ezlabor at the end of their shifts"
        objDictionary.Add "March 6, 2020", "Time to approve your employees ezlabor at the end of their shifts"
        objDictionary.Add "March 22, 2020", "Time to approve your employees ezlabor at the end of their shifts"
        objDictionary.Add "April 7, 2020", "Time to approve your employees ezlabor at the end of their shifts"
        objDictionary.Add "April 23, 2020", "Time to approve your employees ezlabor at the end of their shifts"
        objDictionary.Add "May 8, 2020", "Time to approve your employees ezlabor at the end of their shifts"
        objDictionary.Add "May 22, 2020", "Time to approve your employees ezlabor at the end of their shifts"
        objDictionary.Add "June 7, 2020", "Time to approve your employees ezlabor at the end of their shifts"
        objDictionary.Add "June 23, 2020", "Time to approve your employees ezlabor at the end of their shifts"
        objDictionary.Add "July 8, 2020", "Time to approve your employees ezlabor at the end of their shifts"
        objDictionary.Add "July 23, 2020", "Time to approve your employees ezlabor at the end of their shifts"
        objDictionary.Add "August 7, 2020", "Time to approve your employees ezlabor at the end of their shifts"
        objDictionary.Add "August 22, 2020", "Time to approve your employees ezlabor at the end of their shifts"
        objDictionary.Add "September 8, 2020", "Time to approve your employees ezlabor at the end of their shifts"
        objDictionary.Add "September 22, 2020", "Time to approve your employees ezlabor at the end of their shifts"
        objDictionary.Add "October 7, 2020", "Time to approve your employees ezlabor at the end of their shifts"
        objDictionary.Add "October 22, 2020", "Time to approve your employees ezlabor at the end of their shifts"
        objDictionary.Add "November 6, 2020", "Time to approve your employees ezlabor at the end of their shifts"
        objDictionary.Add "November 21, 2020", "Time to approve your employees ezlabor at the end of their shifts"
        objDictionary.Add "December 5, 2020", "Time to approve your employees ezlabor at the end of their shifts"
        objDictionary.Add "December 23, 2020", "Time to approve your employees ezlabor at the end of their shifts"
       
  colKeys = objDictionary.Keys

        For Each strKey in colKeys
  dtmHolidayDate = strKey
  strHolidayName = objDictionary.Item(strKey)

  Set objHoliday = objOutlook.CreateItem(olAppointmentItem)
 
  objHoliday.Subject = strHolidayName
  objHoliday.Start = dtmHolidayDate & " 7:00 AM"
  objHoliday.End = dtmHolidayDate & " 5:00 PM"
  objHoliday.AllDayEvent = False
  objHoliday.ReminderSet = True
  objHoliday.ReminderMinutesBeforeStart = 1440
  objHoliday.BusyStatus = olFree
  objHoliday.Save
            objHoliday.ReminderSet = True
            objHoliday.Save

        Next

  MsgBox "The ADP reminders have been added to your calendar, please do not click the link again or the remiders will be added a second time. Have a great day."  

    Wscript.Quit

End If
