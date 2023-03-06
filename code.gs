function onOpen(){
  SpreadsheetApp.getUi().createMenu("Events").addItem("Create Event","recurringEvent").addToUi()
}

function recurringEvent()
{
  var ws = SpreadsheetApp.getActiveSpreadsheet()
  var ss = ws.getSheetByName('Events')
  var row = 2
  var lastRow = ss.getLastRow();
  var range = ss.getRange(2, 1, lastRow, 1);
  var numRows = range.getValues().filter(String).length;


  for(var index = 0; index < numRows ; index ++)
  {
    
     var title = ss.getRange(row, 1).getValue()
      var startDate = ss.getRange(row, 2).getValue()
      var duration = ss.getRange(row, 7).getValue()
      var repetition = ss.getRange(row, 6).getValue()
      var description = ss.getRange(row, 8).getValue()
      var location = ss.getRange(row, 9).getValue()
      
      var endDate = new Date(startDate)
      endDate.setTime(endDate.getTime() + 1000 * 60 * duration)

      var recurrence = ""
      if(repetition == 'None')
        recurrence = CalendarApp.newRecurrence().addDate(new Date(startDate))
      else if(repetition == "Daily")
        recurrence = CalendarApp.newRecurrence().addDailyRule()
      else if(repetition == "Weekly")
        recurrence = CalendarApp.newRecurrence().addWeeklyRule()
      else if(repetition == "Monthly")
        recurrence = CalendarApp.newRecurrence().addMonthlyRule()

      var events = CalendarApp.createEventSeries(title, startDate, endDate, recurrence)
      events.setDescription(description)
      events.setVisibility(CalendarApp.Visibility.PUBLIC)
      events.setLocation(location)
      events.removeAllReminders();
  
      // for(var jdex = 0; jdex < events.length; jdex ++)
      // {
      //   var event = events[jdex];
      //   event.setVisibility(CalendarApp.Visibility.PUBLIC)
      // }
      row ++
  }
  
 
}