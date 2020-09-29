function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Sync to Calendar')
      .addItem('Schedule Meetings', 'scheduleEvents')
      .addSubMenu(ui.createMenu('Delete')
          .addItem('Delete events from Calendar', 'clearCalendar'))
      .addToUi();
}


function scheduleEvents() {
  // Creates Calendar events from google sheet based
  // on parameter info provided. 
  
  var spreadsheet = SpreadsheetApp.getActiveSheet();
  var calendarId = spreadsheet.getRange("B1").getValue();
  
  var eventCal = CalendarApp.getCalendarById(calendarId);

  // Hard coded data range
  var signups = spreadsheet.getRange("A5:F20").getValues(); 

  // Loops through data range
  for (x=0; x<signups.length; x++) {
    var shift = signups[x];
    if ((shift[0] == '') || (shift[1] == '') || (shift[2] == '')) {
      Logger.log("Warning: Empty value for cell. Skipping line");
      Logger.log('"Start: ' + shift[0] + ", End: " + shift[1] + ", Title: " + shift[2] + '"');
      continue;
    } else {
      var description_tm = shift[3] + "\n" + "\n" + "\n" + shift[2] +" was created by Meeting-Scheduler" + "\n" + "Contact meeting-scheduler@exampledomain.com for support";
      var startTime = shift[0];
      var endTime = shift[1];
      var eventTitle = shift[2];
      var options = {
        description: description_tm,
        guests: shift[4],
        sendInvites: shift[5]
      };
      Logger.log("Creating event with the following details: " + eventTitle + ", " + startTime + ", " + endTime + ", " + shift[3]);
      eventCal.createEvent(eventTitle, startTime, endTime, options);
    }   
  }
}

function clearCalendar() {
  // This one removes all of the shifts from the event
  // Calendar, so I put it in a sub-menu to make it
  // difficult to click by accident!
  
  var spreadsheet = SpreadsheetApp.getActiveSheet();
  var calendarId = spreadsheet.getRange("B1").getValue();
  
  var eventCal = CalendarApp.getCalendarById(calendarId);
  
  // Hard coded data range
  var signups = spreadsheet.getRange("A5:E20").getValues(); 
  
  // Loops through data range
  for (x=0; x<signups.length; x++) {
    var shift = signups[x];
    if ((shift[0] == '') || (shift[1] == '') || (shift[2] == '')) {
      Logger.log("Warning: Empty value for cell. Skipping line");
      Logger.log('"Start: ' + shift[0] + ", End: " + shift[1] + ", Title: " + shift[2] + '"');
      continue;
    } else {
      var description_tag = shift[2] + " was created by Meeting-Scheduler";
      Logger.log("description tag is: " + description_tag)
      var startTime = shift[0];
      var endTime = shift[1];
      var options = {
        search: description_tag,
        guests: shift[4],
        max: 1
      };
      Logger.log("Deleting event with the following details: ")
      Logger.log("Start: " + startTime + ", End: " + endTime + ", Title: " + shift[2]);
      var events = eventCal.getEvents(startTime, endTime, options);
      for(var i=0; i<events.length;i++){
        var ev = events[i];
        Logger.log(ev.getTitle()); // show event name in log
        ev.deleteEvent();
      }   
    }
  }
}
