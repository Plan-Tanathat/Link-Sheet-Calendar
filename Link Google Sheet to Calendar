function createTimeDrivenTrigger() {
 var triggers = ScriptApp.getProjectTriggers();
 for (var i = 0; i < triggers.length; i++) {
   if (triggers[i].getHandlerFunction() == 'createCalendarEvents') {
     ScriptApp.deleteTrigger(triggers[i]);
   }
 }
 //create a new trigger to run the function every 10 minutes.
 ScriptApp.newTrigger('createCalendarEvents')
   .timeBased()
   .everyMinutes(10)
   .create();
}

function createCalendarEvents() {
 var sheet = SpreadsheetApp.getActiveSheet();
 var data = sheet.getDataRange().getValues();
 var calendar = CalendarApp.getCalendarById('469752d4aecb92a4b0628751ef6cc11925b102156098843f598d8c096e1361b0@group.calendar.google.com');

 for (var i = 1; i < data.length; i++) {
   var row = data[i];
   var primaryKey = row[1];           //Primary key
   var startDate = new Date(row[11]);
   var endDate = new Date(row[12]);

   // if startDate and endDate do not have a time, set the time to 8:00 for startDate and 16:00 for endDate.
   if (!startDate || isNaN(startDate.getTime())) {
     continue;
   }
   startDate.setHours(8, 0, 0);

   if (!endDate || isNaN(endDate.getTime())) {
     endDate = new Date(startDate);
   }
   endDate.setHours(16, 0, 0);

   var title = row[8];
   var description =           // Set Column
       "ชื่อ " + row[4] + " " + row[5] + " " + row[7] + "\n" + "\n" +
       "1. " + row[15] + "\n" +
       "2. " + row[16] + "\n" +
       "3. " + row[17] + "\n" +
       "4. " + row[18] + "\n" +
       "5. " + row[19] + "\n" +
       "อื่นๆ " + row[20];

   // Check if there are already events with the same primary key.
   var existingEvents = calendar.getEvents(startDate, endDate);
   var eventExists = existingEvents.some(function(event) {
     return event.getDescription().includes(primaryKey);
   });

   // If there is no event created from the same primary key, create a new event.
   if (!eventExists) {
     calendar.createEvent(title, startDate, endDate, {
       description: description + "\n\nPrimary Key: " + primaryKey
     });
   }
 }
}

// Delete all events from the calendar
function deleteAllCalendarEvents() {
 var calendar = CalendarApp.getCalendarById('469752d4aecb92a4b0628751ef6cc11925b102156098843f598d8c096e1361b0@group.calendar.google.com');
 var now = new Date();
 var events = calendar.getEvents(new Date(now.getFullYear(), 0, 1), new Date(now.getFullYear(), 11, 31)); // ดึงกิจกรรมทั้งปี
 for (var i = 0; i < events.length; i++) {
   events[i].deleteEvent();
 }
}

// Create a menu for automatic syncing with Calendar and deleting events
function onOpen() {
 var ui = SpreadsheetApp.getUi();
 ui.createMenu('Sync to Calendar')
   .addItem('Add events to calendar now', 'createCalendarEvents')
   .addItem('Set 10-min Auto Sync', 'createTimeDrivenTrigger')
   .addItem('Delete all events from calendar', 'deleteAllCalendarEvents')
   .addToUi();
}
