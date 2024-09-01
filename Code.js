/*********************************************************************
 *          stu-srvcs-pac-automated-cal-appointment-project          *
 *                                                                   *
 * Author: Alvaro Gomez                                              *
 *         Academic Technology Coach                                 *
 *         Northside Independent School District                     *
 *         alvaro.gomez@nisd.net                                     *
 *         Office: +1-210-397-9408                                   *
 *         Mobile: +1-210-363-1577                                   *
 *                                                                   *
 * Purpose: This script will run when a submission is made on the    *
 *          associated Google Form. The script will invite a         *
 *          registrant by adding the submitted email address to a    *
 *          calendar appointment.                                    *
 *                                                                   *
 * Usage: @todo Create a Google Form with the necessary fields (i.e  *
 *        Name, email address). @todo Create the appointment for the *
 *        event on the organizer's Google Calendar. @todo Get the    *
 *        get the calendar id and the appointment id.                *
 *        listEventsForSpecificDate() is a helper function that      *
 *        can be used to identify the event's id. @todo Set a        *
 *        trigger to run addParticipantsToEvent() when a form is     *
 *        submitted.                                                 *
 *                                                                   *
 ********************************************************************/

function addParticipantsToEvent() {
  /** 
   * @todo Add the organizer's Google Calendar ID to const calendarId
  */
  const calendarId = 'alvaro.gomez@nisd.net';

  /** 
   * @todo Define the event's IDs. For this example, we have two events.
   * If necessary, use the listEventsForSpecificDate() function to get the event IDs.
  */ 
  const eventId1 = '2dacit6908ttabd9bmmgb5gnpp@google.com'; // First Event ID
  const eventId2 = '761c3vvedhhquconp9sb8r5jdu@google.com'; // Second Event ID

  // Access the calendar
  const calendar = CalendarApp.getCalendarById(calendarId);
  
  if (!calendar) {
    Logger.log('Calendar not found. Check the Calendar ID.');
    return;
  }

  // Open the active sheet
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Get the data range that includes email addresses in column B, event selections in column C, and notes in column D
  const dataRange = sheet.getRange(2, 2, sheet.getLastRow() - 1, 3); // Email in column B, Selection in column C, Notes in column D
  
  // Get all data as a 2D array
  const data = dataRange.getValues();
  
  // Get the current date and time
  const now = new Date();
  const timestamp = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  
  // Add each email address to the selected event and update the sheet
  data.forEach((row, index) => {
    const email = row[0];
    const eventSelection = row[1]; // Column C for event selection
    const note = row[2]; // Column D for note
    let eventId;

    // This is a check to keep from adding a person to the event everytime the function is run. It will skip if note is not empty
    if (note) {
      Logger.log(`Skipping ${email} - Already added`);
      return;
    }

    // Determine which event to add the participant to based on selection
    if (eventSelection === 'EAST - Tuesday, September 24th, 6:00-7:30 PM, Rawlinson MS, 14100 Vance Jackson, San Antonio, TX 78249') {
      eventId = eventId1;
    } else if (eventSelection === 'WEST - Tuesday, October 8th, 6:00-7:30 PM, Vale MS, 2120 Ellison Drive, San Antonio, TX 78251') {
      eventId = eventId2;
    } else {
      Logger.log(`Invalid event selection: ${eventSelection}`);
      // sheet.getRange(index + 2, 4).setValue(`Invalid selection on ${timestamp}`);
      return;
    }

    // Retrieve the event by ID
    const event = calendar.getEventById(eventId);
    
    if (!event) {
      Logger.log(`Event not found. Please check the Event ID and ensure it matches exactly.`);
      //sheet.getRange(index + 2, 4).setValue(`Event not found on ${timestamp}`);
      return;
    }

    Logger.log('Event found: ' + event.getTitle());

    if (email) {
      try {
        event.addGuest(email.trim());
        Logger.log(`Added: ${email}`);
        sheet.getRange(index + 2, 4).setValue(`Added to ${event.getTitle()} on ${timestamp}`); // This updates the note column
      } catch (error) {
        Logger.log(`Failed to add ${email}: ${error.message}`);
        sheet.getRange(index + 2, 4).setValue(`Failed to add on ${timestamp}: ${error.message}`);
      }
    }
  });
}

/************************************************************************ 
 * listEventsForSpecificDate() is a helper function that can be used    *
 * to identify the event's ID.                                          *
 *                                                                      *
 * listEventsForSpecificDate will list (in the Log) all events for the  *
 * specified date in const specificDate.                                *
************************************************************************/
function listEventsForSpecificDate() {
  const calendarId = 'alvaro.gomez@nisd.net'; // Your Calendar ID
  const calendar = CalendarApp.getCalendarById(calendarId);
  
  if (!calendar) {
    Logger.log('Calendar not found. Check the Calendar ID.');
    return;
  }

  /** 
   * @todo Define the specific date to retrieve events for.
  */
  const specificDate = new Date(2024, 7, 31); // Months are 0-indexed: 7 = August
  
  // Retrieve events for the specific date
  const events = calendar.getEventsForDay(specificDate);
  
  // Log all events with their IDs
  events.forEach(ev => {
    Logger.log(`Event Found: ${ev.getTitle()} - ID: ${ev.getId()}`);
  });
}
