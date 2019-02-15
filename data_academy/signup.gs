
/*
  Inserts  menu on open
 */
function onOpen() {
  var menu = [{name: 'Set up Courses', functionName: 'setUpConference_'}];
  SpreadsheetApp.getActive().addMenu('Courses Setup', menu);
}

/**
 * A set-up function that uses the data in the spreadsheet to create
 * Google Calendar events, a Google Form, and a trigger that allows the script
 * to react to form responses.
 */
function setUpConference_() {
  if (ScriptProperties.getProperty('calId')) {
    Browser.msgBox('Your courses are set up. Look in your calendar!');
  }
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName('Data Academy Courses');
  var range = sheet.getDataRange();
  var values = range.getValues();
  setUpCalendar_(values, range);
  setUpForm_(ss, values);
  ScriptApp.newTrigger('onFormSubmit').forSpreadsheet(ss).onFormSubmit()
      .create();
  ss.removeMenu('Courses Setup');
}

/**
 * Creates a Google Calendar with events for each workshop or training session in the
 * spreadsheet, then writes the event IDs to the spreadsheet for future use.
 *
 * @param {String[][]} values Cell values for the spreadsheet range.
 * @param {Range} range A spreadsheet range that contains conference data.
 */
function setUpCalendar_(values, range) {
  var cal = CalendarApp.createCalendar('Data Academy Calendar');
  for (var i = 1; i < values.length; i++) {
    var session = values[i];
    var title = session[0];
    var start = joinDateAndTime_(session[1], session[2]);
    var end = joinDateAndTime_(session[1], session[3]);
    var options = {location: session[4], description: session[5], sendInvites: true};  

    var event = cal.createEvent(title, start, end, options)
        .setGuestsCanSeeGuests(true);
    session[6] = event.getId();
  }
  range.setValues(values);

  // Store the ID for the Calendar, which is needed to retrieve events by ID.
  ScriptProperties.setProperty('calId', cal.getId());
}

/**
 * Creates a single Date object from separate date and time cells.
 *
 * @param {Date} date A Date object from which to extract the date.
 * @param {Date} time A Date object from which to extract the time.
 * @return {Date} A Date object representing the combined date and time.
 */
function joinDateAndTime_(date, time) {
  date = new Date(date);
  date.setHours(time.getHours());
  date.setMinutes(time.getMinutes());
  return date;
}

/**
 * Creates a Google Form that allows respondents to select which training
 * sessions they would like to attend, grouped by date and start time.
 *
 * @param {Spreadsheet} ss The spreadsheet that contains the conference data.
 * @param {String[][]} values Cell values for the spreadsheet range.
 */
function setUpForm_(ss, values) {
  // Group the sessions by title so that they can be passed to the form. 
  var schedule = {};
  for (var i = 1; i < values.length; i++) {
    var session = values[i];
    var title = session[0];
    var time = joinDateAndTime_(session[1], session[2]) ;
    if (!schedule[title]) {
      schedule[title] = [];
    }
    schedule[title].push(time);   
  }

  // Create the form and add a multiple-choice question for each timeslot.
  var form = FormApp.create('Data Academy Sign-up Form');
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
  form.addTextItem().setTitle('Name').setRequired(true);
  form.addTextItem().setTitle('Email').setRequired(true);
  var deptItem = form.addListItem();
      deptItem.setTitle('Department').setRequired(true);
      deptItem.setChoices([
          deptItem.createChoice('Analytics'),
          deptItem.createChoice('Communications'),
          deptItem.createChoice('Customer Support'),
          deptItem.createChoice('Didactics'),
          deptItem.createChoice('Engineering'),
          deptItem.createChoice('Finance & Legal'),
          deptItem.createChoice('Marketing'),
          deptItem.createChoice('Product')
      ]);
  
  
  for (var title in schedule) {
    var item = form.addMultipleChoiceItem();
    item.setTitle(title);
    item.setChoiceValues(schedule[title]);
 
  }
  
}

/**
 * A trigger-driven function to send out calendar invitations and a
 * personalized Google Docs itinerary after a user responds to the form.
 *
 * @param {Object} e The event parameter for form submission to a spreadsheet;
 *     see https://developers.google.com/apps-script/understanding_events
 */
function onFormSubmit(e) {
  var user = {name: e.namedValues['Name'][0], email: e.namedValues['Email'][0]};

  // Grab the session data again to match it to the user's choices.
  var response = [];
  var values = SpreadsheetApp.getActive().getSheetByName('Data Academy Courses')
     .getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    var session = values[i];
    var course_title = session[0];
    var time = joinDateAndTime_(session[1],session[2]);

    // For every selection in the response, find the matching title 
    // in the spreadsheet and add the session data to the response array.
    if (e.namedValues[course_title] == time.toString()) {
      response.push(session);
    }
  }
  sendInvites_(user, response);
  sendDoc_(user, response);
}

/**
 * Add the user as a guest for every session he or she selected.
 *
 * @param {Object} user An object that contains the user's name and email.
 * @param {String[][]} response An array of data for the user's session choices.
 */
function sendInvites_(user, response) {
  var id = ScriptProperties.getProperty('calId');
  var cal = CalendarApp.getCalendarById(id);
  for (var i = 0; i < response.length; i++) {
    cal.getEventSeriesById(response[i][6]).addGuest(user.email);
  }
}

/**
 * Create and share a personalized Google Doc with each user's signups.
 *
 * @param {Object} user An object that contains the user's name and email.
 * @param {String[][]} response An array of data for the user's session choices.
 */
function sendDoc_(user, response) {
  var doc = DocumentApp.create('Courses  for ' + user.name)
      .addEditor(user.email);
  var body = doc.getBody();
  var table = [['Course', 'Date', 'Location','Description']];
  for (var i = 0; i < response.length; i++) {
    table.push([response[i][0], response[i][1].toLocaleDateString(),
       response[i][4], response[i][5]]);
  }
  body.insertParagraph(0, doc.getName())
      .setHeading(DocumentApp.ParagraphHeading.HEADING1);
  table = body.appendTable(table);
  table.getRow(0).editAsText().setBold(true);
  doc.saveAndClose();

  // Email a link to the Doc as well as a PDF copy.
  MailApp.sendEmail({
    to: user.email,
    subject: doc.getName(),
    body: 'Thanks for registering for a course with the Data Academy! Here is an overview for the courses you signed up for: ' + doc.getUrl(),
    attachments: doc.getAs(MimeType.PDF),
  });
}
