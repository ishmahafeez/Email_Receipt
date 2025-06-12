function handleCheckboxEdit(e) {
  if (!e) return;

  const sheet = e.source.getActiveSheet();
  const editedCol = e.range.getColumn();
  const row = e.range.getRow();
  const newValue = e.value;

  Logger.log(`Edited cell: Row ${row}, Column ${editedCol}`);

  // Column 20: Set Meeting
  if (editedCol === 20 && String(newValue).toLowerCase() === 'true') {
    Logger.log("✅ Set Meeting checkbox checked — creating calendar event.");

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const rowData = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
    Logger.log("Headers: " + JSON.stringify(headers));
    Logger.log("RowData: " + JSON.stringify(rowData));
    const getValue = (header) => rowData[headers.indexOf(header)];

    const meetingDate = getValue("Meeting Date");
    const meetingTime = getValue("Meeting Time");
    const email = getValue("Email Address");
    const mapTemplate = getValue("Map Template");

    Logger.log("Meeting Date from sheet: " + meetingDate);
    Logger.log("Meeting Time from sheet: " + meetingTime);
    Logger.log("Email Address: " + email);

    if (meetingDate && meetingTime) {
      try {
        Logger.log("Attempting to get calendar...");
        const calendar = CalendarApp.getDefaultCalendar();
        Logger.log("Calendar obtained successfully");
        
        // Parse the meeting date and time
        const eventDate = new Date(meetingDate);
        Logger.log("Initial event date: " + eventDate.toString());
        
        // Parse time string (e.g., "5:00 PM")
        const timeMatch = meetingTime.match(/(\d+):(\d+)\s*(AM|PM)/i);
        if (!timeMatch) {
          throw new Error("Invalid time format: " + meetingTime);
        }
        
        const [_, hours, minutes, period] = timeMatch;
        let hour = parseInt(hours);
        
        Logger.log(`Raw time values - Hours: ${hours}, Minutes: ${minutes}, Period: ${period}`);
        
        // Convert to 24-hour format
        if (period.toUpperCase() === 'PM' && hour !== 12) {
          hour += 12;
        } else if (period.toUpperCase() === 'AM' && hour === 12) {
          hour = 0;
        }
        
        Logger.log(`Converted hour to 24-hour format: ${hour}`);
        
        // Create a new date object with the correct time
        const finalDate = new Date(eventDate);
        finalDate.setHours(hour, parseInt(minutes), 0, 0);
        Logger.log("Final event date with time: " + finalDate.toString());
        
        // Create end time (1 hour later)
        const endDate = new Date(finalDate);
        endDate.setHours(endDate.getHours() + 1);
        Logger.log("End date: " + endDate.toString());

        Logger.log("Creating calendar event with the following details:");
        Logger.log("- Title: IshMaps - Map Of US");
        Logger.log("- Start: " + finalDate.toString());
        Logger.log("- End: " + endDate.toString());
        Logger.log("- Guest: " + email);
        Logger.log("- Send Invites: true");

        // Create the event with guests
        const event = calendar.createEvent(
          "IshMaps - Map Of US",
          finalDate,
          endDate,
          {
            description: "Custom Maps",
            guests: email,
            sendInvites: true
          }
        );
        
        Logger.log("📅 Calendar event created successfully!");
        Logger.log("Event ID: " + event.getId());
        Logger.log("Event Title: " + event.getTitle());
        Logger.log("Start Time: " + event.getStartTime());
        Logger.log("End Time: " + event.getEndTime());
        Logger.log("Guest Email: " + email);

        // Verify guest was added
        const guests = event.getGuestList();
        Logger.log("Number of guests: " + guests.length);
        guests.forEach(guest => {
          Logger.log("Guest: " + guest.getEmail() + " (Status: " + guest.getGuestStatus() + ")");
        });

        // Try to manually send invites
        try {
          Logger.log("Attempting to manually send invites...");
          event.setGuestsCanModify(true);
          event.setGuestsCanSeeGuests(true);
          Logger.log("Guest permissions updated");
        } catch (inviteError) {
          Logger.log("❌ Error updating guest permissions: " + inviteError.toString());
        }

      } catch (error) {
        Logger.log("❌ Error creating calendar event: " + error.toString());
        Logger.log("Error details: " + JSON.stringify(error));
      }
    } else {
      Logger.log("❌ Missing meeting date or time");
    }
  }

  // Column 21: Check Out (send email)
  if (editedCol === 21 && String(newValue).toLowerCase() === 'true') {
    Logger.log("✅ Checkbox was checked — proceeding to send email.");

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const rowData = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
    Logger.log("Headers: " + JSON.stringify(headers));
    Logger.log("RowData: " + JSON.stringify(rowData));
    Logger.log("Index of 'Email Address': " + headers.indexOf("Email Address"));
    Logger.log("Index of 'Timestamp': " + headers.indexOf("Timestamp"));
    const getValue = (header) => rowData[headers.indexOf(header)];
    Logger.log("✅Getting values.");

    // Get all values from the row
    const timestamp = getValue("Timestamp");
    const email = getValue("Email Address");
    const mapTemplate = getValue("Map Template");
    const character1 = getValue("Character 1");
    const character2 = getValue("Character 2");
    const frame = getValue("Frame");
    let mapLink = getValue("Map Link");
    const price = getValue("Price");
    const framePrice = getValue("Frame Price");
    const subtotal = getValue("SubTotal");
    const meetingTime = getValue("Meeting Time");
    const setMeeting = getValue("Set Meeting");
        Logger.log("✅got valuess.");


    let html = receipt();

    const formattedDate = Utilities.formatDate(new Date(timestamp), Session.getScriptTimeZone(), 'M/d/yy');
    html = html.replace('{{date}}', formattedDate);
    html = html.replace('{{mapTemplate}}', mapTemplate);
    html = html.replace('{{character1}}', character1 || '');
    html = html.replace('{{character2}}', character2 || '');
    html = html.replace('{{price}}', `$${price}`);
    html = html.replace('{{subtotal}}', `$${subtotal}`);
    html = html.replace('{{downloadLink}}', mapLink);
        Logger.log("✅converted values.");


    if (String(frame).toLowerCase() === 'yes') {
      html = html.replace('{{frameSection}}', `
        <div class="item">
          <span>Frame</span>
          <span style="text-align: right;"> : $${framePrice}</span>
        </div>`);
    } else {
      html = html.replace('{{frameSection}}', '');
    }

    MailApp.sendEmail({
      to: email,
      subject: `Map of Us Receipt`,
      htmlBody: html
    });

    Logger.log("📩 Email sent to: " + email);
  }
}

function receipt() {
  return HtmlService.createHtmlOutputFromFile("receipt.html").getContent();
}


function onEdit(e) {
  handleCheckboxEdit(e);
}

function testSimpleEmail() {
  MailApp.sendEmail("your@email.com", "Test", "This is a test.");
}

function testCalendarEvent() {
  try {
    const calendar = CalendarApp.getDefaultCalendar();
    const startDate = new Date(); // Current time
    const endDate = new Date();
    endDate.setHours(endDate.getHours() + 1); // 1 hour later

    const event = calendar.createEvent(
      "Map Of US - Test Event",
      startDate,
      endDate,
      {
        description: "custom maps - Test Event"
      }
    );
    
    Logger.log("📅 Test Calendar event created successfully!");
    Logger.log("Event ID: " + event.getId());
    Logger.log("Event Title: " + event.getTitle());
    Logger.log("Start Time: " + event.getStartTime());
    Logger.log("End Time: " + event.getEndTime());
    
    return "Test event created successfully! Check the logs for details.";
  } catch (error) {
    Logger.log("❌ Error creating test calendar event: " + error.toString());
    return "Error creating test event: " + error.toString();
  }
}
