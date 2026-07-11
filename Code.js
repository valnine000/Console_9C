// ==========================================
// Code.js (SAFE FOR PUBLIC REPOSITORY)
// ==========================================

var url = CONFIG.SPREADSHEET_URL;

function doGet(e) {
  return HtmlService
    .createTemplateFromFile('Viewer')
    .evaluate()
    .setTitle('Console 9C')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
}

// ==========================================
// QUERY SPACE CALENDAR EVENT CREATOR (Server-Side)
// ==========================================
function createQuerySpaceEvent(payload) {
  try {
    const qsCalendarId = CONFIG.QUERY_SPACE_CALENDAR_ID;
    const calendar = CalendarApp.getCalendarById(qsCalendarId);

    if (!calendar) {
      throw new Error("Could not locate the Query Space calendar.");
    }

    const startTime = new Date(payload.startIso);
    const endTime = new Date(payload.endIso);

    const event = calendar.createEvent(
      payload.title,
      startTime,
      endTime,
      { description: payload.description }
    );

    return event.getId();

  } catch (error) {
    throw new Error(error.message);
  }
}

// ==========================================
// FETCH ENTRIES IN QUERY SPACE CALENDAR
// ==========================================
function fetchQuerySpacesData() {
  const calendarId = CONFIG.QUERY_SPACE_CALENDAR_ID;

  try {
    const cal = CalendarApp.getCalendarById(calendarId);
    if (!cal) {
      return { status: 'error', message: 'Query Space Calendar not found' };
    }

    const now = new Date();
    const past = new Date(now.getFullYear() - 5, 0, 1);
    const future = new Date(now.getFullYear() + 5, 0, 1);

    const events = cal.getEvents(past, future);

    const rawData = events.map(e => ({
      eventId: e.getId(),
      title: e.getTitle(),
      description: e.getDescription() || ''
    }));

    return { status: 'success', data: rawData };

  } catch (e) {
    return { status: 'error', message: e.toString() };
  }
}

// ==========================================
// QUERY SPACE CALENDAR EVENT UPDATER (Server-Side)
// ==========================================
function updateQuerySpaceEvent(payload) {
  try {
    const qsCalendarId = CONFIG.QUERY_SPACE_CALENDAR_ID;
    const calendar = CalendarApp.getCalendarById(qsCalendarId);

    if (!calendar) {
      throw new Error("Could not locate the Query Space calendar.");
    }

    const event = calendar.getEventById(payload.eventId);

    if (!event) {
      throw new Error("Event not found on Query Space Calendar. It may have been deleted.");
    }

    // Conditionally update only what is provided in the payload
    if (payload.title) {
      event.setTitle(payload.title);
    }

    if (payload.description) {
      event.setDescription(payload.description);
    }

    if (payload.startIso && payload.endIso) {
      const startTime = new Date(payload.startIso);
      const endTime = new Date(payload.endIso);
      event.setTime(startTime, endTime);
    }

    return event.getId();

  } catch (error) {
    throw new Error(error.message);
  }
}

// ==========================================
// CALENDAR EVENT UPDATER (Server-Side)
// ==========================================
function updateGoogleCalendarEvent(payload) {
  try {
    const calendar = CalendarApp.getDefaultCalendar();
    const event = calendar.getEventById(payload.eventId);

    if (!event) {
      throw new Error("Event not found on Google Calendar.");
    }

    const startTime = new Date(payload.startIso);
    const endTime = new Date(payload.endIso);

    event.setTitle(payload.title);
    event.setTime(startTime, endTime);
    event.setDescription(payload.description);

    return event.getId();

  } catch (error) {
    throw new Error(error.message);
  }
}

// ==========================================
// CALENDAR EVENT CREATOR (Server-Side)
// ==========================================
function createGoogleCalendarEvent(payload) {
  try {
    const calendar = CalendarApp.getDefaultCalendar();

    const startTime = new Date(payload.startIso);
    const endTime = new Date(payload.endIso);

    const event = calendar.createEvent(
      payload.title,
      startTime,
      endTime,
      { description: payload.description }
    );

    return event.getId();

  } catch (error) {
    throw new Error(error.message);
  }
}

// ==========================================
// GET CALENDAR EVENTS, ALL CALENDARS
// ==========================================
function getGoogleCalendarThings(startDateStr, endDateStr, lastSyncTime) {
  const calendarIds = [
    CONFIG.PRIMARY_CALENDAR_ID,
    CONFIG.QUERY_SPACE_CALENDAR_ID
  ];

  let events = [];
  let errors = [];

  for (const calendarId of calendarIds) {
    try {
      const optionalArgs = {
        timeMin: new Date(startDateStr).toISOString(),
        timeMax: new Date(endDateStr).toISOString(),
        singleEvents: true,
        orderBy: 'startTime',
        maxResults: 2500,
        eventTypes: ['default', 'focusTime', 'outOfOffice', 'workingLocation']
      };

      if (lastSyncTime) {
        optionalArgs.updatedMin = new Date(lastSyncTime).toISOString();
        optionalArgs.showDeleted = true;
      }

      let pageToken = null;
      do {
        if (pageToken) {
          optionalArgs.pageToken = pageToken;
        }

        const response = Calendar.Events.list(calendarId, optionalArgs);
        const items = response.items;

        if (items && items.length > 0) {
          const eventSource = (calendarId === CONFIG.PRIMARY_CALENDAR_ID) ? 'primary' : 'queryspace';
          for (let i = 0; i < items.length; i++) {
            const event = items[i];
            const start = event.start.dateTime || event.start.date;
            const end = event.end.dateTime || event.end.date;

            let startTime = start;
            let endTime = end;

            if (event.start.date) {
              const sParts = event.start.date.split('-');
              startTime = new Date(sParts[0], sParts[1] - 1, sParts[2]).toISOString();
            }

            if (event.end.date) {
              const eParts = event.end.date.split('-');
              endTime = new Date(eParts[0], eParts[1] - 1, eParts[2]).toISOString();
            }

            events.push({
              id: event.id,
              googleCalendarId: event.id,
              title: event.summary || 'Untitled Event',
              startTime: startTime,
              endTime: endTime,
              description: event.description || '',
              sourceCalendar: calendarId,
              source: eventSource,
              status: event.status
            });
          }
        }
        pageToken = response.nextPageToken;
      } while (pageToken);
    } catch (e) {
      errors.push({
        calendarId: calendarId,
        error: e.toString()
      });
    }
  }

  events.sort((a, b) => new Date(a.startTime) - new Date(b.startTime));

  return {
    events: events,
    errors: errors
  };
}