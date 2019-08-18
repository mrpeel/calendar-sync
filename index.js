const pMap = require("p-map")
const delay = require("delay")
const { google } = require("googleapis");
const h2p = require('html2plaintext');


if (process.env.NODE_ENV !== "production") {
    require("dotenv").config();
}

const randomDelay = () => {
    const ms = Math.floor(Math.random() * 10000);
    // console.log('delaying', ms);
    return delay(ms);
};

/*
    Auth with Google calendar
*/
const authGcal = () => {
    const jsonStr = Buffer.from(process.env.GOOGLE_APPLICATION_CREDENTIALS_BASE64, 'base64').toString();
    const authKeys = JSON.parse(jsonStr);
    const auth = google.auth.fromJSON(authKeys);
    auth.scopes = ['https://www.googleapis.com/auth/calendar.events'];
    return google.calendar({
        auth,
        version: 'v3'
    });
};

/* 
    Return the max num event starting at cut-off time (one hour ago)
    from the Google calendar
*/
const getGcalEvents = async (gCal, minDate, maxDate, maxNumEvents) => {
    const params = {
        timeMin: minDate.toISOString(),
        timeMax: maxDate.toISOString(),
        maxResults: maxNumEvents,
        calendarId: process.env.GOOGLE_CALENDAR_ID,
    };
    const events = await gCal.events.list(params);
    return events.data.items;
};

/* 
    Filter Outlook events against a cut-off time (one hour ago), remove the leave organiser 
    which puts leave events in the calendar that shouldn't bbe synced, and limit by the maximum
    number of events to sync at one time
*/
const filterOutlookEvents = (events, minDate, removeOrganizer, maxNumEvents) => {
    // Filter out events before the cut off or with the wrong organizer
    console.log(`Received ${events.length || 0} Outlook events`);
    let filteredEvents = events.filter(event => (new Date(event.Start).toISOString() >= minDate.toISOString()
        && event.Organizer != removeOrganizer
        && event.RequiredAttendees.indexOf(removeOrganizer) == -1));
    console.log(`Filtered to  ${filteredEvents.length} events`);

    // Order events by start time
    filteredEvents.sort((a, b) => {
        return new Date(a.Start).getTime() - new Date(b.Start).getTime();
    });
    // Only take max num events
    if (filteredEvents.length > maxNumEvents) {
        filteredEvents = filteredEvents.slice(0, maxNumEvents);
    }
    return filteredEvents;
};

/* 
    Loop through each Google Calendar event looking for the iCalUID in the outlook 
    event array.  All events which aren't found in outlook, must have been deleted
    so delete the events from Google calendar
*/
const deleteGCalEvents = async (gCal, outlookEvents, gCalEvents) => {
    const deletedEvents = gCalEvents.filter(gCalEvent =>
        !outlookEvents.find(outlookEvent => gCalEvent.iCalUID === outlookEvent.Id)
    );

    if (!deletedEvents.length) {
        return;
    }

    console.log(`deleting ${deletedEvents.length} events`);

    const deleteEvent = async (event) => {
        await randomDelay();
        console.log(`deleting ${event.id}`);
        const params = {
            calendarId: process.env.GOOGLE_CALENDAR_ID,
            eventId: event.id
        };
        return gCal.events.delete(params);
    };

    return pMap(deletedEvents, deleteEvent, { concurrency: 5 });
};

/*
    Execute the Google import function.  This imports and updates or creates events based on their
    iCalUID.
*/
const importGCalEvents = async (gCal, outlookEvents) => {
    if (!outlookEvents.length) {
        return;
    }

    console.log(`found ${outlookEvents.length} events`);

    const importEvent = async event => {
        await randomDelay();
        console.log(`importing '${event.Subject}'`);
        const params = {
            calendarId: process.env.GOOGLE_CALENDAR_ID,
            resource: {
                iCalUID: event.Id,
                start: {
                },
                end: {
                },
                summary: event.Subject,
                description: h2p(event.Body),
                location: event.Location
            }
        };
        // Set paraemeter to either a date for all-day events, or a date-time for normal events
        if (event.isAllDay) {
            params.resource.start.date = new Date(event.Start).toISOString().slice(0, 10);
            params.resource.end.date = new Date(event.End).toISOString().slice(0, 10);

        } else {
            params.resource.start.dateTime = new Date(event.Start).toISOString();
            params.resource.end.dateTime = new Date(event.End).toISOString();
        }
        return gCal.events.import(params)
            .catch(err => {
                console.error(`event ${event.uid} failed`)
                console.error(err)
            });
    }

    return pMap(outlookEvents, importEvent, { concurrency: 5 })
        .then(() => outlookEvents.map(event => event.uid));
}


const processCalendars = async (outlookEvents) => {
    const minDate = new Date();
    minDate.setHours(minDate.getHours() - 1);

    const maxDate = new Date();
    maxDate.setDate(maxDate.getDate() + process.env.MAX_NUM_DAYS || 180);

    let removeOrganizer = process.env.REMOVE_ORGANIZER || "";
    let maxNumEvents = process.env.MAX_NUM_EVENTS || 2500;

    let filteredOutlookEvents = filterOutlookEvents(outlookEvents, minDate, removeOrganizer, maxNumEvents);

    let gCal = await authGcal();
    let gCalEvents = await getGcalEvents(gCal, minDate, maxDate, maxNumEvents);
    console.log(`Retrieved ${gCalEvents.length} Google Calendar events`);

    await deleteGCalEvents(gCal, filteredOutlookEvents, gCalEvents);
    await importGCalEvents(gCal, filteredOutlookEvents);
};


let syncHandler = async (event, context, callback) => {
    try {
        const requireds = [
            'GOOGLE_CALENDAR_ID',
            'GOOGLE_APPLICATION_CREDENTIALS_BASE64',
            'MAX_NUM_EVENTS',
        ];

        const hasEnv = requireds.every(env => env in process.env)
        if (!hasEnv) {
            throw new Error(`missing required env vars: ${requireds}`);
        }

        let incomingEventJson = JSON.parse(event.body)

        if (incomingEventJson.events) {
            console.log(`Received ${incomingEventJson.events.length} calendar events for processing...`);
            callback(null, {
                statusCode: 200,
                body: JSON.stringify({
                    message: "Processing initiated"
                })
            });
            await processCalendars(incomingEventJson.events);
        } else {
            console.log(`No events received.  Nothing to process.`);
            callback(null, {
                statusCode: 200,
                body: JSON.stringify({
                    message: "No records to process"
                })
            });
        }
        console.log(`Finished processing`);

        context.succeed();
    } catch (err) {
        console.log(err);
        callback(null, {
            statusCode: 500,
        });
        context.fail();
    }
};

const retrieveGoogleEvent = async (gCal, id) => {
    const params = {
        iCalUID: id,
        calendarId: process.env.GOOGLE_CALENDAR_ID,
    };
    const events = await gCal.events.list(params);
    let returnVal;
    if (events.data.items.length) {
        returnVal = events.data.items[0];
    }

    return returnVal;
}

const importSingleEvent = async (gCal, event) => {
    console.log(`importing '${event.Subject}'`);
    const params = {
        calendarId: process.env.GOOGLE_CALENDAR_ID,
        resource: {
            iCalUID: event.Id,
            start: {
            },
            end: {
            },
            summary: event.Subject,
            description: h2p(event.Body),
            location: event.Location,
        }
    };
    // Set paraemeter to either a date for all-day events, or a date-time for normal events
    if (event.isAllDay) {
        params.resource.start.date = new Date(event.Start).toISOString().slice(0, 10);
        params.resource.end.date = new Date(event.End).toISOString().slice(0, 10);

    } else {
        params.resource.start.dateTime = new Date(event.Start).toISOString();
        params.resource.end.dateTime = new Date(event.End).toISOString();
    }
    return gCal.events.import(params)
        .catch(err => {
            console.error(`event ${event.uid} failed`)
            console.error(err)
        });
};

const deleteSingleEvent = async (gCal, eventId) => {
    console.log(`deleting ${eventId}`);
    const params = {
        calendarId: process.env.GOOGLE_CALENDAR_ID,
        eventId: eventId
    };
    return gCal.events.delete(params);
};


let eventHandler = async (event, context, callback) => {
    try {
        const requireds = [
            'GOOGLE_CALENDAR_ID',
            'GOOGLE_APPLICATION_CREDENTIALS_BASE64',
        ];

        const hasEnv = requireds.every(env => env in process.env)
        if (!hasEnv) {
            throw new Error(`missing required env vars: ${requireds}`);
        }

        let removeOrganizer = process.env.REMOVE_ORGANIZER || "";
        let incomingEventJson = JSON.parse(event.body)

        /*
            Events use either camel case or Pascal case depending on which version of the event trigger is used
            in MS flow.  Each of the uses of a property attempts to use whichever version was supplied.
        */
        incomingEventJson.Id = incomingEventJson.Id || incomingEventJson.id;
        incomingEventJson.Organizer = incomingEventJson.Organizer || incomingEventJson.organizer;
        incomingEventJson.RequiredAttendees = incomingEventJson.RequiredAttendees || incomingEventJson.requiredAttendees;
        incomingEventJson.Subject = incomingEventJson.Subject || incomingEventJson.subject;
        incomingEventJson.Body = incomingEventJson.Body || incomingEventJson.body;
        incomingEventJson.Location = incomingEventJson.Location || incomingEventJson.location;
        incomingEventJson.Start = incomingEventJson.Start || incomingEventJson.start;
        incomingEventJson.End = incomingEventJson.End || incomingEventJson.end;

        // Check if this is an event of the type which we ignore (no action or has the wrong organizer / attendees)
        if (incomingEventJson.Organizer != removeOrganizer
            && incomingEventJson.RequiredAttendees.indexOf(removeOrganizer) == -1
            && incomingEventJson.ActionType) {
            let gCal = await authGcal();

            if (incomingEventJson.ActionType == "deleted") {
                // perform delete operation
                let googleEvent = await retrieveGoogleEvent(gCal, incomingEventJson.Id);
                if (googleEvent) {
                    await deleteSingleEvent(gCal, googleEvent.id);
                } else {
                    console.log(`Failed to perform delete - unable to find event icalUID ${incomingEventJson.Id}`);
                }
            } else {
                // perform upsert (Google Cal import) operation - if this is not a filtered out 
                await importSingleEvent(gCal, incomingEventJson);
                console.log(`Event id ${incomingEventJson.Id} processed`);
            }

            callback(null, {
                statusCode: 200,
                body: JSON.stringify({
                    message: `Event id ${incomingEventJson.Id} processed`
                })
            });

        } else {
            console.log(`Action type '${incomingEventJson.ActionType || ''}' ignored for organiser '${incomingEventJson.Organizer || ''}' and attendees '${incomingEventJson.RequiredAttendees || ''}'`);
            callback(null, {
                statusCode: 200,
                body: JSON.stringify({
                    message: "No action specified"
                })
            });
        }
        console.log(`Finished processing`);

        context.succeed();
    } catch (err) {
        console.log(err);
        callback(null, {
            statusCode: 500,
        });
        context.fail();
    }
};

module.exports = {
    syncHandler: syncHandler,
    eventHandler: eventHandler,
};





