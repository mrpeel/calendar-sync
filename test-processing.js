let index = require('./index');
let completeEvents = require('./samples/complete-events');

let processTestInvoke = async () => {
    await index.apiHandler({
        "body": JSON.stringify(completeEvents)
    }, null, () => { console.log(`done`); });
};

processTestInvoke();

let processTest1 = async () => {
    await index.eventHandler(
        {
            "body": JSON.stringify({
                "ActionType": "added",
                "Subject": "Two dayer test",
                "Start": "2019-08-18T05:15:00.0000000+00:00",
                "End": "2019-08-18T05:45:00.0000000+00:00",
                "ShowAs": 0,
                "Recurrence": 2,
                "ResponseType": 1,
                "ResponseTime": "0001-01-01T00:00:00+00:00",
                "Importance": 0,
                "Id": "AAMkADAwYzBmZGI0LTQwZjItNDQ3NS1hMTcyLWNjN2JkZTA5NjZiMAFRAAgI1yNvA04AAEYAAAAAv88qugCiDEW1BqUBZQ7fbwcA9jNec41JiEGOicqeWIomlgAAAIeglAAAwqvbQhkSS0qCWki0oEm71wABzjykggAAEA==",
                "DateTimeCreated": "2019-08-17T03:21:37.9642298+00:00",
                "DateTimeLastModified": null,
                "Organizer": "neil.kloot@jbhifi.com.au",
                "TimeZone": "AUS Eastern Standard Time",
                "SeriesMasterId": "AAMkADAwYzBmZGI0LTQwZjItNDQ3NS1hMTcyLWNjN2JkZTA5NjZiMAFRAAgI1yNvA04AAEYAAAAAv88qugCiDEW1BqUBZQ7fbwcA9jNec41JiEGOicqeWIomlgAAAIeglAAAwqvbQhkSS0qCWki0oEm71wABzjykggAAEA==",
                "Categories": [],
                "RequiredAttendees": "",
                "OptionalAttendees": "",
                "ResourceAttendees": "",
                "Body": "<html><head><meta name=\"Generator\" content=\"Microsoft Exchange Server\">\r\n<!-- converted from text -->\r\n<style><!-- .EmailQuote { margin-left: 1pt; padding-left: 4pt; border-left: #800000 2px solid; } --></style></head>\r\n<body>\r\n<font size=\"2\"><span style=\"font-size:11pt;\"><div class=\"PlainText\">&nbsp;</div></span></font>\r\n</body>\r\n</html>\r\n",
                "IsHtml": true,
                "Location": "",
                "IsAllDay": false,
                "RecurrenceEnd": "2019-08-18T00:00:00+10:00",
                "NumberOfOccurrences": null,
                "Reminder": 15,
                "ResponseRequested": true
            })
        }
        , null, () => { console.log(`done`); });
};

let processTest2 = async () => {
    await index.eventHandler(
        {
            "body": JSON.stringify({
                "ActionType": "deleted",
                "Subject": null,
                "Start": "",
                "End": "",
                "ShowAs": 0,
                "Recurrence": 0,
                "ResponseType": 0,
                "ResponseTime": null,
                "Importance": 0,
                "Id": "AAMkADAwYzBmZGI0LTQwZjItNDQ3NS1hMTcyLWNjN2JkZTA5NjZiMAFRAAgI1yNvA04AAEYAAAAAv88qugCiDEW1BqUBZQ7fbwcA9jNec41JiEGOicqeWIomlgAAAIeglAAAwqvbQhkSS0qCWki0oEm71wABzjykggAAEA==",
                "DateTimeCreated": null,
                "DateTimeLastModified": null,
                "Organizer": null,
                "TimeZone": null,
                "SeriesMasterId": null,
                "Categories": null,
                "RequiredAttendees": "",
                "OptionalAttendees": "",
                "ResourceAttendees": "",
                "Body": null,
                "IsHtml": false,
                "Location": null,
                "IsAllDay": null,
                "RecurrenceEnd": null,
                "NumberOfOccurrences": null,
                "Reminder": null,
                "ResponseRequested": null
            })
        }, null, () => { console.log(`done`); });
};

// processTest2();
