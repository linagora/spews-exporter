'use strict';

let ews = require('ews-javascript-api');

var exch = new ews.ExchangeService(ews.ExchangeVersion.Exchange2013);
exch.Credentials = new ews.ExchangeCredentials('userName', 'password');
//set ews endpoint url to use
exch.Url = new ews.Uri('https://outlook.office365.com/Ews/Exchange.asmx'); // you can also use exch.AutodiscoverUrl

var attendee =[ new ews.AttendeeInfo('email1@domain.com'), new ews.AttendeeInfo('email2@domain.com')];
//create timewindow object o request avaiability suggestions for next 48 hours, DateTime and TimeSpan object is created to mimic portion of .net datetime/timespan object using momentjs
var timeWindow = new ews.TimeWindow(ews.DateTime.Now, new ews.DateTime(ews.DateTime.Now.TotalMilliSeconds + ews.TimeSpan.FromHours(48).asMilliseconds()));
exch.GetUserAvailability(attendee, timeWindow, ews.AvailabilityData.FreeBusyAndSuggestions)
    .then(function (availabilityResponse) {
        //do what you want with user availability
    }, function (errors) {
        //log errors or do something with errors
    });
