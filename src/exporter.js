'use strict';

let userEmail = 'user1@mslablgs.vm.obm-int.dc1',
    userName = 'MSLABLGS\\user1',
    password = 'L1n4g0r4',
    ews = require('ews-javascript-api'),
    ntlmXHR = require("./ntlmXHRApi"),
    ntlmXHRApi = new ntlmXHR.ntlmXHRApi(userName, password);

process.env['NODE_TLS_REJECT_UNAUTHORIZED'] = '0';

var exch = new ews.ExchangeService(ews.ExchangeVersion.Exchange2010_SP1);

exch.XHRApi = ntlmXHRApi;
exch.Credentials = new ews.ExchangeCredentials('fakeAsWeDoNtml', 'fakeAsWeDoNtml');
exch.Url = new ews.Uri('https://172.16.24.101/Ews/Exchange.asmx'); // you can also use exch.AutodiscoverUrl

var attendee =[ new ews.AttendeeInfo(userEmail)];
//create timewindow object o request avaiability suggestions for next 48 hours, DateTime and TimeSpan object is created to mimic portion of .net datetime/timespan object using momentjs
var timeWindow = new ews.TimeWindow(ews.DateTime.Now, new ews.DateTime(ews.DateTime.Now.TotalMilliSeconds + ews.TimeSpan.FromHours(48).asMilliseconds()));
exch.GetUserAvailability(attendee, timeWindow, ews.AvailabilityData.FreeBusyAndSuggestions)
    .then(function (availabilityResponse) {
      console.log('11', availabilityResponse);
      //do what you want with user availability
    }, function (errors) {
      console.log('22', errors);
        //log errors or do something with errors
    });
