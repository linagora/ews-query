'use strict';

var ews = require('node-ews'),
    options = require('node-getopt').create([
      ['h', 'host=ARG', 'Exchange host'],
      ['u', 'username=ARG', 'Username'],
      ['p', 'password=ARG', 'Password'],
      ['', 'target=ARG', 'Target user to impersonate']
    ]).parseSystem().options,
    util = require('util');

ews.ignoreSSL = true;
ews.auth(options.username, options.password, options.host);

ews.run('FindItem', {
  attributes: {
    Traversal: 'Shallow'
  },
  ItemShape: {
    BaseShape: 'IdOnly'
  },
  CalendarView: {
    attributes: {
      MaxEntriesReturned: 10,
      StartDate: '2016-01-01T00:00:00Z',
      EndDate: '2016-12-31T00:00:00Z'
    }
  },
  ParentFolderIds: {
    DistinguishedFolderId: {
      attributes: {
        Id: 'calendar'
      }
    }
  }
}, [{
  RequestServerVersion: {
    attributes: {
      Version: 'Exchange2010'
    }
  },
  ExchangeImpersonation: {
    ConnectingSID: {
      SmtpAddress: options.target
    }
  }
}], function(err, result) {
  console.log(util.inspect(err));
  console.log(util.inspect(result, false, null));
});
