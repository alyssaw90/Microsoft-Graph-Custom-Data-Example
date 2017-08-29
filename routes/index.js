const express = require('express');
const router = express.Router();
const passport = require('passport');
const MicrosoftGraph = require('@microsoft/microsoft-graph-client');
// ////const fs = require('fs');
// ////const path = require('path');

// Get the home page.
router.get('/', (req, res) => {
  // check if user is authenticated
  if (!req.isAuthenticated()) {
    res.render('login');
  } else {
    // renderSendMail(req, res);
    res.render('sendMail')
  }
});

// Authentication request.
router.get('/login',
  passport.authenticate('azuread-openidconnect', { failureRedirect: '/' }),
    (req, res) => {
      res.redirect('/');
    });

// Authentication callback.
// After we have an access token, get user data and load the sendMail page.
router.get('/token',
  passport.authenticate('azuread-openidconnect', { failureRedirect: '/' }),
    (req, res) => {
      graphHelper.getUserData(req.user.accessToken, (err, user) => {
        if (!err) {
          // req.user.profile.displayName = user.body.displayName;
          // req.user.profile.emails = [{ address: user.body.mail || user.body.userPrincipalName }];
          // renderSendMail(req, res);
          res.render('sendMail')
        } else {
          // renderError(err, res);
          res.render('sendMail')
        }
      });
    });

/* START OF CALENDAR ROUTES*/
router.get('/calendars', (req,res) => {
  var client = MicrosoftGraph.Client.init({
    authProvider: (done) => {
      done(null, req.user.accessToken)
    }
  });
  client
    .api('/me/calendars')
    .get((err, response) => {
      if(!err) {
        console.log(response.value)
        res.render('sendMail', {data: response.value})
      } else {
        console.log(err)
        res.render('sendMail')
      }
    })
})
/* END OF CALEDNAR ROUTES */

/* START OF EVENT ROUTES */
router.get('/events', (req, res) => {
  var client = MicrosoftGraph.Client.init({
    authProvider: (done) => {
      done(null, req.user.accessToken)
    }
  });
  client 
    .api('/me/events')
    .get((err, response) => {
      if(!err) {
        console.log(response.value)
        res.render('sendMail', {data: response.value})
      } else {
        console.log(err)
        res.render('sendMail')
      }
    })
})
/* END OF EVENT ROUTES */

/* START OF OPEN EXTENSION ROUTES*/
router.post('/open', (req,res) => {
  var data = {
    "@odata.type": "Microsoft.Graph.OpenTypeExtension",
    "extensionName": "Session.Tag",
    "tagName": ['RDP', 'SMB']
  };
  var client = MicrosoftGraph.Client.init({
    authProvider: (done) => {
      done(null, req.user.accessToken)
    }
  });
  client
    .api('/me/events/<EVENT_ID_GOES_HERE>/extensions')
    .post(data, (err, response) => {
      if(!err) {
        console.log(response)
        res.render('sendMail')
      } else {
        console.log(err)
        res.render('sendMail')
      }
    })
});

router.get('/open-extensions', (req,res) => {
  var client = MicrosoftGraph.Client.init({
    authProvider: (done) => {
      done(null, req.user.accessToken)
    }
  })
  client
    .api('/me/events/<SAME_EVENT_ID_FROM_ABOVE_GOES_HERE>/extensions/Microsoft.OutlookServices.OpenTypeExtension.Session.Tag')
    .get((err, response) => {
      if(!err) {
        console.log(response)
        res.render('sendMail', {data: response.value})
      } else {
        console.log(err)
        res.render('sendMail')
      }
    })
});
/* END OF OPEN EXTENSION ROUTES*/

/* START OF SCHEMA EXTENSION ROUTES*/
router.get('/schema-extensions', (req,res) => {
  var client = MicrosoftGraph.Client.init({
    authProvider: (done) => {
      done(null, req.user.accessToken)
    }
  });
  client
    .api('/schemaExtensions')
    .get((err, response) => {
      if(!err){
        console.log(response.value[response.value.length-1])
        res.render('sendMail', {data: response.value})
      } else {
        console.log(err)
        res.render('sendMail')
      }
    })
});

router.post('/schema', (req, res) => {
  var data = {
    "id":"RDPSessionTag",
    "description": "Session Tag",
    "targetTypes": [
        "Event"
    ],
    "properties": [
        {
            "name": "sessionTag",
            "type": "String"
        }
    ]
  }
  var client = MicrosoftGraph.Client.init({
    authProvider: (done) => {
      done(null, req.user.accessToken)
    }
  })
  client
    .api('/schemaExtensions')
    .post(data, (err, response) => {
      if(!err){
        console.log(response)
        res.render('sendMail')
      } else {
        console.log(err)
        res.render('sendMail')
      }
    })
  
});

router.post('/event-schema-extension', (req,res) => {
  var data = {
    "<SCHEMA_EXTENSION_ID_GOES_HERE>": {
      "sessionTag": "Session tag example"
    }
  }
  var client = MicrosoftGraph.Client.init({
    authProvider: (done) => {
      done(null, req.user.accessToken)
    }
  })
  client
    .api('/me/events/<EVENT_ID_GOES_HERE>')
    .patch(data, (err, response) => {
      if(!err) {
        console.log(response)
        res.render('sendMail')
      } else {
        console.log(err)
        res.render('sendMail')
      }
    })
})
/* END OF SCHEMA EXTENSION ROUTES*/

/* START OF OUTLOOK EXTENDED PROPERTIES ROUTES */
router.post('/event-outlook-extended-properites', (req, res) => {
  var client = MicrosoftGraph.Client.init({
    authProvider: (done) => {
      done(null, req.user.accessToken)
    }
  });
  var data = {
    "subject": "Extended Property",
    "body": {
      "contentType": "HTML",
      "content": "Test Event with Extended Property"
    },
    "start": {
        "dateTime": "2017-08-26T09:00:00",
        "timeZone": "Pacific Standard Time"
    },
    "end": {
        "dateTime": "2017-08-29T21:00:00",
        "timeZone": "Pacific Standard Time"
    },
    "attendees": [],
    "multiValueExtendedProperties": [
       {
             "id":"StringArray {66f5a359-4659-4830-9070-00050ec6ac6e} Name Session",
             "value": ["Tag1", "Tag2", "Tag3"]
       }
    ]
  };
  client
    .api('/me/events')
    .post(data, (err, response) => {
      if(!err) {
        console.log(response)
        res.render('sendMail')
      } else {
        console.log(err)
        res.render('sendMail')
      }
    })
});

router.get('/get-event-with-outlook-extended-property', (req, res) => {
  var client = MicrosoftGraph.Client.init({
    debugLogging: true,
    authProvider: (done) => {
      done(null, req.user.accessToken)
    }
  })
  client
    .api("/me/events/<ID_FROM_EVENT_CREATED_IN_ROUTE_ABOVE>")
    .expand('multiValueExtendedProperties')
    // .filter('StringArray {66f5a359-4659-4830-9070-00050ec6ac6e} Name Session')
    .get((err, response) => {
      if(!err) {
        console.log(response)
        res.render('sendMail')
      } else {
        console.log(err)
        res.render('sendMail')
      }
    })
})
/* END OF OUTLOOK EXTENDED PROPERTIES ROUTES */

router.get('/disconnect', (req, res) => {
  req.session.destroy(() => {
    req.logOut();
    res.clearCookie('graphNodeCookie');
    res.status(200);
    res.redirect('/');
  });
});

// helpers
function hasAccessTokenExpired(e) {
  let expired;
  if (!e.innerError) {
    expired = false;
  } else {
    expired = e.forbidden &&
      e.message === 'InvalidAuthenticationToken' &&
      e.response.error.message === 'Access token has expired.';
  }
  return expired;
}

function renderError(e, res) {
  e.innerError = (e.response) ? e.response.text : '';
  res.render('error', {
    error: e
  });
}

module.exports = router;