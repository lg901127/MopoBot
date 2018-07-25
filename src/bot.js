'use strict';

module.exports.setup = function(app) {
    var builder = require('botbuilder');
    var teams = require('botbuilder-teams');
    var config = require('config');
    var botConfig = config.get('bot');
    
    // Create a connector to handle the conversations
    var connector = new teams.TeamsChatConnector({
        // It is a bad idea to store secrets in config files. We try to read the settings from
        // the environment variables first, and fallback to the config file.
        // See node config module on how to create config files correctly per NODE environment
        appId: process.env.MICROSOFT_APP_ID || botConfig.microsoftAppId,
        appPassword: process.env.MICROSOFT_APP_PASSWORD || botConfig.microsoftAppPassword
    });
    
    // Define a simple bot with the above connector that echoes what it received
    // var bot = new builder.UniversalBot(connector, function(session) {
    //     // Message might contain @mentions which we would like to strip off in the response
    //     var text = teams.TeamsMessage.getTextWithoutMentions(session.message);
    //     session.send('You said: %s', text);
    // });
    var bot = new builder.UniversalBot(connector);
    // var stripBotAtMentions = new teams.StripBotAtMentions();
    // bot.use(stripBotAtMentions);
    bot.dialog('/', [
        function(session) {
            builder.Prompts.choice(session, 'Pick your choice', ['List all users', 'Add user', 'Delete user', 'Get Policies Assigned To Users']);
        },
        function(session, results) {
            switch (results.response.index) {
                case 0:
                    session.send('List users');
                    break;
                case 1:
                    session.send('Add user');
                    break;
                case 2:
                    session.send('Delete user');
                    break;
                case 3: 
                    bot.dialog('policiesAssignDialog', [
                      function (session) {
                        session.beginDialog('askName');
                      },

                      function (session, results) {
                        switch(results.response) {
                          case 'Arabic Test': 
                            session.endDialog('Assigned Policies: Messaging Policy - Teams Messaging Policy 1520270091696, Meeting policy - RestrictedAnonymousAccess');
                          default: 
                            session.endDialog('We cannot find any assigned policies for the specified user');
                        }
                      }
                    ]);

                    bot.dialog('askName', [
                      function (session) {
                        builder.Prompts.text(session, 'Please enter the user name');
                      }, 

                      function (session, results) {
                        session.endDialogWithResult(results);
                      }
                    ]);
                default:
                    session.send('default');
                    break;
            }
        }
    ])

    // Setup an endpoint on the router for the bot to listen.
    // NOTE: This endpoint cannot be changed and must be api/messages
    app.post('/api/messages', connector.listen());

    // Export the connector for any downstream integration - e.g. registering a messaging extension
    module.exports.connector = connector;
};
