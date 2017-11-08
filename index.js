'use strict';

const functions = require('firebase-functions'); // Cloud Functions for Firebase library
const DialogflowApp = require('actions-on-google').DialogflowApp; // Google Assistant helper library

const docusign = require('docusign-esign');

const userName = '<YOUR ACCOUNT EMAIL>';    // your account email
const password = '<YOUR ACCOUNT PASSWORD>'; // your account password
const integratorKey = '<INTEGRATOR KEY>';	  // your account Integrator Key (found on Preferences -> API page)
const templateId = '<TEMPLATE ID>';         // valid templateId from a template in your account
const templateRoleName = '<ROLE NAME>';     // template role that exists on above template
const baseUrl = 'https://demo.docusign.net/restapi';    // we will retrieve this

const googleAssistantRequest = 'google'; // Constant to identify Google Assistant requests

exports.dialogflowFirebaseFulfillment = functions.https.onRequest((request, response) => {
  console.log('Request headers: ' + JSON.stringify(request.headers));
  console.log('Request body: ' + JSON.stringify(request.body));

  // An action is a string used to identify what needs to be done in fulfillment
  let action = request.body.result.action; // https://dialogflow.com/docs/actions-and-parameters

  // Parameters are any entites that Dialogflow has extracted from the request.
  const parameters = request.body.result.parameters; // https://dialogflow.com/docs/actions-and-parameters

  // Contexts are objects used to track and store conversation state
  const inputContexts = request.body.result.contexts; // https://dialogflow.com/docs/contexts

  // Get the request source (Google Assistant, Slack, API, etc) and initialize DialogflowApp
  const requestSource = (request.body.originalRequest) ? request.body.originalRequest.source : undefined;
  const app = new DialogflowApp({request: request, response: response});

  // Create handlers for Dialogflow actions as well as a 'default' handler
  const actionHandlers = {
    // The default welcome intent has been matched, welcome the user (https://dialogflow.com/docs/events#default_welcome_intent)
    'input.welcome': () => {
      // Use the Actions on Google lib to respond to Google requests; for other requests use JSON
      if (requestSource === googleAssistantRequest) {
        sendGoogleResponse('Welcome to DocuSign！To whom will you send what?'); // Send simple response to user
      } else {
        sendResponse('Welcome to DocuSign！To whom will you send what?'); // Send simple response to user
      }
    },
    // The default fallback intent has been matched, try to recover (https://dialogflow.com/docs/intents#fallback_intents)
    'input.unknown': () => {
      // Use the Actions on Google lib to respond to Google requests; for other requests use JSON
      if (requestSource === googleAssistantRequest) {
        sendGoogleResponse('Sorry, would you let me know To whom will you send what?'); // Send simple response to user
      } else {
        sendResponse('Sorry, would you let me know To whom will you send what?'); // Send simple response to user
      }
    },
    // Default handler for unknown or undefined actions
    'default': () => {
        new RequestSignatureFromTemplate(parameters);       //Send DocuSign
        
        // Use the Actions on Google lib to respond to Google requests; for other requests use JSON
        // "TemplateEntity" and "ReceiverEntity" are defined in Dialogflow
      if (requestSource === googleAssistantRequest) {
        let responseToUser = {
          speech: 'DocuSign sent ' + parameters.TemplateEntity + ' to ' + parameters.ReceiverEntity, // spoken response
          displayText: 'DocuSign sent ' + parameters.TemplateEntity + ' to ' + parameters.ReceiverEntity + ':-)' // displayed response
        };
        
        sendGoogleResponse(responseToUser);
      } else {
        let responseToUser = {
          speech: 'DocuSign sent ' + parameters.TemplateEntity + ' to ' + parameters.ReceiverEntity, // spoken response
          displayText: 'DocuSign sent ' + parameters.TemplateEntity + ' to ' + parameters.ReceiverEntity + ':-)' // displayed response
        };
        sendResponse(responseToUser);
      }
    }
  };

  // If undefined or unknown action use the default handler
  if (!actionHandlers[action]) {
    action = 'default';
  }

  // Run the proper handler function to handle the request from Dialogflow
  actionHandlers[action]();

  // Function to send correctly formatted Google Assistant responses to Dialogflow which are then sent to the user
  function sendGoogleResponse (responseToUser) {
    if (typeof responseToUser === 'string') {
      app.ask(responseToUser); // Google Assistant response
    } else {
      // If speech or displayText is defined use it to respond
      let googleResponse = app.buildRichResponse().addSimpleResponse({
        speech: responseToUser.speech || responseToUser.displayText,
        displayText: responseToUser.displayText || responseToUser.speech
      });

      app.ask(googleResponse); // Send response to Dialogflow and Google Assistant
    }
  }

  // Function to send correctly formatted responses to Dialogflow which are then sent to the user
  function sendResponse (responseToUser) {
    // if the response is a string send it as a response to the user
    if (typeof responseToUser === 'string') {
      let responseJson = {};
      responseJson.speech = responseToUser; // spoken response
      responseJson.displayText = responseToUser; // displayed response
      response.json(responseJson); // Send response to Dialogflow
    } else {
      // If the response to the user includes rich responses or contexts send them to Dialogflow
      let responseJson = {};

      // If speech or displayText is defined, use it to respond (if one isn't defined use the other's value)
      responseJson.speech = responseToUser.speech || responseToUser.displayText;
      responseJson.displayText = responseToUser.displayText || responseToUser.speech;

      response.json(responseJson); // Send response to Dialogflow
    }
  }

    // Function to send correctly formatted Google Assistant responses to Dialogflow which are then sent to the user
    function RequestSignatureFromTemplate (parameters) {
        
        if (parameters) {
            // initialize the api client
            var apiClient = new docusign.ApiClient();
            apiClient.setBasePath(baseUrl);

            // create JSON formatted auth header
            var creds = '{"Username":"' + userName + '","Password":"' + password + '","IntegratorKey":"' + integratorKey + '"}';
            apiClient.addDefaultHeader('X-DocuSign-Authentication', creds);

            // assign api client to the Configuration object
            docusign.Configuration.default.setDefaultApiClient(apiClient);

            // ===============================================================================
            // Step 1:  Login() API
            // ===============================================================================
            // login call available off the AuthenticationApi
            var authApi = new docusign.AuthenticationApi();

            // login has some optional parameters we can set
            var loginOps = {};
            loginOps.apiPassword = 'true';
            loginOps.includeAccountIdGuid = 'true';
            authApi.login(loginOps, function (error, loginInfo, response) {
                if (error) {
                    console.log('Error: ' + error);
                    return;
                }

                if (loginInfo) {
                    // list of user account(s)
                    // note that a given user may be a member of multiple accounts
                    var loginAccounts = loginInfo.loginAccounts;
                    console.log('LoginInformation: ' + JSON.stringify(loginAccounts));

                    // ===============================================================================
                    // Step 2:  Create Envelope API (AKA Signature Request) from a Template
                    // ===============================================================================

                    // create a new envelope object that we will manage the signature request through
                    var envDef = new docusign.EnvelopeDefinition();
                    envDef.emailSubject = 'Please sign this document sent from Gooele Home';
                    envDef.templateId = templateId;

                    // specify receiver info
                    var signerName = 'Default Signer';
                    var signerEmail = 'default.signer@example.com';

                    if (parameters.ReceiverEntity == 'John Doe') {
                        signerName = 'John Doe';
                        signerEmail = 'john.doe@example.com';
                    }

                    // create a template role with a valid templateId and roleName and assign signer info
                    var tRole = new docusign.TemplateRole();
                    tRole.roleName = templateRoleName;
                    tRole.name = signerName;
                    tRole.email = signerEmail;

                    // create a list of template roles and add our newly created role
                    var templateRolesList = [];
                    templateRolesList.push(tRole);

                    // assign template role(s) to the envelope
                    envDef.templateRoles = templateRolesList;

                    // send the envelope by setting |status| to "sent". To save as a draft set to "created"
                    envDef.status = 'sent';

                    // use the |accountId| we retrieved through the Login API to create the Envelope
                    var accountId = loginAccounts[0].accountId;

                    // instantiate a new EnvelopesApi object
                    var envelopesApi = new docusign.EnvelopesApi();

                    // call the createEnvelope() API
                    envelopesApi.createEnvelope(accountId, {'envelopeDefinition': envDef}, function (err, envelopeSummary, response) {
                        if (error) {
                            console.log('Error: ' + error);
                            return;
                        }

                        if (envelopeSummary) {
                            console.log('EnvelopeSummary: ' + JSON.stringify(envelopeSummary));
                        }
                    });
                }
            });
        }
    }
});

// Construct rich response for Google Assistant
const app = new DialogflowApp();
