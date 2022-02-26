# outlook-cors-sample

This repository contains an example Outlook add-in that uses event-based activation to authenticate the user, get user data from the Microsoft Graph using CORS and insert the information into the signature of the email.

## References

This example is based on the SSO quickstart available here:<br>
https://docs.microsoft.com/en-us/office/dev/add-ins/quickstarts/sso-quickstart

Documentation for event-based activation can be found here:<br>
https://docs.microsoft.com/en-us/office/dev/add-ins/outlook/autolaunch

Guidence for enabling SSO in event-based addins:<br>
https://docs.microsoft.com/en-us/office/dev/add-ins/outlook/use-sso-in-event-based-activation

office-addin-sso authentication backend implementation:<br>
https://github.com/OfficeDev/Office-Addin-Scripts/blob/master/packages/office-addin-sso/src/authRoute.ts

## Setup

Replace instances of INSERT_CLIENT_ID with your App Registration ID.

Replace instances of INSERT_TOKEN_EXCHANGE_SERVER with your token exchange server. If you are running the server using `localhost` and it works in the taskpane but not in the Outlook for Windows event-based runtime, then host it under another domain locally or deploy the server on another host. Make sure the server uses HTTPS.
