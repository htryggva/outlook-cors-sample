# outlook-cors-sample

This repository contains an example Outlook add-in that uses event-based activation to authenticate the user, get user data from the Microsoft Graph using CORS and insert the information into the signature of the email.

## Setup

Replace instances of `INSERT_CLIENT_ID` with your App Registration ID.

**IMPORTANT** Make sure you set the following registry value before you start Outlook:

```
HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Wef\Developer\[Add-in ID]
UseDirectDebugger 1
```

This will allow you to make XHR calls to `https://localhost`, where the token exchange server is hosted in this example.

## Notes

### UseDirectDebugger

The JS Runtime uses a different sandbox depending on if this value is set in the registry or not.

You therefore need to make sure you test your add-in with the flag set to **0** before shipping it to your users.

- UseDirectDebugger set to 1: **Developer mode**
  - Execution engine: V8
  - VS Code debugging enabled
    - [Debugging Documentation](https://docs.microsoft.com/en-us/office/dev/add-ins/outlook/debug-autolaunch)
  - Calls to https://localhost are allowed
- UseDirectDebugger set to 0: **Production mode**
  - Execution engine: Chakra (Edge Legacy)
  - VS Code debugging disabled
  - Calls to https://localhost are **not** allowed

### console.log in Runtime

Enable `RuntimeLogging` in Registry for console.log to work:

```
HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Developer\RuntimeLogging
(Default) C:\PathToLog\file.txt
```

[Runtime Logging Documentation](https://docs.microsoft.com/en-us/office/dev/add-ins/testing/runtime-logging#runtime-logging-on-windows)

## References

This example is based on the SSO quickstart available here:<br>
https://docs.microsoft.com/en-us/office/dev/add-ins/quickstarts/sso-quickstart

Documentation for event-based activation can be found here:<br>
https://docs.microsoft.com/en-us/office/dev/add-ins/outlook/autolaunch

Guidence for enabling SSO in event-based addins:<br>
https://docs.microsoft.com/en-us/office/dev/add-ins/outlook/use-sso-in-event-based-activation

office-addin-sso authentication backend implementation:<br>
https://github.com/OfficeDev/Office-Addin-Scripts/blob/master/packages/office-addin-sso/src/authRoute.ts
