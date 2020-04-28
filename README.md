# Example Add-in to illustrate Set-Cookies Issue on iOS

## Summary

This project contains an example Microsoft Web Add-in implementation derived from the Outlook quickstart example [here](https://docs.microsoft.com/en-us/office/dev/add-ins/quickstarts/outlook-quickstart?tabs=yeomangenerator). This example is to illustrate a bug in Outlook add-ins where on opening a dialog and navigating to an external page that passes backa  Set-Cookie header, on iOS that Set-Cookie header is not respected and no cookie is stored. This is not a problem on desktop clients e.g. Windows Outlook Desktop Client.

## Reproduce Issue

### Start Add-in Service

From this directory run the following command.

```
npm run dev-server
```

This should start the add-in server. Verify it is running by navigating to https://localhost:3000 in a browser.

### Install Manifest

Get a manifest for the add-in by navigating to https://localhost:3000/manifest.xml. Sideload this into your Office account so that it appears as a button in the add-ins menu.