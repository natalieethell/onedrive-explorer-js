# OneDrive Explorer sample web app

This sample illustrates the basic concepts for interacting with OneDrive via Microsoft Graph API to create a file explorer web app.

[See it in action!](https://dev.onedrive.com/odx/index.html)

Files included in this project:

* index.html - The controls and layout for the sample app.
* scripts\OneDriveExplorerApp.js - The JavaScript code for the application which signs in using MSAL and retrieves data from Microsoft Graph.
* scripts\Msal.js - [Microsoft Authentication library for JavaScript](https://github.com/AzureAD/microsoft-authentication-library-for-js)
* scripts\AuthenticationHandler.js - Helper file for working with Msal.js

## Getting started

To get started with this sample project, register a new application via the [Microsoft App Registration Portal](https://apps.dev.microsoft.com).

* Add the web platform to your application, and enter the redirect URL where you will host the sample. For example, `http://localhost:9999`.
* Save changes to your app.
* Copy the application ID for your newly created app into [scripts\OneDriveExplorerApp.js:5](blob/master/scripts/OneDriveExplorerApp.js#L5).

### Using IIS Express

If you have IIS Express installed, you can launch the sample app from the commandline:

```cmd
C:\Program Files\IIS Express\iisexpress.exe /Path:C:\Path\To\Sample /Port:9999
```

This will launch IIS express and start hosting the files at `C:\Path\To\Sample` on `http://localhost:9999`.
You can then navigate in your browser to this URL to use the sample.


## License

Copyright (c) Microsoft Corporation. All Rights Reserved. Licensed under the MIT [license](LICENSE.txt).

## Contributing

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.
