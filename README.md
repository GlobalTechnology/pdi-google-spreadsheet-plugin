pdi-google-spreadsheet-plugin
=============================

Plugin for Pentaho Data Integration(Kettle) allowing reading and writing of Google Spreadsheets.

Limitations
-----------
* This plugin can currently only read/write to existing spreadsheets/worksheets.
* The first row of a worksheet is always considered the header row and is used to name the fields read.
* Writing to a Worksheet completely replaces existing content.

See [Google Spreadsheets API](https://developers.google.com/google-apps/spreadsheets/) for more limitations on the API.

Authentication
--------------
This plugin uses a mechanism of OAuth 2.0 called [Service Accounts](https://developers.google.com/accounts/docs/OAuth2ServiceAccount).
Service accounts were designed for Server to Server authentication.

A Service Account can be created through the [Google API Console](https://code.google.com/apis/console). The plugin
requires the Service Account E-Mail address as well as the Service Account Private Key(.p12) file.

Permissions to read/write specific Spreadsheets are granted by Sharing the Spreadsheet (or containing Drive folder)
with the Service Account E-Mail Address.

Building
--------
This Plugin is built with Maven.
```
$ git clone git@github.com:GlobalTechnology/pdi-google-spreadsheet-plugin.git
$ cd pdi-google-spreadsheet-plugin
$ mvn package
```

This will produce a Kettle plugin in `target/pdi-google-spreadsheet-plugin-{version}.zip`.
This file can be extracted into your Pentaho Data Integrations plugin directory.

Step Configuration
------------------
| Property               | Description |
|:-----------------------|:------------|
| Email Address          | Service Account E-Mail Address. This is provided in the Google API Console. |
| Private Key (p12) file | Private Key associated with the Email Address Entered. This is provided when the Service Account is created in the Google API Console. The Client Id of the Private Key will be displayed and should match the beginning of the Email Address |
| Spreadsheet Key        | Unique Spreadsheet Key. *Browse* will present a list of all Spreadsheets the current Service Account has access to. |
| Worksheet Id           | Worksheet Id. *Browse* will list all Worksheets in selected Spreadsheet  |

Clicking **Test Connection** on the Service Account tab will test if the provided Email and Private Key can be used to access Google Spreadsheets.
The Connection must be successful before you can *Browse* for a Spreadsheet or Worksheet.

When using the Google Spreadsheet Input Step, you must also fetch the list of Fields before the step will return any data.
