# Script Converter Tool

## TLDR;
Tool created for accessing script on the cloud and converting that data into useable object
for later use down the pipeline.

## Credits
Programmer / Designer - **Jeffrey M Johnson**

## What
Director of play had live document that he updated daily with data the application being built
needed. Not wanting to have copies floating around and recognizing the customer had no desire 
to switch tools, I made this tool to bridge the gap.

It access' the script spreadsheet file using OAuth2.0 and then sifts through the data, converting 
what is needed to an xml schema designed for this purpose.

## How
#### Tools / Libraries
* C# Console Application

#### Interesting Code
* [Script Class](https://github.com/JeffreyMJohnson/ScriptXMLConverter/blob/master/Script.cs)
* [GoogleDataSource Class](https://github.com/JeffreyMJohnson/ScriptXMLConverter/blob/master/GoogleDataSource.cs)

## Current Status
This tool was created for a one-time use on a short-term project that is now complete.

