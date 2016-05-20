# Office Add-in DialogApi Example

_Applies to: Word 2016_

## Sample description
Sample that illustrates the use of [add-in commands](https://github.com/officedev/office-Add-in-Commands-Samples) along with [dialog API](http://dev.office.com/reference/add-ins/shared/officeui), **`Office.context.ui.displayDialog`** for simple scenarios. While the sample is designed to work with Word you can easily use the same code for other Office applications including Excel, PowerPoint and Outlook. 

The first command uses the ShowTaskpane action and then an HTML button inside the taskpane to trigger a dialog from within a taskpane. The second command uses the ExecuteFunction action to display a dialog directly.


![Office Add-in Dialog API Sample](http://i.imgur.com/EQ8jxDI.png)

## Try it out
### With Visual Studio
1.  Clone this repo and then open the SimpleDialogSample.sln in Visual Studio.
2.  Press F5 to build and deploy the sample add-in. Word launches and the add-in will be installed
3.  Click on one of the buttons

### Without Visual Studio
1. Clone this repo
2. Deploy the WebApp, SimpleWebApp folder, in your web server of choice. Make sure it supports httpS. 
3. Modify the SimpleDialogSample.xml file, point all URLs to your web server (replace the ~placeholders)
4. Register you add-in with office by sideloading it via a [network share](https://msdn.microsoft.com/EN-US/library/office/fp123503.aspx)
5.  Test and run the add-in. 

    a.  In the **Insert tab** in Word 2016, choose **My Add-ins**. 
    
    b.  In the **Office Add-ins** dialog box, choose **Shared Folder**.
    
    c.  Choose **Simple Dialog Example**>**OK**. The add-in will install and buttons displayed on the Ribbon
6. Click one of the buttons


## Additional samples
The following samples also make use of the dialog API for authentication scenarios:


- [PowerPoint Add-in in Microsoft Graph ASP.Net Insert Chart](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart)
- [Office Add-in Server Authentication Sample for ASP.net MVC](https://github.com/dougperkes/Office-Add-in-AspNetMvc-ServerAuth/tree/Office2016DisplayDialog)