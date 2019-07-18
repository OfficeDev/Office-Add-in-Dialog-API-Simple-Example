---
topic: sample
products:
- office-excel
- office-word
- office-powerpoint
- office-outlook
- office-365
languages:
- javascript
- xml
extensions:
  contentType: samples
  technologies:
  - Add-ins
  createdDate: 5/13/2016 10:06:29 AM
---
# Office Add-in Dialog API Example

_Applies to: Word 2016_

## Sample description
Learn how to open popup dialogs in Office Add-ins with the [Dialog API's](http://dev.office.com/reference/add-ins/shared/officeui) **`Office.context.ui.displayDialogAsync`** method and also the use of [add-in commands](https://github.com/officedev/office-Add-in-Commands-Samples). While the sample uses Word you can easily use the same code for other Office applications including Excel, PowerPoint and Outlook. 

The first command uses the ShowTaskpane action and then an HTML button inside the taskpane to trigger a dialog from within a taskpane. The second command uses the ExecuteFunction action to display a dialog directly.


![Office Add-in Dialog API Sample](http://i.imgur.com/EQ8jxDI.png)

## Set up the project
### With Visual Studio
1.  Clone this repo and then open the SimpleDialogSample.sln in Visual Studio.
2.  Press F5 to build and deploy the sample add-in. Word launches and the add-in will be installed.

### Without Visual Studio
1. Clone this repo.
2. Deploy the SimpleDialogSampleWeb folder in your web server of choice. Make sure it supports http**S**. 
3. Modify the \SimpleDialogSample\SimpleDialogSample.xml file so that all URLs point to your web server (replace the ~placeholders)
1. Create a network share, or [share a folder to the network](https://technet.microsoft.com/en-us/library/cc770880.aspx).
2. Place a copy of the SimpleDialogSample.xml manifest file into the shared folder.
3. Launch Word and open a document.
4. Choose the **File** tab, and then choose **Options**.
5. Choose **Trust Center**, and then choose the **Trust Center Settings** button.
6. Choose **Trusted Add-ins Catalogs**.
7. In the **Catalog Url** field, enter the network path to the folder share that contains SimpleDialogSample.xml, and then choose **Add Catalog**.
8. Select the **Show in Menu** check box, and then choose **OK**.
9. A message is displayed to inform you that your settings will be applied the next time you start Microsoft Office. Close Word.
10. Restart Word and open a Word document.
2. On the **Insert** tab in Word 2016, choose **My Add-ins**.
3. Select the **SHARED FOLDER** tab.
4. Choose **Simple Dialog Exampmle**, and then select **OK**.
6. On the **Home** ribbon is a new group called **Sample Group** with a buttons labeled **Dialog from Task Pane** and **Dialog from Function**. 

## Try out the sample

1. Click the **Dialog from Task Pane** button and a task pane opens.
2. Click **Pick a number!** to open a dialog. 
3. On the dialog, click one of the number buttons. The number you chose appears in a message banner at the top of the task pane and the dialog closes.
4. Click **Pick a number!** again. 
5. Click it again immediately to see the error when a user tries to open more than one dialog.
6. Click the **X** button in the upper right of the dialog to see a sample message in the task pane that responds to the user closing the dialog.
7. Close the task pane.
8. On the ribbon, click the **Dialog from Function** button. A short message appears in the document. You may have to drag the dialog out of the way to see it.
9. On the dialog, click one of the number buttons. The number you chose appears in the document and the dialog closes.
10. Click the **Dialog from Function** button again.
11. Click the **X** button in the upper right of the dialog to see a sample message in the document that responds to the user closing the dialog.

## Additional samples
The following samples also make use of the dialog API for authentication scenarios:

- [PowerPoint Add-in in Microsoft Graph ASP.Net Insert Chart](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart)
- [Office Add-in Auth0](https://github.com/OfficeDev/Office-Add-in-Auth0)
- [Excel Add-in ASP.NET QuickBooks](https://github.com/OfficeDev/Excel-Add-in-ASPNET-QuickBooks)
- [Office Add-in Server Authentication Sample for ASP.net MVC](https://github.com/dougperkes/Office-Add-in-AspNetMvc-ServerAuth/tree/Office2016DisplayDialog)

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.
