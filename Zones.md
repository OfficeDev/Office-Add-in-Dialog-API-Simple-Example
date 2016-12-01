##Troubleshoot: An Office Add-in Dialog cannot be displayed
##Affected applications and symptoms
**Applicable Office Applications**: 
- Affected: Online(Web) versions of Word, Excel, PowerPoint, OneNote, Outlook. 
- Not affected: Windows, Mac and IPad versions of Office

**Applicable Browsers**: 
- Affected: Internet Explorer and Edge. 
- Not affected: Chrome, Firefox and Safari. 

**Symptoms**: 
When using an Office add-in, you are asked to allow a dialog to be displayed. Upon allowing the dialog your receive the following message: 

```
The security settings in your browser prevent us from creating a dialog box. Try a different browser, or configure your browser, so that [URL] and the domain shown in your address bar are in the same security zone
```

![](http://i.imgur.com/3mqmlgE.png)

##Resolution
###End users and Administrators: Add add-in domain to trusted sites

Often times, the easiest way to resolve this issue is to add the domain of the add-in site to the list of **Trusted Sites** 

**Important Security Reminder: Do NOT add sites of Add-ins you don't trust**

2. While in Internet Explorer go to **Internet Options>Settings>Security** (You need to do this from Internet Explorer even if you are using Edge)
![](http://i.imgur.com/JwJLPg0.png)

3. Select the **Trusted Sites** zone and click on **Sites**. 
4. Enter the URL the warning originally gave you into the textbox and click "Add" 

![](http://i.imgur.com/ytHeuBZ.png)

5. Try to use the add-in again. If the problem persist you may need to verify the settings of all other security zones and ensure the add-in domain is on the same zone as the URL that is displayed in your addressbar for the office application.

###Developers: Use displayInFrame
The issue depicted in this article only occurs when the [Dialog API](https://dev.office.com/reference/add-ins/shared/officeui.displaydialogasync) is used in pop-up mode. Use the [displayInFrame](https://dev.office.com/reference/add-ins/shared/officeui.displaydialogasync) flag to avoid hitting this complication. This alternative does require your page to support being displayed inside and IFrame. Here is an example of the code a developer could use:

```
Office.context.ui.displayDialogAsync(startAddress, {displayInFrame:true}, callback);
```

##Additional Technical information
Internet Explorer has a security feature called "Security Zones". Websites on different security zones, particularly **sites that have different Protected Mode settings cannot communicate with each** other via the browser. 

When developers use the Dialog API in pop-up mode, a new window is created and navigated to the URL of the add-in. If the domain of the dialog URL is in a different security zone (with different protected mode setting) as the domain of the website hosting the Office Online aplication, the dialog cannot communicate back with Office. 