/*Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. 
4  See LICENSE in the project root for license information */

var dialog;

function dialogCallback(asyncResult) {
    if (asyncResult.status == "failed") {

        // In addition to general system errors, there are 3 specific errors for 
        // displayDialogAsync that you can handle individually.
        switch (asyncResult.error.code) {
            case 12004:
                showNotification("Domain is not trusted");
                break;
            case 12005:
                showNotification("HTTPS is required");
                break;
            case 12007:
                showNotification("A dialog is already opened.");
                break;
            default:
                showNotification(asyncResult.error.message);
                break;
        }
    }
    else {
        dialog = asyncResult.value;
        /*Messages are sent by developers programatically from the dialog using office.context.ui.messageParent(...)*/
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, messageHandler);

        /*Events are sent by the platform in response to user actions or errors. For example, the dialog is closed via the 'x' button*/
        dialog.addEventHandler(Office.EventType.DialogEventReceived, eventHandler);
    }
}


function messageHandler(arg) {
    dialog.close();
    showNotification(arg.message);
}


function eventHandler(arg) {

    // In addition to general system errors, there are 2 specific errors 
    // and one event that you can handle individually.
    switch (arg.error) {
        case 12002:
            showNotification("Cannot load URL, no such page or bad URL syntax.");
            break;
        case 12003:
            showNotification("HTTPS is required.");
            break;
        case 12006:
            // The dialog was closed, typically because the user the pressed X button.
            showNotification("Opps, we need more information. Please reopen the dialog.");
            break;
        default:
            showNotification("Undefined error in popup window");
            break;
    }
}

function openDialog() {
    Office.context.ui.displayDialogAsync("https://localhost:44328/Dialog.html",
        { height: 50, width: 50 }, dialogCallback);
}


