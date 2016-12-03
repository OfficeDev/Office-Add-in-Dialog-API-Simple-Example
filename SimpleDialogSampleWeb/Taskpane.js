/// <reference path="/Scripts/FabricUI/message.banner.js" />
/*Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. 
  See LICENSE in the project root for license information */


    "use strict";

    var messageBanner;

    // The initialize function must be defined each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {

            // Initialize the FabricUI notification mechanism and hide it
            var element = document.querySelector('.ms-MessageBanner');
            messageBanner = new app.notification.MessageBanner(element);
            messageBanner.hideBanner();

            // If not using Word 2016, use fallback logic.
            if (!Office.context.requirements.isSetSupported('WordApi', '1.1')) {
                $('#subtitle').text("Opps!");
                $("#template-description").text("Sorry, this sample requires Word 2016 or later. The button will not open a dialog.");
                $('#button-text').text("Button");
                $('#button-desc').text("Button that opens dialog only on Word 2016 or later.");
                return;
            }

            //Button that displays dialog as pop-up
            $('#button-text').text("Pick a number (pop-up)");
            $('#button-desc').text("Pick your favorite number");
            $('#action-button').click(
        openDialog);

            //Button that displays dialog as IFrame. 
            $('#button-text2').text("Pick a number (IFrame**)");
            $('#button-desc2').text("Pick your favorite number");
            $('#action-button2').click(
                openDialogAsIframe);
        });
    };

    function errorHandler(error) {
        showNotification(error);
    }

    // Display notifications in message banner at the top of the task pane.
    function showNotification(content) {
        $("#notificationBody").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }
