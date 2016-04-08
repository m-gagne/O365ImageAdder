/// <reference path="../App.js" />

(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();

            $('#set-image').click(insertImageIntoDocument);
        });
    };

    function insertImageIntoDocument() {
        var imageTextBox = $('#image');
        var base64EncodedImageStr = imageTextBox.val();

        if (Office.context.requirements.isSetSupported('ImageCoercion', '1.1'))
        {
            writeDebug("This office context supports ImageCoercion 1.1");
        }
        else {
            writeDebug("This office context DOES NOT support ImageCoercion 1.1");
        }

        if (Office.context.requirements.isSetSupported('ImageCoercion', '1.1')) {
            Office.context.document.setSelectedDataAsync(base64EncodedImageStr, {
                coercionType: Office.CoercionType.Image
            },
            function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    writeDebug("Action failed with error: " + asyncResult.error.message);
                }
            });
        } else {
            writeDebug("Image Coercion Not Supported! Trying to insert image into clipboard...");
            try
            {
                // append image off canvas and copy it
                $("#content-main").append('<div id="imageCanvas" src="' + base64EncodedImageStr + '"></div>');
                $("#imageCanvas").execCommand('copy');
                writeDebug("Clipboard copy successful");
                writeDebug("Ctrl+P into your document");
            }
            catch (err) {
                writeDebug("Unable to copy to clipboard");
            }
        }
    }

    function writeDebug(message) {
        $("#debug").append("<p><strong>" + message + "</strong></p>");
    }
})();