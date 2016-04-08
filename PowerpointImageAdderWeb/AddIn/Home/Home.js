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
        var base64EncodedImageStr = $('#image').val();

        if (Office.context.requirements.isSetSupported('ImageCoercion', '1.1'))
        {
            writeDebug("This office context supports ImageCoercion 1.1");
        }
        else {
            writeDebug("This office context DOES NOT support ImageCoercion 1.1");
        }

        //if (Office.context.requirements.isSetSupported('ImageCoercion', '1.1')) {
            Office.context.document.setSelectedDataAsync(base64EncodedImageStr, {
                coercionType: Office.CoercionType.Image
            },
            function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    writeDebug("Action failed with error: " + asyncResult.error.message);
                }
            });
        //} else {
        //    writeDebug("Image Coercion Not Supported!");
        //}
    }

    function writeDebug(message) {
        $("#debug").append("<p><strong>" + message + "</strong></p>");
    }
})();