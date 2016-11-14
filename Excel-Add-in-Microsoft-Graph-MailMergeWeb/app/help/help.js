﻿/* Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See LICENSE in the project root for license information. */

(function () {
    "use strict";
    Office.initialize = function (reason) {
        $(document).ready(function () {
            $('#cancel-button').click(messageParent);
        });
    };

    // Sends a message back to the calling API to close the dialog window.
    function messageParent() {
        Office.context.ui.messageParent('ok');
    }

}());
