// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. 
// See full license at the root of the repo.

/**
* This module initializes the click handler to restart the template page.
**/
module M {
    "use strict";
    $(document).ready(() => {
        $('#restart').click(() => {
            window.location.href = '../template/index.html';
        });
    });
}