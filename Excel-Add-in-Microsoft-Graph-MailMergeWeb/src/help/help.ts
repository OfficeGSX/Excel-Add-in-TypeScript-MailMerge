/* Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. 
See LICENSE in the project root for license information. */

import * as $ from 'jquery';
import { Dialog, Utilities } from '@microsoft/office-js-helpers';

class Help {
    constructor() {
        $('#cancel-button').click(() => Dialog.close('close'));
    }
}

/**
* Public method called to open the Help Dialog.
*/
export async function showHelp() {
    try {
        let dialog = new Dialog<string>(location.origin + '/dist/help/help.html');
        await dialog.result;
    }
    catch (error) {
        Utilities.log(error);
    }
}

new Help();
