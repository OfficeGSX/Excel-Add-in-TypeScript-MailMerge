// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. 
// See full license at the root of the repo.

import * as $ from 'jquery';
import 'core-js';
import 'fabric';
import { Storage, Dialog, Utilities } from '@microsoft/office-js-helpers';

/**
* The Preview class serves methods for the mail preview functionality.
*/
class Preview {

    firstRowDataContainer: Storage<any>;
    selectedEmailContentContainer: Storage<any>;
    selectedTemplateContainer: Storage<any>;
    profileInfoContainer: Storage<any>;

    constructor() {
        this.firstRowDataContainer = new Storage('FirstRowData');
        this.selectedEmailContentContainer = new Storage('SelectedEmailContent');
        this.selectedTemplateContainer = new Storage('SelectedTemplate');
        this.profileInfoContainer = new Storage('ProfileInfo');
        this.initialize();
		$('#close-button').click(() => Dialog.close('close'));
    }

    private initialize() {
        var senderDisplayName = this.profileInfoContainer.get(this.profileInfoContainer.keys()[0]);
        var senderEmailAddress = this.profileInfoContainer.keys()[0];
        var recipientEmail = this.firstRowDataContainer.keys()[0];
        var subject = this.selectedTemplateContainer.get(this.selectedTemplateContainer.keys()[0]);  
        var placeHolders = this.firstRowDataContainer.get(this.firstRowDataContainer.keys()[0]);
        var emailBody = this.selectedEmailContentContainer.get(subject).content;
        var sentDateTime = Date.now();
        
        // In the email body, replace the placeholders with the values from the Excel table.
        for (var key in placeHolders) {
            if (key != "EmailAddress") {
                emailBody = this.replaceAll(emailBody, `&lt;${key}&gt;`, placeHolders[key]);
            }         
        }

		// Create the message view in a dialog window.
        $('#header').append('<h3>' + senderDisplayName + '</h2>');
        $('#header').append('<p>' + 'Subject:  ' + subject + '</p>');
        $('#header').append('<p>' + 'To:  ' + recipientEmail + '</p>');
        $('#header').append('<hr />');
        $('#body').append(emailBody);
    }

	/**
	* Searches a string for a value, and replaces
	* with another value.
	* 
	* @param input The string to search in.
	* @param search The string to search for.
	* @param replacement The new value to replace with.
	* return Returns a new string with the search string replaced with replacement.
	*/
    replaceAll(input, search, replacement) {
        return input.split(search).join(replacement);
    }
}

let preview = new Preview();