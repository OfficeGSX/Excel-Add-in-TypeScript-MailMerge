// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.

import { Storage } from '@microsoft/office-js-helpers';

// The Preview class serves methods for the preview functionality.
export class Preview {

    firstRowDataContainer: Storage<any>;
    selectedEmailContentContainer: Storage<any>;
    selectedTemplateContainer: Storage<any>;
    profileInfoContainer: Storage<any>;

    constructor() {
        this.firstRowDataContainer = new Storage('FirstRowData');
        this.selectedEmailContentContainer = new Storage('SelectedEmailContent');
        this.selectedTemplateContainer = new Storage('SelectedTemplate');
        this.profileInfoContainer = new Storage('ProfileInfo');
    }

    initialize() {
        var senderDisplayName = this.profileInfoContainer.get(this.profileInfoContainer.keys()[0]);
        var senderEmailAddress = this.profileInfoContainer.keys()[0];
        var recipientEmail = this.firstRowDataContainer.keys()[0];
        var subject = this.selectedTemplateContainer.get(this.selectedTemplateContainer.keys()[0]);  
        var placeHolders = this.firstRowDataContainer.get(this.firstRowDataContainer.keys()[0]);
        var emailBody = this.selectedEmailContentContainer.get(subject).content;
        var sentDateTime = Date.now();
        
        // Perform mail merge in the email body.
        for (var key in placeHolders) {
            if (key != "EmailAddress") {
                emailBody = this.replaceAll(emailBody, `&lt;${key}&gt;`, placeHolders[key]);
            }         
        }

        $('#header').append('<h3>' + senderDisplayName + '</h2>');
        $('#header').append('<p>' + 'Subject:  ' + subject + '</p>');
        $('#header').append('<p>' + 'To:  ' + recipientEmail + '</p>');
        $('#header').append('<hr />');
        $('#body').append(emailBody);
    }

    // Send a message to the add-in to close the dialog.
    messageParent() {
        Office.context.ui.messageParent('ok');
    }

    replaceAll(input, search, replacement) {
        return input.split(search).join(replacement);
    }


}