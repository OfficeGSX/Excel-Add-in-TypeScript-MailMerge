// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the See full license at the root of this repo.

import { Authenticator, DefaultEndpoints, IToken } from '@microsoft/office-js-helpers';
import { Storage } from '@microsoft/office-js-helpers';

export class OutlookHelper {

    authenticator: Authenticator;
    token: IToken;
    templatesFolderId: string;
    emailTemplates: string;
    profileInfoContainer: Storage<any>;
    firstRowDataContainer: Storage<any>;
    selectedEmailContentContainer: Storage<any>;
    selectedTemplateContainer: Storage<any>;
    emailAddressesContainer: Storage<any>;
    excelDataContainer: Storage<any>;


    constructor() {
        this.profileInfoContainer = new Storage('ProfileInfo');
        this.authenticator = new Authenticator();
        this.token = this.authenticator.tokens.get(DefaultEndpoints.Microsoft);
    }

    // Get the user profile of the current user.
    getUserProfile() {
        // Create the request URL.
        var requestUrl = 'https://graph.microsoft.com/v1.0/me';

        return $.ajax({
            method: "GET",
            url: requestUrl,
            dataType: "json",
            headers: {
                Authorization: 'Bearer ' + this.token.access_token
            }
        })
            .done((data) => {
                this.profileInfoContainer.clear();
                this.profileInfoContainer.insert(data.mail, data.displayName)
            })
            .fail((jqXHR, textStatus, errorThrown) => {
                var response = $.parseJSON(jqXHR.responseText);
                console.log(JSON.stringify(jqXHR));
            });
    }

    // Get the email messages in the templates folder from Outlook.
    getTemplates(folderId: string) {

        // Create the request URL.
        var requestUrl = 'https://graph.microsoft.com/v1.0/me/mailFolders/' + folderId + '/messages';

        return $.ajax({
            method: "GET",
            url: requestUrl,
            dataType: "json",
            headers: {
                Authorization: 'Bearer ' + this.token.access_token
            }
        })
            .done(data => {
                var emailSubjects: any;
                if (data == null || data.value == null) return null;
                emailSubjects = data.value.map(value => data.value.subject);
            })
            .fail((jqXHR, textStatus, errorThrown) => {
                var response = $.parseJSON(jqXHR.responseText);
            });
    }

    // Get all mail folders in Outlook.
    getOutlookFolders() {
        // Create the request URL.
        var requestUrl = 'https://graph.microsoft.com/v1.0/me/mailFolders/';

        return $.ajax({
            method: "GET",
            url: requestUrl,
            dataType: "json",
            headers: {
                Authorization: 'Bearer ' + this.token.access_token
            }
        })
    }

    // Helper function.
    replaceAll(input, search, replacement) {
        return input.split(search).join(replacement);
    }

    // Send a single message.
    sendMail(requestUrl, data) {
        return $.ajax({
            method: "POST",
            url: requestUrl,
            dataType: "json",
            headers: {
                Authorization: 'Bearer' + this.token.access_token,
                'Content-Type': 'application/json'
            },
            data: JSON.stringify(data)
        })

    }

    // Send all of the mail merged messages.
    sendMessages() {
        this.selectedEmailContentContainer = new Storage('SelectedEmailContent');
        this.selectedTemplateContainer = new Storage('SelectedTemplate');
        this.emailAddressesContainer = new Storage('EmailAddresses');
        this.firstRowDataContainer = new Storage('FirstRowData');
        this.excelDataContainer = new Storage('Data');

        var placeHolders = this.excelDataContainer.values();

        var subject = this.selectedTemplateContainer.get(this.selectedTemplateContainer.keys()[0]);
        var emailBody = this.selectedEmailContentContainer.get(subject).content;
        var sentDateTime = Date.now();

        for (var i = 0; i < this.emailAddressesContainer.keys().length; i++) {
            var emailBody = this.selectedEmailContentContainer.get(subject).content;

            var placeHolder = placeHolders[i];
            for (var key in placeHolder) {
                if (key != "EmailAddress") {
                    emailBody = this.replaceAll(emailBody, `&lt;${key}&gt;`, placeHolder[key]);
                    console.log(`&lt;${key}&gt` + "   " + placeHolder[key]);
                }
            }

            var emailAddress = this.emailAddressesContainer.get(this.emailAddressesContainer.keys()[i]);

            var requestUrl = 'https://graph.microsoft.com/v1.0/me/sendMail';

            var data = {
                "Message": {
                    "subject": subject,
                    "body": {
                        "contentType": "html",
                        "content": emailBody
                    },
                    "toRecipients": [
                        {
                            "emailAddress": {
                                "address": emailAddress
                            }
                        }
                    ],
                    "isDeliveryReceiptRequested": true,
                    "isReadReceiptRequested": true
                },
                "SaveToSentItems": true
            };

            this.sendMail(requestUrl, data)
                .then(() => {
                    console.log("Mail sent");
                })
        }
    }
}