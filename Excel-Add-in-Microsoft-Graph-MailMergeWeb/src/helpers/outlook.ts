// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. 
// See full license at the root of this repo.

import { Storage } from '@microsoft/office-js-helpers';
import { request } from './request';

export class OutlookHelper {
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
    }

    /**
	* Get the user profile of the current user.
	*/
    async getUserProfile() {
        /**
         * We only extract the mail and displayName property of the result
         * as we expect to have it, else is throws an exception which will bubble up
         * back to our main view page.
         *
         * This syntax is called Object destructing (https://www.typescriptlang.org/docs/handbook/variable-declarations.html#object-destructuring)
         */
        let {mail, displayName} = await request.get<any>('https://graph.microsoft.com/v1.0/me');
        this.profileInfoContainer.clear();
        this.profileInfoContainer.insert(mail, displayName);
    }

    /**
	* Get the email messages in the templates folder from Outlook.
	* @param folderId The ID of the folder.
	*/
    async getTemplates(folderId: string) {
        /** Extract the value property of the result */
        let {value} = await request.get<any>(`https://graph.microsoft.com/v1.0/me/mailFolders/${folderId}/messages`);
        return value;
    }

    /**
	* Get all mail folders in Outlook.
	*/
    async getOutlookFolders() {
        let {value} = await request.get<any>('https://graph.microsoft.com/v1.0/me/mailFolders');
        return value.find(({displayName}) => displayName === 'Mail Merge Email Templates');
    }

    /** 
	* Helper function to replace a string with another string.
	*/
    replaceAll(input, search, replacement) {
        return input.split(search).join(replacement);
    }

    /**
	* Send all of the mail merged messages.
	*/
    sendMessages() {
        this.selectedEmailContentContainer = new Storage('SelectedEmailContent');
        this.selectedTemplateContainer = new Storage('SelectedTemplate');
        this.emailAddressesContainer = new Storage('EmailAddresses');
        this.firstRowDataContainer = new Storage('FirstRowData');
        this.excelDataContainer = new Storage('Data');

        // Store information about the emails to be sent.
        let placeHolders = this.excelDataContainer.values();
        let subject = this.selectedTemplateContainer.values()[0];

        // Build the email messages and send.
        this.emailAddressesContainer.keys().forEach((item, index) => {
            let placeHolder = placeHolders[index];
            let emailBody = this.selectedEmailContentContainer.get(subject).content;

            // Create the body with placeholder text
            for (let key in placeHolder) {
                if (key !== 'EmailAddress') {
                    emailBody = this.replaceAll(emailBody, `&lt;${key}&gt;`, placeHolder[key]);
                }
            }

            let emailAddress = this.emailAddressesContainer.get(item);
            let data = {
                'Message': {
                    'subject': subject,
                    'body': {
                        'contentType': 'html',
                        'content': emailBody
                    },
                    'toRecipients': [
                        {
                            'emailAddress': {
                                'address': emailAddress
                            }
                        }
                    ],
                    'isDeliveryReceiptRequested': true,
                    'isReadReceiptRequested': true
                },
                'SaveToSentItems': true
            };

            request.post('https://graph.microsoft.com/v1.0/me/sendMail', data);
        });
    }
}
