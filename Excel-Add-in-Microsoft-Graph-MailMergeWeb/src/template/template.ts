// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. 
// See full license at the root of this repo.

import * as $ from 'jquery';
import { Authenticator, Storage, IToken, Dialog, Utilities } from '@microsoft/office-js-helpers';
import { OutlookHelper } from '../helpers/index';
import { showHelp } from '../help/help';

/**
* This class serves methods that are required to show email templates.
*/
class Template {

    authenticator: Authenticator;
    token: IToken;
    emailSubjects = [];
    templateList: any;
    outlookHelper: OutlookHelper;
    selectedTemplate: string;
    selectedTemplateContainer: Storage<any>;
    selectedEmailContentContainer: Storage<any>;
    templatesContainer: Storage<any>;
    tokenContainer: Storage<any>;

    constructor() {
        this.authenticator = new Authenticator();
        this.selectedTemplateContainer = new Storage('SelectedTemplate');
        this.selectedEmailContentContainer = new Storage('SelectedEmailContent');
        this.templatesContainer = new Storage('Templates');
        this.outlookHelper = new OutlookHelper();
        this._initialize();
    }

    private _initialize() {
        this._registerHelp();
        this._registerSignOut();
        this._loadData();
    }

    /**
	* Click handler for selecting an email template in the UI.
	*/
    private _listItemSelected(element: Element) {
        this.selectedTemplateContainer.clear();
        this.selectedTemplateContainer.add('1', element.innerHTML);
        window.location.replace('../merge/index.html');
    }

	/**
	* Handles the click event for the help dialog window.
	*/
    private _registerHelp() {
        $('#openHelpDialog').click(async evt => showHelp());
    }

	/** 
	* Sign the user out by clearing the tokens.
	*/
    private _registerSignOut() {
        $('#signOut').click(() => {
            this.authenticator.tokens.clear();
            window.location.replace('../home/index.html');
        });
    }

	/**
	* Get templates from the user's Outlook and store the message body and 
	* placeholders for each template.
	*/
    private async _loadData() {
        try {
            // Get the mail template folder in Outlook.
            let folder = await this.outlookHelper.getOutlookFolders();

            // Using the template folder ID, get all the messages in that folder.
            let emails = await this.outlookHelper.getTemplates(folder.id);
            this.templatesContainer.clear();
            this.selectedEmailContentContainer.clear();

            emails.forEach(mail => {
                // Get the body of the message.
                let body = mail.bodyPreview;
                // Look for placeholders in the email body and put them in an array
                // of placeholders.
                let placeHolders = body.match(/<.+?>/gmi);

                // Remove the angle brackets from each placeholder.
                for (let index in placeHolders) {
                    placeHolders[index] = placeHolders[index].replace('<', '').replace('>', '');
                }
                // Store email subject and placeholders in local storage.
                this.templatesContainer.insert(mail.subject, placeHolders);
                this.emailSubjects.push(mail.subject);
                this.selectedEmailContentContainer.insert(mail.subject, mail.body);
            });

            if (this.emailSubjects != null) {
                // Display the email templates in the UI.
                this.emailSubjects.forEach(subject => {
                    $('#listSection').append('<div class="listItem"><i class="ms-Icon ms-Icon--mail ms-fontColor-neutralSecondary"></i><a class=" link ms-fontColor-neutralSecondary ms-fontSize-m ms-fontWeight-semibold">' + subject + '</a></div>');
                });
				// Attach a click event handler for when user selects a template, and store that template.
                $('.link').click((evt) => {
                    let target = evt.currentTarget;
                    target.classList.add('selected');
                    this._listItemSelected(evt.target);
                });
            }
            else {
                console.log('no emails');
				$('#templateError').append('<div class="ms-font-l"><p>We couldn\'t find any templates in your Mail Merge Templates folder.</p><p class="ms-fontColor-themePrimary" id="helpDialog">How do I make a template?</p>');
            }
        }
        catch (error) {
		    // Display a message in the UI that no templates were found.
            Utilities.log(error);
            $('#templateError').append('<div class="ms-font-l"><p>We couldn\'t find any templates in your Mail Merge Templates folder.</p><p class="ms-fontColor-themePrimary" id="helpDialog">How do I make a template?</p>');
        };
    }
}

/**
 * Create a new instance of the Template class and
 * start the application.
 */
new Template();
