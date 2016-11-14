// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the See full license at the root of this repo.

import { OutlookHelper } from '../helpers/index';
import {Storage} from '@microsoft/office-js-helpers';
import {Authenticator, DefaultEndpoints, IToken} from '@microsoft/office-js-helpers';

// This class serves methods that are required for the email templates page.
export class Template {

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
        $('#openHelpDialog').click(evt => (window as any).app.openDialog(location.origin + "/app/help/help.html", 65, 65));
        $('#signOut').click(() => this.signOut());
        console.log('initialized');
    }

    initialize() {
        // Get the mail folders in Outlook.
        return this.outlookHelper.getOutlookFolders()
            .then(data => {
                if (data == null || data.value == null) return null;
                var folders: any[] = data.value;
                folders.forEach(folder => {
                    // Check for the folder name.
                    if (folder.displayName === "Mail Merge Email Templates") {
                        // Using the ID, get all the messages in that folder.
                        return this.outlookHelper.getTemplates(folder.id)
                            .then(data => {
                                this.templatesContainer.clear();
                                this.selectedEmailContentContainer.clear();
                                data.value.forEach(value => {
                                    var body = value.bodyPreview;
                                    // Look for placeholders in the email body.
                                    var placeHolders = body.match(/<.+?>/gmi);
                                    for (var i = 0; i < placeHolders.length; i++) {
                                        // Remove the angle brackets.
                                        placeHolders[i] = placeHolders[i].replace('<', '').replace('>', '');
                                    }
                                    // Store email subject and placeholders in local storage.
                                    this.templatesContainer.insert(value.subject, placeHolders);
                                    this.emailSubjects.push(value.subject);
                                    this.selectedEmailContentContainer.insert(value.subject, value.body);
                                })
                                return this.templatesContainer;
                            })
                            .then(() => {
                                if (this.emailSubjects != null) {
                                    // Display the email templates in the UI.
                                    this.emailSubjects.forEach(subject => {
                                        $('#listSection').append('<div class="listItem"><i class="ms-Icon ms-Icon--mail ms-fontColor-neutralSecondary"></i><a class=" link ms-fontColor-neutralSecondary ms-fontSize-m ms-fontWeight-semibold">' + subject + '</a></div>');
                                    });
                                    $('.link').click(evt => this.listItemSelected(evt.target));
                                }
                                else {
                                    console.log("no emails");
                                }
                            });
                    }
                })
            })

            .fail(error => {
                console.error(error);
                // Provides an error notification.
                $('#templateError').append('<div class="ms-font-l"><p>We couldn\'t find any templates in your Mail Merge Templates folder.</p><p class="ms-fontColor-themePrimary" id="helpDialog">How do I make a template?</p>');
                $('#helpDialog').click(evt => (window as any).app.openDialog(location.origin + "/app/help/help.html", 65, 65));
            });
    }

    signOut() {
        this.authenticator.tokens.clear();
        window.location.href = '/index.html';
    }

    // Click handler for selecting an email template in the UI.
    listItemSelected(element: Element) {
        console.log(element.innerHTML);
        this.selectedTemplateContainer.clear();
        this.selectedTemplateContainer.add('1', element.innerHTML);
        window.location.href = '/app/data/data.html';
    }
}