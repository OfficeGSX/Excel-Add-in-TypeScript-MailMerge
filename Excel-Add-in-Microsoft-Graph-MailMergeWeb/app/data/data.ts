// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.

import { ExcelHelper } from '../helpers/index';
import {OutlookHelper} from '../helpers/index';
import {Storage} from '@microsoft/office-js-helpers';
import {Authenticator, DefaultEndpoints, IToken} from '@microsoft/office-js-helpers';

// The Data class serves methods used in the data page.
export class Data {
    authenticator: Authenticator;
    excelHelper: ExcelHelper;
    outlookHelper: OutlookHelper;
    columnHeaders: any;
    selectedTemplate: string;
    templatesContainer: Storage<any>;
    selectedTemplateContainer: Storage<any>;


    constructor() {
        this.authenticator = new Authenticator();
        this.excelHelper = new ExcelHelper();
        this.outlookHelper = new OutlookHelper();
        this.templatesContainer = new Storage('Templates');
        this.selectedTemplateContainer = new Storage('SelectedTemplate');

        this.selectedTemplate = this.selectedTemplateContainer.get('1');
        this.columnHeaders = this.templatesContainer.get(this.selectedTemplate);
        this.columnHeaders.splice(0, 0, "EmailAddress");

        console.log('initialized');
        $('#previewEmail').click(() => {
            // Get the first row of data from the Excel table.
            this.excelHelper.getFirstRowData()
                // Get the user profile details.
                .then(() => this.outlookHelper.getUserProfile())
                // Display the preview dialog with the first user's information.
                .then(() => (window as any).app.openDialog(location.origin + "/app/data/preview.html", 35, 25));
        });

        $('#sendEmail').click(() => {
            $("#spinner").show();
            $("#mainpage").hide();
            // Get all of the email addresses from Excel.
            this.excelHelper.getEmailAddresses()
                // Get all of the mail merge data from Excel.
                .then(() => this.excelHelper.getData())
                .then(() => this.outlookHelper.sendMessages())
                .then(() => {
                    window.location.href = 'sendstatus.html';
                    console.log('emails sent');
                })
                .catch(() => { console.log('atleast 1 email failed'); });
        });
        $('#openHelpDialog').click(evt => (window as any).app.openDialog(location.origin + "/app/help/help.html", 65, 65));
        $('#signOut').click(() => this.signOut());

    }

    initialize() {
        this.excelHelper.createMailMergeTable(this.columnHeaders);
    }

    // Sign the user out.
    signOut() {
        this.authenticator.tokens.clear();
        window.location.href = '/index.html';
    }
}