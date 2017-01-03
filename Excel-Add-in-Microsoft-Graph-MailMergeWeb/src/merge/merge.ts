// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. 
// See full license at the bottom of this file.

import * as $ from 'jquery';
import 'fabric';
import { Authenticator, Storage, IToken, Dialog, Utilities } from '@microsoft/office-js-helpers';
import { ExcelHelper, OutlookHelper } from '../helpers/index';
import { showHelp } from '../help/help';

/**
 * The Merge class serves methods to preview and send mail.
 */
class Merge {
    authenticator: Authenticator;
    token: IToken;
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
        this._initialize();
    }

    private async _initialize() {
        this._registerHelp();
        this._registerSignOut();
        this._registerPreview();
        this._registerSend();
        this._loadData();
    }

	/**
	* Sets data for the message templates.
	*/
    private _loadData() {
        this.selectedTemplate = this.selectedTemplateContainer.get('1');
        this.columnHeaders = this.templatesContainer.get(this.selectedTemplate);
        this.columnHeaders.splice(0, 0, 'EmailAddress');
        this.excelHelper.createMailMergeTable(this.columnHeaders);
    }

	/**
	* Handles the click event for the help dialog window.
	*/
    private _registerHelp() {
        $('#openHelpDialog').click(async evt => showHelp());
    }

	/**
	* Opens a dialog window that shows a mail preview with data from the first 
	* row of the table in the Excel spreadsheet.
	*/
    private _registerPreview() {
        $('#previewEmail')
            .click(async () => {
                await Promise.all([
				    // Get data from the first row in the table.
                    this.excelHelper.getFirstRowData(),
					// Get the user profile.
                    this.outlookHelper.getUserProfile()
                ]);

                await new Dialog<string>(`${location.origin}/dist/preview/index.html`, 1024, 768).result;
            });
    }

	/**
	* Attaches a handler to send mail.
	*/
    private _registerSend() {
        $('#sendEmail').click(async () => {
            $('#spinner').show();
            $('#mainpage').hide();

            try {
                // Get all of the email addresses from Excel.
                await this.excelHelper.getEmailAddresses();

                // Get all of the mail merge data from Excel.
                await this.excelHelper.getData();

				// Send the messages with the data loaded in the template.
                await this.outlookHelper.sendMessages();
                window.location.href = '../sent/index.html';
            }
            catch (error) {
                Utilities.log(error);
                window.location.href = '../home/index.html';
            };
        });
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
}

/**
 * Create a new instance of the Merge class and
 * load the data.
 */
new Merge();
