// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
// See full license at the root of this repo.

import { Dictionary } from '@microsoft/office-js-helpers';
import { Storage } from '@microsoft/office-js-helpers';

export class ExcelHelper {
    mailMergeData: Dictionary<any>;
    emailAddresses: string[];
    emailAddressesRangeAddress: string;
    firstRowDataContainer: Storage<any>;
    emailAddressesContainer: Storage<any>;
    excelDataContainer: Storage<any>;
    selectedTemplateContainer: Storage<any>;

    constructor() {
        this.mailMergeData = new Dictionary();
        this.firstRowDataContainer = new Storage('FirstRowData');
        this.firstRowDataContainer.clear();
        this.emailAddressesContainer = new Storage('EmailAddresses');
        this.emailAddressesContainer.clear();
        this.excelDataContainer = new Storage('Data');
        this.selectedTemplateContainer = new Storage('SelectedTemplate');
        this.excelDataContainer.clear();
    }

    /**
	* Create a table in Excel that contains all the placeholders from an email template.
	*/
    createMailMergeTable(columnHeaders: string[]) {
        try {
            // Run a batch operation against the Excel object model.
            return Excel.run(async ctx => {

                // Queue a command to get the worksheet collection of existing sheets.
                let worksheets = ctx.workbook.worksheets;

                // Queue a command to load the name property of each worksheet in the collection
                // We will use this later to hide all the existing sheets from view.
                worksheets.load('name');

                // Run the queued-up commands, and return a promise to indicate task completion.
                await ctx.sync()
                // Search for and delete the worksheet named DataSheet.
                for (let i = 0; i < worksheets.items.length; i++) {
                    if (worksheets.items[i].name === 'DataSheet') {
                        worksheets.items[i].delete();
                    }
                }

                await ctx.sync();
                // Queue a command to add a new worksheet to store the transactions.
                let dataSheet = ctx.workbook.worksheets.add('DataSheet');

                // Fill white color in the sheet to remove gridlines.
                dataSheet.getRange().format.fill.color = 'white';

                await ctx.sync();

                // Queue a command to add a new table.
                let lastColumnName = this.columnName(columnHeaders.length - 1);
                let masterTable = ctx.workbook.tables.add('DataSheet!A2:' + lastColumnName + '6', true);
                masterTable.name = 'MailMergeTable';

                let sheet = ctx.workbook.worksheets.getItem('DataSheet');

                // Queue a command to set the header row.
                masterTable.getHeaderRowRange().values = [columnHeaders];

                let selectedTemplate = this.selectedTemplateContainer.get('1');
                let values;
                if (selectedTemplate === 'Absence Limit Exceeded') {
                    // Create an array containing sample data.
                    values = [
                        ['alexd@MOD265542.onmicrosoft.com', 'Alex', 'Janet', 'Mrs. Zrinka'],
                        ['robinc@MOD265542.onmicrosoft.com', 'Robin', 'Molly', 'Mrs. Zrinka'],
                        ['garretv@MOD265542.onmicrosoft.com', 'Garrett', 'Anne', 'Mrs. Zrinka'],
                        ['belindan@MOD265542.onmicrosoft.com', 'Belinda', 'Garth', 'Mrs. Zrinka']];

                } else if (selectedTemplate === 'Progress Report') {
                    values = [
                        ['alexd@MOD265542.onmicrosoft.com', 'Alex', '4', '5', '4', 'Mrs. Zrinka'],
                        ['robinc@MOD265542.onmicrosoft.com', 'Robin', '4', '5', '4', 'Mrs. Zrinka'],
                        ['garretv@MOD265542.onmicrosoft.com', 'Garrett', '4', '5', '4', 'Mrs. Zrinka'],
                        ['belindan@MOD265542.onmicrosoft.com', 'Belinda', '4', '5', '4', 'Mrs. Zrinka']];
                }
                else if (selectedTemplate === 'Parent Teacher Conference') {
                    values = [
                        ['alexd@MOD265542.onmicrosoft.com', 'Alex', 'Janet', 'Mrs. Zrinka'],
                        ['robinc@MOD265542.onmicrosoft.com', 'Robin', 'Molly', 'Mrs. Zrinka'],
                        ['garretv@MOD265542.onmicrosoft.com', 'Garrett', 'Anne', 'Mrs. Zrinka'],
                        ['belindan@MOD265542.onmicrosoft.com', 'Belinda', 'Garth', 'Mrs. Zrinka']];
                }

                // Queue a command to write the sample data to the table.
                masterTable.getDataBodyRange().values = values;

                // Format the table header and data rows.
                this.addContentToWorksheet(sheet, 'A2:' + lastColumnName + '2', '', 'TableHeaderRow');

                // Queue commands to auto-fit columns and rows.
                sheet.getUsedRange().getEntireColumn().format.autofitColumns();
                sheet.getUsedRange().getEntireRow().format.autofitRows();


                // Queue a command to activate the Transactions sheet.
                sheet.activate();


                // Run the queued-up commands, and return a promise to indicate task completion.
                await ctx.sync();
            });
        }
        catch (error) {
            this.handleError(error);
        }
    }

    /**
	* Get the first row data from the table to show a preview.
	*/
    getFirstRowData(): Promise<Dictionary<any>> {
        try {
            return Excel.run(async ctx => {
                // Get the table.
                let mailMergeTable = ctx.workbook.tables.getItem('MailMergeTable');

                // Get the data from the table.
                let headerRowData = mailMergeTable.getHeaderRowRange().load('columnCount, values');
                let firstRow = mailMergeTable.getDataBodyRange().getRow(0).load('values');

                await ctx.sync();

                // Convert values from the 2d array.
                let headerRowDataValueArray = headerRowData.values;
                let firstRowValueArray = firstRow.values;

                let emailAddress = firstRowValueArray[0][0];

                let mergedData = {};
                this.firstRowDataContainer.clear();

                firstRowValueArray[0].forEach((item, index) => {
                    mergedData[headerRowDataValueArray[0][index]] = item;
                });

                this.mailMergeData.insert(emailAddress, mergedData);

                this.firstRowDataContainer.insert(emailAddress, mergedData);
                return this.mailMergeData;
            }) as Promise<Dictionary<any>>;
        }
        catch (error) {
            this.handleError(error);
        }
    }

    /**
	* Get all the email addresses from the table in Excel.
	*/
    getEmailAddresses(): Promise<Dictionary<any>> {
        try {
            return Excel.run(async ctx => {
                // Get the table.
                let mailMergeTable = ctx.workbook.tables.getItem('MailMergeTable');

                // Get the email address column.
                let emailAddressColumn = mailMergeTable.getDataBodyRange().getColumn(0).load('rowCount, values');


                await ctx.sync();
                this.emailAddressesContainer.clear();

                // Convert values from the 2d array.
                let emailAddressColumnValues = emailAddressColumn.values;

                // Store the values of the first column, which contains the email addresses into an array.
                emailAddressColumnValues.forEach((item, index) => {
                    this.emailAddressesContainer.insert(<string><any>index, item[0]);
                });

                return this.emailAddressesContainer;
            }) as Promise<Dictionary<any>>;
        }
        catch (error) {
            this.handleError(error);
        }
    }

    /**
     *  Get merged data from an Excel table.
     */
    getData(): Promise<any> {
        try {
            return Excel.run(async ctx => {
                // Get the table.
                let mailMergeTable = ctx.workbook.tables.getItem('MailMergeTable');

                let headerRowData = mailMergeTable.getHeaderRowRange().load('columnCount, values');
                let dataRows = mailMergeTable.getDataBodyRange().load('rowCount, values');

                await ctx.sync();
                // Get the values.
                let headerRowDataValueArray = headerRowData.values;
                let dataRowsValueArray = dataRows.values;

                this.excelDataContainer.clear();

                dataRowsValueArray.forEach((item, index) => {
                    let mergedData = {};
                    item.forEach((item, index) => {
                        mergedData[headerRowDataValueArray[0][index]] = item;
                    });
                    this.excelDataContainer.insert(item[0], mergedData);
                });

                return this.excelDataContainer;
            }) as Promise<any>;
        }
        catch (error) {
            this.handleError(error);
        }
    }

    /**
    * Returns the column name based on a zero-based column index.
    * For example, columnName(4) = 5th column = "E". Meanwhile, columnName(1000) = 1001st column = "ALM".
    * @param index Zero-based column index.
    * @returns {String} Locale-independent column name (e.g., a string comprised of one or more letters in the range "A:Z").
    */
    columnName(index) {
        if (typeof index !== 'number' || isNaN(index) || index < 0) {
            // Throw exception here.
            throw 'Error: Parameter is not a number.';
        }
        let letters = [];
        while (index >= 0) {
            letters.push(getSingleLetter(index % 26));
            index = Math.floor(index / 26) - 1;
        }
        return letters.reverse().join('');
        function getSingleLetter(zeroThrough25Index) {
            return String.fromCharCode(zeroThrough25Index + 65); // ASCII code for "A" is 65
        }
    }

    // Handle errors.
    handleError(error) {
        // Always be sure to catch any accumulated errors that bubble up from the Excel.run execution.
        console.log('Error: ' + error);
    }

    /**
     *  Helper function to add and format content in the worksheet.
     */
    addContentToWorksheet(sheetObject, rangeAddress, displayText, typeOfText) {
        let range;

        // Format differently by the type of content.
        switch (typeOfText) {
            case 'TableHeading':
                range = sheetObject.getRange(rangeAddress);
                range.values = displayText;
                range.format.font.name = 'Segoe UI';
                range.format.font.size = 12;
                range.format.font.color = '#00b3b3';
                range.merge();
                break;
            case 'TableHeaderRow':
                range = sheetObject.getRange(rangeAddress);
                range.format.font.name = 'Segoe UI';
                range.format.font.size = 10;
                range.format.font.bold = true;
                range.format.font.color = 'black';
                break;
            default:
                break;
        }
    }
}
