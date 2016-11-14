// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the root of this repo.

import { Dictionary } from '@microsoft/office-js-helpers';
import {Storage} from '@microsoft/office-js-helpers';

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

    createMailMergeTable(columnHeaders: string[]) {

        // Run a batch operation against the Excel object model.
        return Excel.run(ctx => {

            // Queue a command to add a new worksheet to store the transactions.
            var dataSheet = ctx.workbook.worksheets.add("DataSheet");

            // Fill white color in the sheet to remove gridlines.
            dataSheet.getRange().format.fill.color = "white";

            return ctx.sync()
                .then(() => {

                    // Queue a command to add a new table.
                    var lastColumnName = this.columnName(columnHeaders.length - 1);
                    var masterTable = ctx.workbook.tables.add('DataSheet!A2:' + lastColumnName + '6', true);
                    masterTable.name = "MailMergeTable";

                    // Queue a command to set the header row.
                    masterTable.getHeaderRowRange().values = [columnHeaders];

                    var selectedTemplate = this.selectedTemplateContainer.get('1');
                    if (selectedTemplate == "Absence Limit Exceeded") {
                        // Create an array containing sample data.
                        var values = [
                            ["alexd@MOD265542.onmicrosoft.com", "Alex", "Janet", "Mrs. Zrinka"],
                            ["robinc@MOD265542.onmicrosoft.com", "Robin", "Molly", "Mrs. Zrinka"],
                            ["garretv@MOD265542.onmicrosoft.com", "Garrett", "Anne", "Mrs. Zrinka"],
                            ["belindan@MOD265542.onmicrosoft.com", "Belinda", "Garth", "Mrs. Zrinka"]];

                        // Queue a command to write the sample data to the table.
                        masterTable.getDataBodyRange().values = values;
                    }

                    // Format the table header and data rows.
                    this.addContentToWorksheet(dataSheet, 'A2:' + lastColumnName + '2', "", "TableHeaderRow");

                    // Queue commands to auto-fit columns and rows.
                    dataSheet.getUsedRange().getEntireColumn().format.autofitColumns();
                    dataSheet.getUsedRange().getEntireRow().format.autofitRows();


                    // Queue a command to activate the Transactions sheet.
                    dataSheet.activate();


                    // Run the queued-up commands, and return a promise to indicate task completion.
                    return ctx.sync();
                });
        })
            .catch(error => this.handleError(error));
    }


    getFirstRowData(): Promise<Dictionary<any>> {
        return Excel.run(ctx => {
            // Get the table.
            var mailMergeTable = ctx.workbook.tables.getItem("MailMergeTable");

            // Get the data from the table.
            var headerRowData = mailMergeTable.getHeaderRowRange().load("columnCount, values");
            var firstRow = mailMergeTable.getDataBodyRange().getRow(0).load("values");


            return ctx.sync()
                .then(() => {

                    // Convert values from the 2d array.
                    var headerRowDataValueArray = headerRowData.values;
                    var firstRowValueArray = firstRow.values;

                    var emailAddress = firstRowValueArray[0][0];

                    var mergedData = {};
                    this.firstRowDataContainer.clear();

                    for (var i = 0; i < headerRowData.columnCount; i++) {
                        mergedData[headerRowDataValueArray[0][i]] = firstRowValueArray[0][i];
                    }
                    console.log("merged data" + mergedData);

                    this.mailMergeData.add(emailAddress, mergedData);
                    console.log(this.mailMergeData);

                    this.firstRowDataContainer.insert(emailAddress, mergedData);
                    return this.mailMergeData;
                });
        }) as Promise<Dictionary<any>>;
    }


    getEmailAddresses(): Promise<Dictionary<any>> {
        return Excel.run(ctx => {
            // Get the table.
            var mailMergeTable = ctx.workbook.tables.getItem("MailMergeTable");

            // Get the email address column.
            var emailAddressColumn = mailMergeTable.getDataBodyRange().getColumn(0).load("rowCount, values");


            return ctx.sync()
                .then(() => {
                    this.emailAddressesContainer.clear();

                    // Convert values from the 2d array.
                    var emailAddressColumnValues = emailAddressColumn.values;

                    for (var i = 0; i < emailAddressColumn.rowCount; i++) {
                        this.emailAddressesContainer.insert(i.toString(), emailAddressColumnValues[i][0]);
                        console.log(i + " " + emailAddressColumnValues[i][0]);
                    }
                    return this.emailAddressesContainer;
                });
        }) as Promise<Dictionary<any>>;

    }

    getData(): Promise<any> {
        return Excel.run(ctx => {
            // Get the table.
            var mailMergeTable = ctx.workbook.tables.getItem("MailMergeTable");

            var headerRowData = mailMergeTable.getHeaderRowRange().load("columnCount, values");
            var dataRows = mailMergeTable.getDataBodyRange().load("rowCount, values");


            return ctx.sync()
                .then(() => {
                    // Get the values.
                    var headerRowDataValueArray = headerRowData.values;
                    var dataRowsValueArray = dataRows.values;

                    this.excelDataContainer.clear();

                    for (var i = 0; i < dataRows.rowCount; i++) {
                        var emailAddress = dataRowsValueArray[i][0];

                        var mergedData = {};

                        for (var j = 0; j < headerRowData.columnCount; j++) {
                            mergedData[headerRowDataValueArray[0][j]] = dataRowsValueArray[i][j];
                        }

                        this.excelDataContainer.insert(emailAddress, mergedData);
                    }

                    return this.excelDataContainer;
                });
        }) as Promise<any>;

    }

  /**
  * Returns the column name based on a zero-based column index.
  * For example, columnName(4) = 5th column = "E". Meanwhile, columnName(1000) = 1001st column = "ALM".
  * @param index Zero-based column index.
  * @returns {String} Locale-independent column name (e.g., a string comprised of one or more letters in the range "A:Z").
  */
    columnName(index) {
        if (typeof index !== 'number' || isNaN(index) || index < 0) {
            // throw exception here. 
        }
        var letters = [];
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
        console.log("Error: " + error);        
    }

    // Helper function to add and format content in the workbook.
    addContentToWorksheet(sheetObject, rangeAddress, displayText, typeOfText) {

        // Format differently by the type of content.
        switch (typeOfText) {
            case "TableHeading":
                var range = sheetObject.getRange(rangeAddress);
                range.values = displayText;
                range.format.font.name = "Corbel";
                range.format.font.size = 12;
                range.format.font.color = "#00b3b3";
                range.merge();
                break;
            case "TableHeaderRow":
                var range = sheetObject.getRange(rangeAddress);
                range.format.font.name = "Corbel";
                range.format.font.size = 10;
                range.format.font.bold = true;
                range.format.font.color = "black";
                break;
            default:
                break;
        }
    }

}