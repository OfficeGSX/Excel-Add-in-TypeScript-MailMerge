// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
// See full license at the root of this repo.

import { Authenticator, DefaultEndpoints, IToken } from '@microsoft/office-js-helpers';

/**
* Helper class to serve methods for ajax request calls.
*/
export class RequestHelper {
    authenticator: Authenticator;
    token: IToken;

    constructor() {
        this.authenticator = new Authenticator();
        // Insert the client ID here.
        this.authenticator.endpoints.registerMicrosoftAuth('[Enter your clientID here]', {
            scope: 'User.Read Mail.ReadWrite Mail.Send',
            redirectUrl: '[redirect Url]'
        });
    }

	/**
	* GET request to the endpoint url.
	*/
    async get<T>(url: string): Promise<T> {
        await this._isAuthenticated();

        let xhr = $.ajax({
            method: 'GET',
            url: url,
            dataType: 'json',
            headers: {
                'Authorization': `Bearer ${this.token.access_token}`
            }
        });

        return this._promise$<T>(xhr);
    }

	/**
	* POST request to the endpoint url.
	*/
    async post<T>(urL: string, data: any): Promise<T> {
        await this._isAuthenticated();

        let xhr = $.ajax({
            method: 'POST',
            url: urL,
            dataType: 'json',
            headers: {
                'Authorization': `Bearer ${this.token.access_token}`,
                'Content-Type': 'application/json'
            },
            data: JSON.stringify(data)
        });

        return this._promise$<T>(xhr);
    }

	/**
	* Promise to handle the ajax request.
	*/
    private _promise$<T>(xhr: JQueryXHR): Promise<T> {
        return new Promise((resolve, reject) => {
            xhr.done(e => resolve(e))
                .fail(e => reject(e));
        });
    }

	/**
	* Sets a token for authentication.
	*/
    private async _isAuthenticated() {
        this.token = await this.authenticator.authenticate('Microsoft');
    }
}

export const request = new RequestHelper();
