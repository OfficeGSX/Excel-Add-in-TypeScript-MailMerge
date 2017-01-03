// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
// See full license at the root of this repo.

import * as $ from 'jquery';
import { Authenticator, Utilities } from '@microsoft/office-js-helpers';

class Home {
    constructor() {
        /**
         * The home.js is loaded again inside of the Authentication Dialog as we are redirected
		 * from Microsoft.com after the OAuth process completes.
		 * By using Authenticator.isAuthDialog() we avoid running the code again inside
		 * of the dialog and inform the library to close the dialog instead and report
		 * the success/failure of the OAuth process.
         */
        if (!Authenticator.isAuthDialog()) {
            this._initialize();
        }
    }

    private async _initialize() {
        let authenticator = new Authenticator();
        let token = authenticator.tokens.get('Microsoft');

        /**
         * If we have received a previous token, then go ahead and navigate to the next page
         * else register the Microsoft OAuth endpoint using the helper provided in the library
		 * and call the authenticate function to get the token.
         */
        if (token == null) {
            $('#loginO365PopupButton').click(async () => {
                try {
                    /**
                     * By default registerMicrosoftAuth registers a provider with the name 'Microsoft'
                     */
                    authenticator.endpoints.registerMicrosoftAuth('[Enter your client ID here]', {
                        scope: 'User.Read Mail.ReadWrite Mail.Send',
                        redirectUrl: '[redirect Url]'
                    });
					/**
                     * Using TypeScript 2.1 async/await, we wait to get the token from the
                     * authenticate method and proceed to the next page. Internally the authenticate
                     * method stores a successful token inside of localStorage which we'll reuse later.
                     */
                    await authenticator.authenticate('Microsoft');
                    window.location.replace('../template/index.html');
                }
                catch (e) {
                    /**
                     * The Utilities is a console.error helper that assists in logging errors descriptively
                     * Think of it as a prettier error logger.
                     */
                    Utilities.log(e);
                }
            });
        }
        else {
            window.location.replace('../template/index.html');
        }
    }
};

/**
 * Create a new instance of the Home class and
 * start the application
 */
new Home();
