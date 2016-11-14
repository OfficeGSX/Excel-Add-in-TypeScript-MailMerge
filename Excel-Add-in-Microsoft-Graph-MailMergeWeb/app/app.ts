// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the See full license at the root of this repo.

import { Authenticator, TokenManager, DefaultEndpoints, IToken } from '@microsoft/office-js-helpers';

// Main entry point.
export class App {
    authenticator: Authenticator;
    token: IToken;

    constructor() {
        $('.ms-progress-component__footer').hide();
        if (Authenticator.isAuthDialog()) { return; }
        this.authenticator = new Authenticator();
        this.authenticator.tokens.clear();
        this.authenticator.endpoints.registerMicrosoftAuth('your client ID here', {
            scope: 'openid Mail.ReadWrite Mail.Send'
        });
        this.token = this.authenticator.tokens.get(DefaultEndpoints.Microsoft);
    }

    initialize() {
        $('.ms-progress-component__footer').hide();
        if (this.token == null) {
            $('#loginO365PopupButton').click(() => {
                $('.ms-progress-component__footer').show();
                // Start the authentication.
                this.authenticator.authenticate(DefaultEndpoints.Microsoft)
                    .then(token => {
                        this.token = token as IToken;
                        window.location.href = '/app/template/template.html';
                        $('.ms-progress-component__footer').hide();
                    });
            });
        }
        else {
            window.location.href = '/app/template/template.html';
        }
    }
}
