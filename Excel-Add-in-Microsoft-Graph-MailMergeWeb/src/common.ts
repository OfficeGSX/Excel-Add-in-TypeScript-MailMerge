/*Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in root of repo. */

/***
* Initializes the add-in by providing user log in, registering the app, and getting a token.
***/

require.config({
    paths: {
        'jquery': '/node_modules/jquery/dist/jquery.min',
        'core-js': '/node_modules/core-js/client/core.min',
        'fabric': '/node_modules/office-ui-fabric/dist/js/jquery.fabric.min',
        '@microsoft/office-js-helpers': '/node_modules/@microsoft/office-js-helpers/dist/office.helpers'
    },
    shim: {
        'fabric': ['jquery'],
        '@microsoft/office-js-helpers': ['core-js']
    }
});