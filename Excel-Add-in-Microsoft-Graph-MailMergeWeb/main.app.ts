// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the root of this repo.

require.config({
    paths: {
        'jquery': '/node_modules/jquery/dist/jquery.min',
        'core-js': '/node_modules/core-js/client/core.min',
        'fabric': '/node_modules/office-ui-fabric/dist/js/jquery.fabric.min',
        'app': './app',
        '@microsoft/office-js-helpers': '/node_modules/@microsoft/office-js-helpers/dist/office.helpers.min'
    },
    shim: {
        'fabric': ['jquery'],
        'app': ['fabric', '@microsoft/office-js-helpers'],
        '@microsoft/office-js-helpers':['core-js']
    }
});

require(['app/app', 'jquery', 'core-js', 'fabric'], (Source, $) => {
    $(document).ready(() => {
        Office.initialize = reason => {
            var app = new Source.App();
            app.initialize();
        }
    });
});