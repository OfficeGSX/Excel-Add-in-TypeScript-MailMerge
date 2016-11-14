// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the root of this repo.

require.config({
    paths: {
        'jquery': '/node_modules/jquery/dist/jquery.min',
        'core-js': '/node_modules/core-js/client/core.min',
        'fabric': '/node_modules/office-ui-fabric/dist/js/jquery.fabric.min',
        '@microsoft/office-js-helpers': '/node_modules/@microsoft/office-js-helpers/dist/office.helpers.min',
        'app': './app'
    },
    shim: {
        'fabric': ['jquery'],
        'app': ['fabric', '@microsoft/office-js-helpers'],
        '@microsoft/office-js-helpers': ['core-js']
    }
});

require(['app/data/data', 'jquery', 'core-js', 'fabric', 'app/helpers/dialoghelper'], (Source, $) => {
    $(document).ready(() => {
        Office.initialize = reason => {
            var data = new Source.Data();
            data.initialize();
        }
    });
});