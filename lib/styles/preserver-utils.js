exports.processOptions = processOptions;

function processOptions(options) {
    options = options || {};

    var preservationOptions = options.stylePreservations || {};

    if (preservationOptions === 'all') {
        preservationOptions = {
            useColorSpans: true,
            useFontSizeSpans: true,
            useStrictFontSize: true
        };
    }
    
    if (preservationOptions === 'default') {
        preservationOptions = {
            useColorSpans: true,
            useFontSizeSpans: true,
            useStrictFontSize: false
        };
    }

    if (typeof preservationOptions !== 'object')  {
        preservationOptions = {};
    }

    return preservationOptions;
}
