exports.processOptions = _processOptions;
exports.extractPreservableTableStyles = _extractPreservableTableStyles;
exports.convertPreservableStylesToCssString = _convertPreservableStylesToCssString;

var _ = require('underscore');

function _processOptions(options) {
    options = options || {};

    var preservationOptions = options.stylePreservations || {};

    if (preservationOptions === 'all') {
        preservationOptions = {
            useColorSpans: true,
            useFontSizeSpans: true,
            useStrictFontSize: true,
            applyTableStyles: true,
            reduceCellBorderStylesUsed: true,
            ignoreTableElementBorders: true
        };
    }

    if (preservationOptions === 'default') {
        preservationOptions = {
            useColorSpans: true,
            useFontSizeSpans: true,
            useStrictFontSize: false,
            applyTableStyles: true,
            reduceCellBorderStylesUsed: true,
            ignoreTableElementBorders: true
        };
    }

    if (typeof preservationOptions !== 'object')  {
        preservationOptions = {};
    }

    return preservationOptions;
}


// used for both table defaults (<table>) and cells (<td>)
function _extractPreservableTableStyles(elementType, elementProperties, options) {
    if (elementType !== 'table' && elementType !== 'cell') {
        return '';
    }
    options = options || {};

    var fill = elementProperties.firstOrEmpty('w:shd').attributes['w:fill'];
    var cellMargins = elementProperties.firstOrEmpty(elementType === 'cell' ? 'w:tcMar' : 'w:tblCellMar');
    var borders = elementProperties.firstOrEmpty(elementType === 'cell' ? 'w:tcBorders' : 'w:tblBorders');

    var styles = {
        fill: fill && fill !== 'auto' ? fill : null,
        cellMarginTop: _extractCellMarginStyles('top', cellMargins),
        cellMarginLeft: _extractCellMarginStyles('left', cellMargins),
        cellMarginBottom: _extractCellMarginStyles('bottom', cellMargins),
        cellMarginRight: _extractCellMarginStyles('right', cellMargins),
        borderTop: _extractBorderStyles('top', borders),
        borderLeft: _extractBorderStyles('left', borders),
        borderBottom: _extractBorderStyles('bottom', borders),
        borderRight: _extractBorderStyles('right', borders)
    };

    styles = _reduceBorderStyles(styles);

    if (
      styles.fill === null
      && styles.cellMarginTop === null && styles.cellMarginLeft === null && styles.cellMarginBottom === null && styles.cellMarginRight === null
      && styles.borderTop === null && styles.borderLeft === null && styles.borderBottom === null && styles.borderRight === null
    ) {
        styles = null;
    }

    return styles;
}


function _reduceBorderStyles(styles) {
    var directions = ['Top', 'Left', 'Bottom', 'Right'];
    var widthsAndCounts = {};
    var stylesAndCounts = {};
    var colorsAndCounts = {};

    _.each(directions, function(direction) {
        var borderKey = 'border' + direction;

        if (styles[borderKey]) {
            var directionsStyles = styles[borderKey];
            if (directionsStyles.width) {
                _incrementOrStartCount(widthsAndCounts, directionsStyles.width);
            }
            if (directionsStyles.style) {
                _incrementOrStartCount(stylesAndCounts, directionsStyles.style);
            }
            if (directionsStyles.color) {
                _incrementOrStartCount(colorsAndCounts, directionsStyles.color);
            }
        }
    });

    var sortedWidths = _.sortBy(_.values(widthsAndCounts), 'count').reverse();
    var sortedStyles = _.sortBy(_.values(stylesAndCounts), 'count').reverse();
    var sortedColors = _.sortBy(_.values(colorsAndCounts), 'count').reverse();

    styles.simplifiedBorder = {
        width: sortedWidths.length ? sortedWidths[0].val : null,
        style: sortedStyles.length ? sortedStyles[0].val : null,
        color: sortedColors.length ? sortedColors[0].val : null
    };

    styles.simplifiedBorder = (styles.simplifiedBorder.width || styles.simplifiedBorder.style || styles.simplifiedBorder.color) ? styles.simplifiedBorder : null;

    return styles;
}


function _incrementOrStartCount(resultObj, key) {
    if (resultObj[key]) {
        resultObj[key].count += 1;
    } else {
        resultObj[key] = {val: key, count: 1};
    }
}


function _extractBorderStyles(whichBorder, borders) {
    return _extractDirectionalInfo(whichBorder, borders, function(directionalBorder) {
        var borderWidth = directionalBorder.attributes['w:sz'];
        var borderStyle = directionalBorder.attributes['w:val'];
        var borderColor = directionalBorder.attributes['w:color'];

        var borderStyles = {
            width: borderWidth || null,
            style: borderStyle || null,
            color: borderColor || null
        };

        return (borderWidth || borderStyle || borderColor ? borderStyles : null);
    });
}


function _extractCellMarginStyles(whichMargin, margins) {
    return _extractDirectionalInfo(whichMargin, margins, function(directionalMargin) {
        return directionalMargin.attributes['w:w'] || null;
    });
}


function _extractDirectionalInfo(direction, elementProperties, attributeExtractor) {
    if (!direction || !elementProperties || !attributeExtractor || typeof attributeExtractor !== 'function') {
        return null;
    }
    if (direction !== 'top' && direction !== 'left' && direction !== 'bottom' && direction !== 'right') {
        return null;
    }

    var directionInfo = elementProperties.firstOrEmpty('w:' + direction);

    return attributeExtractor(directionInfo);
}


function _convertPreservableStylesToCssString(elementStyles, reduceBorderStyles) {
    var cssString = '';

    // @FUTURE: feature toggle each of these?
    cssString += elementStyles.fill ? ('background-color: #' + elementStyles.fill + ';') : '';

    if (reduceBorderStyles) {
        cssString += elementStyles.simplifiedBorder ? _convertBorderStylesToCssString('', elementStyles.simplifiedBorder) : '';
    } else {
        cssString += elementStyles.borderTop ? _convertBorderStylesToCssString('top', elementStyles.borderTop) : '';
        cssString += elementStyles.borderLeft ? _convertBorderStylesToCssString('left', elementStyles.borderLeft) : '';
        cssString += elementStyles.borderBottom ? _convertBorderStylesToCssString('bottom', elementStyles.borderBottom) : '';
        cssString += elementStyles.borderRight ? _convertBorderStylesToCssString('right', elementStyles.borderRight) : '';
    }

    cssString += elementStyles.cellMarginTop ? 'padding-top: ' + (elementStyles.cellMarginTop / 20) + 'px;' : '';
    cssString += elementStyles.cellMarginLeft ? 'padding-left: ' + (elementStyles.cellMarginLeft / 20) + 'px;' : '';
    cssString += elementStyles.cellMarginBottom ? 'padding-bottom: ' + (elementStyles.cellMarginBottom / 20) + 'px;' : '';
    cssString += elementStyles.cellMarginRight ? 'padding-right: ' + (elementStyles.cellMarginRight / 20) + 'px;' : '';

    return cssString;
}


var  _docxBorderStylesToCssStyles = {
    single: 'solid',
    dashDotStroked: null,
    dashed: 'dashed',
    dashSmallGap: null,
    dotDash: 'dashed',
    dotDotDash: 'dotted',
    dotted: 'dotted',
    double: 'double',
    doubleWave: 'double',
    inset: 'inset',
    nil: 'hidden',
    none: 'none',
    outset: 'outset',
    thick: 'solid',
    thickThinLargeGap: 'double',
    thickThinMediumGap: 'double',
    thickThinSmallGap: 'double',
    thinThickLargeGap: 'double',
    thinThickMediumGap: 'double',
    thinThickSmallGap: 'double',
    thinThickThinLargeGap: 'double',
    thinThickThinMediumGap: 'double',
    thinThickThinSmallGap: 'double',
    threeDEmboss: null,
    threeDEngrave: null,
    triple: null,
    wave: null
};

function _convertBorderStylesToCssString(whichBorder, borderStyles) {
    var css = 'border' + (whichBorder ? '-' + whichBorder : '') + ':';

    css += borderStyles.width ? ' ' + (borderStyles.width / 8) + 'px' : '';  // border widths are stored in eights
    css += borderStyles.style && _docxBorderStylesToCssStyles[borderStyles.style] ? ' ' + _docxBorderStylesToCssStyles[borderStyles.style] : '';
    css += borderStyles.color ? ' #' + borderStyles.color : '';

    css += ';';

    return css;
}
