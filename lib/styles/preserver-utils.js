exports.processOptions = _processOptions;
exports.extractPreservableTableStyles = _extractPreservableTableStyles;
exports.convertPreservableStylesToCssString = _convertPreservableStylesToCssString;

function _processOptions(options) {
    options = options || {};

    var preservationOptions = options.stylePreservations || {};

    if (preservationOptions === 'all') {
        preservationOptions = {
            useColorSpans: true,
            useFontSizeSpans: true,
            useStrictFontSize: true,
            applyTableCellStyles: true
        };
    }

    if (preservationOptions === 'default') {
        preservationOptions = {
            useColorSpans: true,
            useFontSizeSpans: true,
            useStrictFontSize: false,
            applyTableCellStyles: true
        };
    }

    if (typeof preservationOptions !== 'object')  {
        preservationOptions = {};
    }

    return preservationOptions;
}


// used for both table defaults (<table>) and cells (<td>)
function _extractPreservableTableStyles(elementType, elementProperties) {
    if (elementType !== 'table' && elementType !== 'cell') {
        return '';
    }

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

    if (
      styles.fill === null
      && styles.cellMarginTop === null && styles.cellMarginLeft === null && styles.cellMarginBottom === null && styles.cellMarginRight === null
      && styles.borderTop === null && styles.borderLeft === null && styles.borderBottom === null && styles.borderRight === null
    ) {
        styles = null;
    }

    return styles;
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


function _convertPreservableStylesToCssString(elementStyles) {
    var cssString = '';

    // @FUTURE: feature toggle each of these?
    cssString += elementStyles.fill ? ('background-color: #' + elementStyles.fill + ';') : '';

    cssString += elementStyles.borderTop ? _convertBorderStylesToCssString('top', elementStyles.borderTop) : '';
    cssString += elementStyles.borderLeft ? _convertBorderStylesToCssString('left', elementStyles.borderLeft) : '';
    cssString += elementStyles.borderBottom ? _convertBorderStylesToCssString('bottom', elementStyles.borderBottom) : '';
    cssString += elementStyles.borderRight ? _convertBorderStylesToCssString('right', elementStyles.borderRight) : '';

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
    if (whichBorder !== 'top' && whichBorder !== 'left' && whichBorder !== 'bottom' && whichBorder !== 'right')  {
        return '';
    }

    var css = 'border-' + whichBorder + ':';

    css += borderStyles.width ? ' ' + (borderStyles.width / 8) + 'px' : '';  // border widths are stored in eights
    css += borderStyles.style && _docxBorderStylesToCssStyles[borderStyles.style] ? ' ' + _docxBorderStylesToCssStyles[borderStyles.style] : '';
    css += borderStyles.color ? ' #' + borderStyles.color : '';

    css += ';';

    return css;
}
