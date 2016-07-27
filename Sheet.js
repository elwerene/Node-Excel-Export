'use strict';

const fs = require('fs');
const async = require('async');

const sheetFront = '<?xml version="1.0" encoding="utf-8"?><x:worksheet xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"><x:sheetPr><x:outlinePr summaryBelow="1" summaryRight="1" /></x:sheetPr><x:sheetViews><x:sheetView tabSelected="0" workbookViewId="0" /></x:sheetViews><x:sheetFormatPr defaultRowHeight="15" />';
const sheetBack = '</x:sheetData><x:printOptions horizontalCentered="0" verticalCentered="0" headings="0" gridLines="0" /><x:pageMargins left="0.75" right="0.75" top="0.75" bottom="0.5" header="0.5" footer="0.75" /><x:pageSetup paperSize="1" scale="100" pageOrder="downThenOver" orientation="default" blackAndWhite="0" draft="0" cellComments="none" errors="displayed" /><x:headerFooter /><x:tableParts count="0" /></x:worksheet>';


module.exports = (path, config, shareStrings, convertedShareStrings) => {
    return new Promise((resolve, reject) => {
        let cols = config.cols, data = config.rows, colsLength = cols.length, p, styleIndex, k = 0, cn = 1, sheetPos = 0;

        let sheet;

        let write = (str, callback) => sheet.write(str, callback);

        let openSheet = (callback) => {
            sheet = fs.createWriteStream(path, {
                autoclose: false
            });

            sheet.on('open', () => {
                callback();
            });
        };

        let writeInitialSheetContent = (callback) => {
            return write(sheetFront, callback);
        };

        let addColumnHeaders = (callback) => {
            return async.waterfall([
                (callback) => {
                    return write('<cols>', callback);
                },
                (callback) => {
                    return async.eachSeries(cols, (col, callback) => {
                        let colStyleIndex = col.styleIndex || 0;
                        let res = '<x:col min="' + cn + '" max="' + cn + '" width="' + (col.width ? col.width : 10) + '" customWidth="1" style="' + colStyleIndex + '"/>';
                        cn++;
                        return write(res, callback);
                    }, callback);
                },
                (callback) => {
                    return write('</cols><x:sheetData>', callback);
                }
            ], callback);
        };

        let writeRows = (callback) => {
            let addMetadataToColumns = (callback) => {
                return async.eachSeries(cols, (col, callback) => {
                    let colStyleIndex = col.captionStyleIndex || 0;
                    let res = addStringCol(getColumnLetter(k + 1) + 1, col.caption, colStyleIndex, shareStrings);
                    k++;
                    convertedShareStrings += res[1];
                    return write('<x:row r="1" spans="1:' + colsLength + '">' + res[0] + '</x:row>', callback);
                }, callback);
            };

            let subscribeToData = (callback) => {
                data.on('data', addRow);
                data.on('end', callback);
            };

            let i = -1;
            let addRow = (r) => {
                if (r == null) return;

                let j, cellData, currRow, cellType;

                i++;
                currRow = i + 2;
                let row = '<x:row r="' + currRow + '" spans="1:' + colsLength + '">';
                for (j = 0; j < colsLength; j++) {
                    styleIndex = cols[j].styleIndex;
                    cellData = r[j];
                    cellType = cols[j].type;
                    if (typeof cols[j].beforeCellWrite === 'function') {
                        let e = {
                            rowNum: currRow,
                            styleIndex: styleIndex,
                            cellType: cellType
                        };
                        cellData = cols[j].beforeCellWrite(r, cellData, e);
                        styleIndex = e.styleIndex || styleIndex;
                        cellType = e.cellType;
                    }
                    switch (cellType) {
                    case 'number':
                        row += addNumberCol(getColumnLetter(j + 1) + currRow, cellData, styleIndex);
                        break;
                    case 'date':
                        row += addDateCol(getColumnLetter(j + 1) + currRow, cellData, styleIndex);
                        break;
                    case 'bool':
                        row += addBoolCol(getColumnLetter(j + 1) + currRow, cellData, styleIndex);
                        break;
                    default:
                        let res = addStringCol(getColumnLetter(j + 1) + currRow, cellData, styleIndex, shareStrings, convertedShareStrings);
                        row += res[0];
                        convertedShareStrings += res[1];
                    }
                }
                row += '</x:row>';

                return write(row, () => { });
            };

            let addRows = (callback) => {
                data.forEach(addRow);

                return callback();
            };

            if (Array.isArray(data)) {
                return async.waterfall([
                    addMetadataToColumns,
                    addRows
                ], callback);
            } else {
                return async.waterfall([
                    addMetadataToColumns,
                    subscribeToData
                ], callback);
            }
        };

        let writeFinalSheetContent = (callback) => write(sheetBack, callback);

        let closeSheet = (callback) => {
            sheet.on('close', callback);
            sheet.destroySoon();
        };

        async.waterfall([
            openSheet,
            writeInitialSheetContent,
            addColumnHeaders,
            writeRows,
            writeFinalSheetContent,
            closeSheet
        ], () => {
            resolve({
                convertedShareStrings: convertedShareStrings,
                shareStrings: shareStrings
            });
        });
    });
};

let startTag = (obj, tagName, closed) => {
    let result = '<' + tagName, p;
    for (p in obj) {
        result += ' ' + p + '=' + obj[p];
    }
    if (!closed) {
        result += '>';
    } else {
        result += '/>';
    }
    return result;
};
let endTag = (tagName) => {
    return '</' + tagName + '>';
};
let addNumberCol = (cellRef, value, styleIndex) => {
    styleIndex = styleIndex || 0;
    if (value === null) {
        return '';
    } else {
        return '<x:c r="' + cellRef + '" s="' + styleIndex + '" t="n"><x:v>' + value + '</x:v></x:c>';
    }
};
let addDateCol = (cellRef, value, styleIndex) => {
    styleIndex = styleIndex || 1;
    if (value === null) {
        return '';
    } else {
        return '<x:c r="' + cellRef + '" s="' + styleIndex + '" t="n"><x:v>' + value + '</x:v></x:c>';
    }
};
let addBoolCol = (cellRef, value, styleIndex) => {
    styleIndex = styleIndex || 0;
    if (value === null) {
        return '';
    }
    if (value) {
        value = 1;
    } else {
        value = 0;
    }
    return '<x:c r="' + cellRef + '" s="' + styleIndex + '" t="b"><x:v>' + value + '</x:v></x:c>';
};
let addStringCol = (cellRef, value, styleIndex, shareStrings) => {
    styleIndex = styleIndex || 0;
    if (value === null) {
        return [
            '',
            ''
        ];
    }
    if (typeof value === 'string') {
        value = value.replace(/&/g, '&amp;').replace(/'/g, '&apos;').replace(/>/g, '&gt;').replace(/</g, '&lt;');
    }
    let convertedShareStrings = '';
    let i = shareStrings.indexOf(value);
    if (i < 0) {
        i = shareStrings.push(value) - 1;
        convertedShareStrings = '<x:si><x:t>' + value + '</x:t></x:si>';
    }
    return [
        '<x:c r="' + cellRef + '" s="' + styleIndex + '" t="s"><x:v>' + i + '</x:v></x:c>',
        convertedShareStrings
    ];
};
let getColumnLetter = (col) => {
    if (col <= 0) {
        throw 'col must be more than 0';
    }
    let array = [];
    while (col > 0) {
        let remainder = col % 26;
        col /= 26;
        col = Math.floor(col);
        if (remainder === 0) {
            remainder = 26;
            col--;
        }
        array.push(64 + remainder);
    }
    return String.fromCharCode.apply(null, array.reverse());
};
