var fs = require('fs');
var temp = require('temp').track();
var path = require('path');
var JSZip = require('jszip');
var async = require('async');
Date.prototype.getJulian = function () {
    return Math.floor(this / 86400000 - this.getTimezoneOffset() / 1440 + 2440587.5);
};
Date.prototype.oaDate = function () {
    return (this - new Date(Date.UTC(1899, 11, 30))) / (24 * 60 * 60 * 1000);
};
var templateXLSX = new Buffer('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', 'base64');
var sheetFront = '<?xml version="1.0" encoding="utf-8"?><x:worksheet xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"><x:sheetPr><x:outlinePr summaryBelow="1" summaryRight="1" /></x:sheetPr><x:sheetViews><x:sheetView tabSelected="0" workbookViewId="0" /></x:sheetViews><x:sheetFormatPr defaultRowHeight="15" /><cols>';
var sheetBack = '</x:sheetData><x:printOptions horizontalCentered="0" verticalCentered="0" headings="0" gridLines="0" /><x:pageMargins left="0.75" right="0.75" top="0.75" bottom="0.5" header="0.5" footer="0.75" /><x:pageSetup paperSize="1" scale="100" pageOrder="downThenOver" orientation="default" blackAndWhite="0" draft="0" cellComments="none" errors="displayed" /><x:headerFooter /><x:tableParts count="0" /></x:worksheet>';
var sharedStringsFront = '<?xml version="1.0" encoding="UTF-8"?><x:sst xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main" uniqueCount="$count" count="$count">';
var sharedStringsBack = '</x:sst>';
exports.execute = function (config, callback) {
    var cols = config.cols, data = config.rows, colsLength = cols.length, p, files = [], styleIndex, k = 0, cn = 1, dirPath, shareStrings = [], shareStringsObj = {}, convertedShareStrings = '', sheet, sheetPos = 0;
    var write = function (str, callback) {
        var buf = new Buffer(str);
        var off = 0;
        var written = 0;
        return async.whilst(function () {
            return written < buf.length;
        }, function (callback) {
            fs.write(sheet, buf, off, buf.length - off, sheetPos, function (err, w) {
                if (err) {
                    return callback(err);
                }
                written += w;
                off += w;
                sheetPos += w;
                return callback();
            });
        }, callback);
    };
    return async.waterfall([
        function (callback) {
            return temp.mkdir('xlsx', function (err, dir) {
                if (err) {
                    return callback(err);
                }
                dirPath = dir;
                return callback();
            });
        },
        function (callback) {
            return fs.mkdir(path.join(dirPath, 'xl'), callback);
        },
        function (callback) {
            return fs.mkdir(path.join(dirPath, 'xl', 'worksheets'), callback);
        },
        function (callback) {
            return async.parallel([
                function (callback) {
                    return fs.writeFile(path.join(dirPath, 'data.zip'), templateXLSX, callback);
                },
                function (callback) {
                    if (!config.stylesXmlFile) {
                        return callback();
                    }
                    p = config.stylesXmlFile || __dirname + '/styles.xml';
                    return fs.readFile(p, 'utf8', function (err, styles) {
                        if (err) {
                            return callback(err);
                        }
                        p = path.join(dirPath, 'xl', 'styles.xml');
                        files.push(p);
                        return fs.writeFile(p, styles, callback);
                    });
                }
            ], function (err) {
                return callback(err);
            });
        },
        function (callback) {
            p = path.join(dirPath, 'xl', 'worksheets', 'sheet.xml');
            files.push(p);
            return fs.open(p, 'a+', function (err, fd) {
                if (err) {
                    return callback(err);
                }
                sheet = fd;
                return callback();
            });
        },
        function (callback) {
            return write(sheetFront, callback);
        },
        function (callback) {
            return async.eachSeries(cols, function (col, callback) {
                var colStyleIndex = col.styleIndex || 0;
                var res = '<x:col min="' + cn + '" max="' + cn + '" width="' + (col.width ? col.width : 10) + '" customWidth="1" style="' + colStyleIndex + '"/>';
                cn++;
                return write(res, callback);
            }, callback);
        },
        function (callback) {
            return write('</cols><x:sheetData>', callback);
        },
        function (callback) {
            return write('<x:row r="1" spans="1:' + colsLength + '">', callback);
        },
        function (callback) {
            return async.eachSeries(cols, function (col, callback) {
                var colStyleIndex = col.captionStyleIndex || 0;
                var res = addStringCol(getColumnLetter(k + 1) + 1, col.caption, colStyleIndex, shareStrings, shareStringsObj);
                k++;
                if (res[1] !== null) {
                    convertedShareStrings += res[1];
                }
                return write(res[0], callback);
            }, callback);
        },
        function (callback) {
            return write('</x:row>', callback);
        },
        function (callback) {
            var j, r, cellData, currRow, cellType;
            var i = -1;
            data.reverse();
            return async.whilst(function () {
                return data.length > 0;
            }, function (callback) {
                i++;
                r = data.pop();
                currRow = i + 2;
                var row = '<x:row r="' + currRow + '" spans="1:' + colsLength + '">';
                for (j = 0; j < colsLength; j++) {
                    styleIndex = cols[j].styleIndex;
                    cellData = r[j];
                    cellType = cols[j].type;
                    if (typeof cols[j].beforeCellWrite === 'function') {
                        var e = {
                            rowNum: currRow,
                            styleIndex: styleIndex,
                            cellType: cellType
                        };
                        cellData = cols[j].beforeCellWrite(r, cellData, e);
                        styleIndex = e.styleIndex || styleIndex;
                        cellType = e.cellType;
                        e = undefined;
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
                        var res = addStringCol(getColumnLetter(j + 1) + currRow, cellData, styleIndex, shareStrings, shareStringsObj);
                        row += res[0];
                        convertedShareStrings += res[1];
                    }
                }
                row += '</x:row>';
                return write(row, callback);
            }, callback);
        },
        function (callback) {
            return write(sheetBack, callback);
        },
        function (callback) {
            return fs.close(sheet, callback);
        },
        function (callback) {
            if (shareStrings.length === 0) {
                return callback();
            }
            sharedStringsFront = sharedStringsFront.replace(/\$count/g, shareStrings.length);
            p = path.join(dirPath, 'xl', 'sharedStrings.xml');
            files.push(p);
            return fs.writeFile(p, sharedStringsFront + convertedShareStrings + sharedStringsBack, callback);
        }
    ], function (err) {
        if (err) {
            return callback(err);
        }
        var prev = fs.readFileSync(path.join(dirPath, 'data.zip'));
        var zip = new JSZip(prev);
        files.forEach(function (file) {
            var relative = path.relative(dirPath, file);
            zip.file(relative, fs.readFileSync(file));
        });
        var data = zip.generate({
            mimeType: 'application/zip',
            type: 'nodebuffer'
        });
        temp.cleanup();
        return callback(null, data);
    });
};
var startTag = function (obj, tagName, closed) {
    var result = '<' + tagName, p;
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
var endTag = function (tagName) {
    return '</' + tagName + '>';
};
var addNumberCol = function (cellRef, value, styleIndex) {
    styleIndex = styleIndex || 0;
    if (value === null) {
        return '';
    } else {
        return '<x:c r="' + cellRef + '" s="' + styleIndex + '" t="n"><x:v>' + value + '</x:v></x:c>';
    }
};
var addDateCol = function (cellRef, value, styleIndex) {
    styleIndex = styleIndex || 1;
    if (value === null) {
        return '';
    } else {
        return '<x:c r="' + cellRef + '" s="' + styleIndex + '" t="n"><x:v>' + value + '</x:v></x:c>';
    }
};
var addBoolCol = function (cellRef, value, styleIndex) {
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
var addStringCol = function (cellRef, value, styleIndex, shareStrings, shareStringsObj) {
    styleIndex = styleIndex || 0;
    if (value === null) {
        return [
            '',
            null
        ];
    }
    if (typeof value === 'string') {
        value = value.replace(/&/g, '&amp;').replace(/'/g, '&apos;').replace(/>/g, '&gt;').replace(/</g, '&lt;');
    }
    var convertedShareStrings = null;
    var i = shareStringsObj[value];
    if (i === undefined) {
        i = shareStrings.push(value) - 1;
        shareStringsObj[value] = i;
        convertedShareStrings = '<x:si><x:t>' + value + '</x:t></x:si>';
    }
    return [
        '<x:c r="' + cellRef + '" s="' + styleIndex + '" t="s"><x:v>' + i + '</x:v></x:c>',
        convertedShareStrings
    ];
};
var getColumnLetter = function (col) {
    if (col <= 0) {
        throw 'col must be more than 0';
    }
    var array = [];
    while (col > 0) {
        var remainder = col % 26;
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
