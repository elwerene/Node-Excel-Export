'use strict';
const fs = require('fs');
const temp = require('temp').track();
const path = require('path');
const JSZip = require('jszip');
const async = require('async');
const Sheet = require('./Sheet');

Date.prototype.getJulian = function () {
    return Math.floor(this / 86400000 - this.getTimezoneOffset() / 1440 + 2440587.5);
};
Date.prototype.oaDate = function () {
    return (this - new Date(Date.UTC(1899, 11, 30))) / (24 * 60 * 60 * 1000);
};

const templateXLSX = new Buffer('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', 'base64');

const sheetsFront = '<?xml version="1.0" encoding="utf-8"?><x:workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><x:workbookPr codeName="ThisWorkbook" /><x:bookViews><x:workbookView firstSheet="0" activeTab="0" /></x:bookViews><x:sheets>';
const sheetsBack = '</x:sheets><x:definedNames /><x:calcPr calcId="125725" /></x:workbook>';


let sharedStringsFront = '<?xml version="1.0" encoding="UTF-8"?><x:sst xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main" uniqueCount="$count" count="$count">';
const sharedStringsBack = '</x:sst>';

const relFront = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
const relBack = '<Relationship Id="rId101" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/><Relationship Id="rId100" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId102" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme.xml"/></Relationships>';

exports.execute = (configs, callback) => {
    let p, files = [], styleIndex, k = 0, cn = 1, dirPath, shareStrings = [], convertedShareStrings = '', sheet, sheetPos = 0;

    if (!Array.isArray(configs)) {
        configs = [configs, configs];
    }

    let makeTemporaryFolder = (callback) => {
        return temp.mkdir('xlsx', (err, dir) => {
            if (err) {
                return callback(err);
            }
            dirPath = dir;
            return callback();
        });
    };

    let makeTemporaryFolderStructure = (callback) => {
        async.waterfall([
            (callback) => {
                return fs.mkdir(path.join(dirPath, 'xl'), callback);
            },
            (callback) => {
                return fs.mkdir(path.join(dirPath, 'xl', 'worksheets'), callback)
            },
            (callback) => {
                return fs.mkdir(path.join(dirPath, 'xl', '_rels'), callback)
            },
            (callback) => {
                return fs.mkdir(path.join(dirPath, '_rels'), callback)
            },
            (callback) => {
                return async.parallel([
                    (callback) => {
                        return fs.writeFile(path.join(dirPath, 'data.zip'), templateXLSX, callback);
                    },
                    (callback) => {
                        let styleSheet = configs.find((config) => config.stylesXmlFile != null);
                        if (styleSheet != null) {
                            return fs.readFile(styleSheet, 'utf8', (err, styles) => {
                                if (err) {
                                    return callback(err);
                                }
                                let p = path.join(dirPath, 'xl', 'styles.xml');
                                files.push(p);
                                return fs.writeFile(p, styles, callback);
                            });

                        } else {
                            callback();
                        }
                    }
                ], (err) => {
                    return callback(err);
                });
            }
        ], callback);
    };

    let generateRel = (callback) => {
        let workbook = relFront;
    	let i = 1;
    	configs.forEach( function(config) {
    		workbook += '<Relationship Id="rId' + (i + 1) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' + i + '.xml"/>';
    		i++;
    	});
    	workbook += relBack;

        return async.parallel([
            (callback) => {
                let p = path.join(dirPath, 'xl', '_rels', 'workbook.xml.rels');
                files.push(p);
                return fs.writeFile(p, workbook, callback);
            },
            (callback) => {
                let p = path.join(dirPath, '_rels', '.rels');
                files.push(p);
                return fs.writeFile(p, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            			  + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            			  + '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>'
            			  + '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>', callback);
            }
        ], () => {
            callback();
        });

    };

    let generateWorkbook = (callback) => {
        let workbook = sheetsFront;
        configs.forEach( (config, i) => {
    		workbook += '<x:sheet name="Sheet ' + (i + 1) + '" sheetId="' + (i + 2) +'" r:id="rId' + (i + 2) + '"/>';
    	});
        workbook += sheetsBack;
        let p = path.join(dirPath, 'xl', 'workbook.xml');
        files.push(p);
    	return fs.writeFile(p, workbook, callback);
    }

    let writeSharedString = (callback) => {
        if (shareStrings.length === 0) {
            return callback();
        }

        sharedStringsFront = sharedStringsFront.replace(/\$count/g, shareStrings.length);
        p = path.join(dirPath, 'xl', 'sharedStrings.xml');
        files.push(p);
        return fs.writeFile(p, sharedStringsFront + convertedShareStrings + sharedStringsBack, callback);
    };

    let zipFile = (err) => {
        if (err) {
            return callback(err);
        }

        fs.readFile(path.join(dirPath, 'data.zip'), (err, prev) => {
            let zip = new JSZip();

            files.forEach((file) => {
                let relative = path.relative(dirPath, file);
                zip.file(relative, fs.createReadStream(file));
            });

            zip.loadAsync(prev)
                .then((zip) => {
                    zip
                        .generateNodeStream({streamFiles:true})
                        .pipe(fs.createWriteStream(dirPath + '.xlsx'))
                        .on('finish', () => {
                            temp.cleanup();
                            return callback(null, dirPath + '.xlsx')
                        });
                });
        });
    };

    let finalizeZip = () => {
        return async.waterfall([
            writeSharedString
        ], () => {
            zipFile();
        });
    };

    let writeSheets = (callback) => {
        async.eachOfSeries(configs, (config, i, callback) => {
            let p = path.join(dirPath, 'xl', 'worksheets', 'sheet' + (i + 1) + '.xml');
            files.push(p);
            Sheet(p, config, shareStrings, convertedShareStrings)
                .then((ammendedSharedStrings) => {
                    convertedShareStrings = ammendedSharedStrings.convertedShareStrings;
                    shareStrings = ammendedSharedStrings.shareStrings;
                    callback();
                });
        }, callback);
    };

    async.waterfall([
        makeTemporaryFolder,
        makeTemporaryFolderStructure,
        generateRel,
        writeSheets,
        generateWorkbook], (err) => {
            finalizeZip();
        });

};
