var async=require('async');
var fs=require('fs');
var temp=require('temp').track();
var path=require('path');
var zipper=require('adm-zip');
var util=require('util');
var child_process=require('child_process');
var template=require('./resources/template');

Date.prototype.getJulian=function() {
	return Math.floor((this/86400000)-
		(this.getTimezoneOffset()/1440)+2440587.5);
};

Date.prototype.oaDate=function() {
	return (this-new Date(Date.UTC(1899, 11, 30)))/(24*60*60*1000);
};

/**
 * build xlsx buffer or file depending on whether target is specified or not
 * @param config per row/column descriptions
 * @param target optional target path/file
 * @param callback function(error, buffer/file)
 */
exports.execute=function(config, target, callback) {
	var cols=config.cols,
		data=config.rows,
		colsLength=cols.length,
		p,
		styleIndex,
		k=0,
		cn=1,
		dirPath,
		shareStrings=[],
		convertedShareStrings="",
		sheet,
		sheetPos=0;

	if(typeof(target)==="function") {
		callback=target;
		target=null;
	}

	var write=function(str, callback) {
		var buf=new Buffer(str);
		var off=0;
		var written=0;

		async.whilst(
			function() {
				return written<buf.length;
			},
			function(callback) {
				fs.write(sheet, buf, off, buf.length-off, sheetPos, function(err, w) {
					if(err) {
						return callback(err);
					}

					written+=w;
					off+=w;
					sheetPos+=w;

					return callback();
				});
			},
			callback
		);
	};

	async.waterfall([
			function(callback) {
				temp.mkdir('xlsx', function(err, dir) {
					dirPath=dir;
					callback(err);
				});
			},
			function(callback) {
				// take our XLXS bare bones template and expand it to the file system. We will amend it.
				try {
					var zip=new zipper(template.XLSX);
					zip.extractAllTo(dirPath);
					callback();
				} catch(error) {
					callback("extraction failed: " + error);
				}
			},
			function(callback) {
				p=config.stylesXmlFile || __dirname+'/resources/styles.xml';
				fs.readFile(p, 'utf8', function(err, styles) {
					if(err) {
						return callback(err);
					}
					p=path.join(dirPath, 'xl', 'styles.xml');
					fs.writeFile(p, styles, callback);
				});
			},
			function(callback) {
				p=path.join(dirPath, 'xl', 'worksheets', 'sheet.xml');
				fs.open(p, 'a+', function(err, fd) {
					sheet=fd;
					callback(err);
				});
			},
			function(callback) {
				write(template.sheetFront, callback);
			},
			function(callback) {
				async.eachSeries(cols, function(col, _callback) {
					var colStyleIndex=col.styleIndex || 0;
					var res=util.format('<x:col min="%d" max="%d" width="%d" customWidth="1" style="%d"/>',
						cn, cn, (col.width ? col.width : 10), colStyleIndex);
					cn++;
					write(res, _callback);
				}, callback);
			},
			function(callback) {
				write('</cols><x:sheetData>', callback);
			},
			function(callback) {
				write('<x:row r="1" spans="1:'+colsLength+'">', callback);
			},
			function(callback) {
				async.eachSeries(cols, function(col, _callback) {
					var colStyleIndex=col.captionStyleIndex || 0;
					var res=addStringCol(getColumnLetter(k+1)+1, col.caption, colStyleIndex, shareStrings);
					k++;
					convertedShareStrings+=res[1];
					write(res[0], _callback);
				}, callback);
			},
			function(callback) {
				write('</x:row>', callback);
			},
			function(callback) {
				var j, r, cellData, currRow, cellType;
				var i=-1;

				data.reverse();

				async.whilst(
					function() {
						return data.length>0;
					},
					function(_callback) {
						i++;
						r=data.pop();
						currRow=i+2;

						var row='<x:row r="'+currRow+'" spans="1:'+colsLength+'">';

						for(j=0; j<colsLength; j++) {
							styleIndex=cols[j].styleIndex;
							cellData=r[j];
							cellType=cols[j].type;

							if(typeof cols[j].beforeCellWrite==='function') {
								var e={
									rowNum: currRow,
									styleIndex: styleIndex,
									cellType: cellType
								};

								cellData=cols[j].beforeCellWrite(r, cellData, e);
								styleIndex=e.styleIndex || styleIndex;
								cellType=e.cellType;
								e=undefined;
							}

							switch(cellType) {
								case 'number':
									row+=addNumberCol(getColumnLetter(j+1)+currRow, cellData, styleIndex);
									break;
								case 'date':
									row+=addDateCol(getColumnLetter(j+1)+currRow, cellData, styleIndex);
									break;
								case 'bool':
									row+=addBoolCol(getColumnLetter(j+1)+currRow, cellData, styleIndex);
									break;
								default:
									var res=addStringCol(getColumnLetter(j+1)+currRow, cellData, styleIndex, shareStrings, convertedShareStrings);
									row+=res[0];
									convertedShareStrings+=res[1];
							}
						}

						row+='</x:row>';

						write(row, _callback);
					}, callback);
			},
			function(callback) {
				write(template.sheetBack, callback);
			},
			function(callback) {
				fs.close(sheet, callback);
			},
			function(callback) {
				if(shareStrings.length===0) {
					return callback();
				}
				var sharedStringsFront=template.sharedStringsFront.replace(/\$count/g, shareStrings.length);
				p=path.join(dirPath, 'xl', 'sharedStrings.xml');
				fs.writeFile(p, sharedStringsFront+convertedShareStrings+template.sharedStringsBack, callback);
			},
			function(callback) {
				var _target=(target)
					? path.resolve(target)
					: path.join(dirPath, "sheet.zip");
				// I'm having trouble with adm-zip's zip encoding. Not sure what the problem is exactly
				// but spending more time on this problem than I meant to - going native.
				async.waterfall([
					function(done) {
						child_process.exec(util.format('zip -r "%s" .', _target), {cwd: dirPath}, function(error) {
							done(error);
						});
					},
					function(done) {
						if(target) {
							callback(null, target);
						} else {
							fs.readFile(_target, done);
						}
					}
				], callback);
			}
		],
		function(err, result) {
			temp.cleanup();
			callback(err, result);
		}
	);
};

var addNumberCol=function(cellRef, value, styleIndex) {
	styleIndex=styleIndex || 0;
	if(value===null) {
		return "";
	} else {
		return '<x:c r="'+cellRef+'" s="'+styleIndex+'" t="n"><x:v>'+value+'</x:v></x:c>';
	}
};

var addDateCol=function(cellRef, value, styleIndex) {
	styleIndex=styleIndex || 1;
	if(value===null) {
		return "";
	} else {
		return '<x:c r="'+cellRef+'" s="'+styleIndex+'" t="n"><x:v>'+value+'</x:v></x:c>';
	}
};

var addBoolCol=function(cellRef, value, styleIndex) {
	styleIndex=styleIndex || 0;
	if(value===null) {
		return "";
	}

	if(value) {
		value=1;
	} else {
		value=0;
	}

	return '<x:c r="'+cellRef+'" s="'+styleIndex+'" t="b"><x:v>'+value+'</x:v></x:c>';
};

var addStringCol=function(cellRef, value, styleIndex, shareStrings) {
	styleIndex=styleIndex || 0;
	if(value===null) {
		return ["", ""];
	}

	if(typeof value==='string') {
		value=value.replace(/&/g, "&amp;").replace(/'/g, "&apos;").replace(/>/g, "&gt;").replace(/</g, "&lt;");
	}

	var convertedShareStrings="";
	var i=shareStrings.indexOf(value);
	if(i<0) {
		i=shareStrings.push(value)-1;
		convertedShareStrings="<x:si><x:t>"+value+"</x:t></x:si>";
	}

	return ['<x:c r="'+cellRef+'" s="'+styleIndex+'" t="s"><x:v>'+i+'</x:v></x:c>', convertedShareStrings];
};

var getColumnLetter=function(col) {
	if(col<=0) {
		throw "col must be more than 0";
	}

	var array=[];
	while(col>0) {
		var remainder=col%26;
		col/=26;
		col=Math.floor(col);

		if(remainder===0) {
			remainder=26;
			col--;
		}

		array.push(64+remainder);
	}

	return String.fromCharCode.apply(null, array.reverse());
};
