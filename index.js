var fs=require('fs');
var temp=require('temp').track();
var path=require('path');
var zipper=require('adm-zip');
var async=require('async');
var util=require('util');
var template=require('./resources/template');

Date.prototype.getJulian=function() {
	return Math.floor((this/86400000)-
		(this.getTimezoneOffset()/1440)+2440587.5);
};

Date.prototype.oaDate=function() {
	return (this-new Date(Date.UTC(1899, 11, 30)))/(24*60*60*1000);
};

exports.execute=function(config, callback) {
	var cols=config.cols,
		data=config.rows,
		colsLength=cols.length,
		p,
		files=[],
		styleIndex,
		k=0,
		cn=1,
		dirPath,
		shareStrings=[],
		convertedShareStrings="",
		sheet,
		sheetPos=0;

	var write=function(str, callback) {
		var buf=new Buffer(str);
		var off=0;
		var written=0;

		return async.whilst(
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

	return async.waterfall([
			function(callback) {
				return temp.mkdir('xlsx', function(err, dir) {
					if(err) {
						return callback(err);
					}

					dirPath=dir;
					return callback();
				});
			},
			function(callback) {
				return fs.mkdir(path.join(dirPath, 'xl'), callback);
			},
			function(callback) {
				return fs.mkdir(path.join(dirPath, 'xl', 'worksheets'), callback);
			},
			function(callback) {
				return async.parallel([
					function(callback) {
						return fs.writeFile(path.join(dirPath, 'data.zip'), template.XLSX, callback);
					},
					function(callback) {
						if(!config.stylesXmlFile) {
							return callback();
						}

						p=config.stylesXmlFile || __dirname+'/resources/styles.xml';
						return fs.readFile(p, 'utf8', function(err, styles) {
							if(err) {
								return callback(err);
							}

							p=path.join(dirPath, 'xl', 'styles.xml');
							files.push(p);
							return fs.writeFile(p, styles, callback);
						});
					}
				], function(err) {
					return callback(err);
				})
			},
			function(callback) {
				p=path.join(dirPath, 'xl', 'worksheets', 'sheet.xml');
				files.push(p);
				return fs.open(p, 'a+', function(err, fd) {
					if(err) {
						return callback(err);
					}

					sheet=fd;

					return callback();
				});
			},
			function(callback) {
				return write(template.sheetFront, callback);
			},
			function(callback) {
				return async.eachSeries(cols, function(col, callback) {
					var colStyleIndex=col.styleIndex || 0;
					var res=util.format('<x:col min="%d" max="%d" width="%d" customWidth="1" style="%d"/>',
						cn, cn, (col.width ? col.width : 10), colStyleIndex);
					cn++;
					return write(res, callback);
				}, callback);
			},
			function(callback) {
				return write('</cols><x:sheetData>', callback);
			},
			function(callback) {
				return write('<x:row r="1" spans="1:'+colsLength+'">', callback);
			},
			function(callback) {
				return async.eachSeries(cols, function(col, callback) {
					var colStyleIndex=col.captionStyleIndex || 0;
					var res=addStringCol(getColumnLetter(k+1)+1, col.caption, colStyleIndex, shareStrings);
					k++;
					convertedShareStrings+=res[1];
					return write(res[0], callback);
				}, callback);
			},
			function(callback) {
				return write('</x:row>', callback);
			},
			function(callback) {
				var j, r, cellData, currRow, cellType;
				var i=-1;

				data.reverse();

				return async.whilst(
					function() {
						return data.length>0;
					},
					function(callback) {
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

						return write(row, callback);
					},
					callback
				);
			},
			function(callback) {
				return write(template.sheetBack, callback);
			},
			function(callback) {
				return fs.close(sheet, callback);
			},
			function(callback) {
				if(shareStrings.length===0) {
					return callback();
				}

				var sharedStringsFront=template.sharedStringsFront.replace(/\$count/g, shareStrings.length);
				p=path.join(dirPath, 'xl', 'sharedStrings.xml');
				files.push(p);
				return fs.writeFile(p, sharedStringsFront+convertedShareStrings+template.sharedStringsBack, callback);
			},
			function(callback) {
				var zipfile=new zipper(path.join(dirPath, 'data.zip'));
				files.forEach(function(file) {
					var relative=path.relative(dirPath, file);
					return zipfile.addLocalFile(file);
				});
				fs.readFile(path.join(dirPath, 'data.zip'), callback);
			}],
		function(err, data) {
			if(err) {
				return callback(err);
			}

			temp.cleanup();
			return callback(null, data);
		}
	);
};

var startTag=function(obj, tagName, closed) {
	var result="<"+tagName, p;
	for(p in obj) {
		result+=" "+p+"="+obj[p];
	}
	if(!closed) {
		result+=">";
	} else {
		result+="/>";
	}

	return result;
};

var endTag=function(tagName) {
	return "</"+tagName+">";
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
