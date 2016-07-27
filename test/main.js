// test/main.js
var assert = require('chai').assert;
var nodeExcel = require('../index');


describe('Simple Excel xlsx Export', function() {
    describe('Export', function() {
        it('returns path', function(done) {
			var conf ={};
			conf.cols = [
				{caption:'string', type:'string'},
				{caption:'date', type:'date'},
				{caption:'bool', type:'bool'},
				{caption:'number 2', type:'number'}
			];
			conf.rows = [
        ['pi', (new Date(Date.UTC(2013, 4, 1))).oaDate(), true, 3.14],
        ["e", (new Date(2012, 4, 1)).oaDate(), false, 2.7182],
        ["M&M<>'", (new Date(Date.UTC(2013, 6, 9))).oaDate(), false, 1.2],
        ["null", null, null, null]
			];

        nodeExcel.execute(conf, function(err, path) {
            if (err) throw err;
            var fs = require('fs');
            fs.exists(path, (exists) => {
                assert(exists, "file does not exist");
                done();
            });
	    });
        });
    });
});
