// test/main.js
var async=require("async");
var should=require('should');
var nodeExcel=require('../index');


describe('Simple Excel xlsx Export', function() {
	describe('Export', function() {
		var conf={
			cols: [
				{caption: 'string', type: 'string'},
				{caption: 'date', type: 'date'},
				{caption: 'bool', type: 'bool'},
				{caption: 'number 2', type: 'number'}
			],
			rows: [
				['pi', (new Date(Date.UTC(2013, 4, 1))).oaDate(), true, 3.14],
				["e", (new Date(2012, 4, 1)).oaDate(), false, 2.7182],
				["M&M<>'", (new Date(Date.UTC(2013, 6, 9))).oaDate(), false, 1.2]
			]
		};

		it('returns buffer', function(done) {
			nodeExcel.execute(conf, function(err, result) {
				should.not.exist(err);
				should.exist(result);
				should.equal(result.constructor, Buffer);
				done();
			});
		});

		it('returns path', function(done) {
			async.times(10, function(index, _done) {
				nodeExcel.execute(conf, "./test/d.xlsx"+index, function(err, result) {
					should.not.exist(err);
					should.equal(result, "./test/d.xlsx"+index);
					_done();
				});
			}, done);
		});
	});
});
