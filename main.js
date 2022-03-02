var express = require('express');
var bodyParser = require('body-parser')
//var excel = require('exceljs')
// var xlsx = require('node-xlsx').default;
var xlsx = require('xlsx')
const aws = require('aws-sdk')
const fs = require('fs');
var workbook;
var ws;
var value;

var port = process.env.PORT || 5000;
const CURRENT_SLOTS_INTENT = 'CurrentSlots'; //this intent is to show the current status of slots/speeches
const S3_BUCKET = process.env.S3_BUCKET;

aws.config.region = 'us-east-2';
// aws.config.accessKeyId = "AKIAIJPPGW2JDEODIW7A";
// aws.config.secretAccessKey = "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";

var appexp = express();
//var workbook = new excel.Workbook();
const s3 = new aws.S3();
var tempbuf = [];

var params = {
  Bucket: S3_BUCKET, 
  Key: "meeting agenda.xlsx"
 };
 
var file = s3.getObject(params).createReadStream();   //reading from AWS S3

file.on('data', function (data) {         //collecting as buffer stream
        tempbuf.push(data);
    });

	file.on('end', function () {
        var buf = Buffer.concat(tempbuf);
		console.log(buf);
        workbook = xlsx.read(buf, {type: "buffer"});
		ws = workbook.Sheets["Sheet1"];
		// console.log("worksheet contents : ", ws);
		console.log("Value of cell c16 is : ", ws['C16'].v);
		// console.log("row value : ", workbook.sheetRows['16']);
    });

appexp.use(bodyParser.json());
appexp.use(bodyParser.urlencoded({ extended : true }));

appexp.get('/', (req, res) => {
	res.send("Hello world from heroku node");
});

appexp.post('/webhook', (req, res) => {
	console.log('Received a post request. body is : ', req.body);
	if(!req.body) console.log('Body is undefined');
	if(!req.body) return res.sendStatus(400);
	console.log('The intent is : ', req.body.queryResult.intent.displayName);
	var response = " ";
	var responseObj;
	if(req.body.queryResult.intent.displayName = CURRENT_SLOTS_INTENT) {
		console.log('Intent matched to current slots intent');
		responseObj={
			"fulfillmentText":response,
			"fulfillmentMessages": [{"text": {"text": [ws['C16'].v]} }],
			"source":"",
		}
	}
	console.log('Here is the response to dialogflow :');
	console.log(responseObj);
	return res.json(responseObj);
		
});

appexp.listen(port, () => {});