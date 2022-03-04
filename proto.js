/**
*	prototype for section tally
* James Mercantante
*
*/

// define requirements
const XLSX = require("xlsx");
const XMLHttpRequest = require("xhr2");
const fs = require("fs");
const xlsx = require("node-xlsx");
const test = new XMLHttpRequest();
// term for url
const current_term = '202230';
// section tally url
var url = "https://banner.rowan.edu/reports/reports.pl?term=" + current_term +
	"&task=Section_Tally&coll=ALL&dept=ALL&subj=ALL&ptrm=ALL&sess=ALL&prof=ALL&attr=ALL&camp=ALL&bldg=ALL&Search=Search&format=excel";
const { MongoClient, ServerApiVersion } = require('mongodb');
// mongodb database url call
const uri = "mongodb+srv://root:Senior-project321@cluster0.u1zph.mongodb.net/saction-tally1?retryWrites=true&w=majority";
const client = new MongoClient(uri, { useNewUrlParser: true, useUnifiedTopology: true, serverApi: ServerApiVersion.v1 });

//call first function
start();

/**
*	Get xls file from secion tally based on the term in url
*	converts to json then writes to file
*
*/
function start(){
		req = new XMLHttpRequest();
		req.responseType = 'arraybuffer';

		req.open('GET', url, true);
		req.send();

		req.onload = function (e) {
      console.log(current_term + " response");
      write_json(to_json_str(req.response));
  };
};


/**
*	Takes http response and converts it to an array
*	then calls function to write json to file
*	@param http response to jsonify
*/
function to_json_str(a) {
  var arraybuffer = a;
  var data = new Uint8Array(arraybuffer);
  var arr = new Array();
  for (var x = 0; x != data.length; ++x) arr[x] = String.fromCharCode(data[x]);
  var bstr = arr.join("");
  var workbook = XLSX.read(bstr, { type: "binary" });
  // read first sheet (identified by first of SheetNames)
  let sheet = workbook.Sheets[workbook.SheetNames[0]];
  // convert to JSON
  var json = XLSX.utils.sheet_to_json(sheet);
  //make_csv()
  var str = JSON.stringify(json);
  return str;
}

/**
*	writes json string to file
*	@param {string} string to write
*/
function write_json(string) {
  fs.writeFile(current_term + ".json", string, function (err) {
    if (err) {
      return console.log(err);
    }
    console.log("written");
  });
	upload();
}

/**
* Connects to collection1 in mongodb database,
*	prints first document in collection and uploads
* sample data
*
*/
function upload(){
	// connect to collection1
	client.connect(err => {
	  const collection = client.db("saction-tally1").collection("collection1");

		// return first document in collection
	  collection.findOne({}, function(err, res){
	  	if (err) throw err;
	  	console.log(res);
	  });

	// example data to be inserted
	var doc = {
	      CRN: "12345",
	      Subj: "JM",
	      Crse: "123456",
	      Sect: "  1",
	      "Part of Term": "Full Term ... \n06-SEP to 21-DEC",
	      Session: "Day",
	      Title: "Test Title.\n",
	      Prof: "James, Merc",
	      "Day  Beg   End   Bldg Room  (Type)":
	        "MW      1230 1345 BUSN 121 (Class)",
	      Campus: "Main",
	      AddlInfo: "",
	      Hrs: 3,
	      Max: 20,
	      MaxResv: 0,
	      LeftResv: 0,
	      Enr: 0,
	      Avail: 20,
	      WaitCap: 0,
	      WaitCount: 0,
	      WaitAvail: 0,
	      "Room Cap": "40",
	    };

		// insert example into database
	  collection.insertOne(doc, function(err, res){
	  	if (err) throw err;
	  	console.log('doc inserted');
			// check for inserted data
			collection.findOne({Subj: 'JM'}, function(err, res){
				if (err) throw err;
				console.log(res);
				client.close();
			});
	  });


	});
}
