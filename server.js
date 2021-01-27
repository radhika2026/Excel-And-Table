const express = require('express')
const bodyParser = require('body-parser')
const multer = require('multer');
const path = require('path');
var mysql = require('mysql');
const app = express();
var fs = require('fs');
var xlsx = require('node-xlsx');
const fastcsv = require("fast-csv");
app.use(bodyParser.urlencoded({
  extended: true
}))
app.set('view engine', 'ejs');
app.get('/', function(req, res) {
  res.sendFile(__dirname + '/home.html');
});
var storage = multer.diskStorage({
  destination: function(req, file, cb) {
    cb(null, './public/uploads')
  },
  filename: function(req, file, cb) {
    var ext = file.originalname.split('.').pop();
      cb(null, 'name' + path.extname(file.originalname));
  }
})
var upload = multer({
  storage: storage
})
app.post('/uploadfile', upload.single('myFile'), (req, res, next) => {
  const file = req.file
  if (!file) {
    const error = new Error('Please upload a file')
    error.httpStatusCode = 400
    return next(error)
  }
  exceltocsv();
  deletetable();
  csvtotable();
  res.sendFile(__dirname + '/index.html');
})

function exceltocsv(){
var obj = xlsx.parse(__dirname + '/public/uploads/name.xlsx'); // parses a file
var rows = [];
var writeStr = "";
//looping through all sheets
for(var i = 0; i < obj.length; i++)
{
    var sheet = obj[i];
    //loop through all rows in the sheet
    for(var j = 0; j < sheet['data'].length; j++)
    {
            //add the row to the rows array
            rows.push(sheet['data'][j]);
    }
}
//creates the csv string to write it to a file
for(var i = 0; i < rows.length; i++)
{
    writeStr += rows[i].join(",") + "\n";
}
//writes to a file, but you will presumably send the csv as a
//response instead
fs.writeFile(__dirname + "/test.csv", writeStr, function(err) {
    if(err) {
        return console.log(err);
    }
    console.log("test.csv was saved in the current directory!");
});
}


function deletetable(){
  var con = mysql.createConnection({
    host: "localhost",
    user: "root",
    password: "1234",
    database: "test"
  });
  con.connect(function(err) {
    if (err) throw err;
    var sql = "DELETE FROM student";
    con.query(sql, function (err, result) {
      if (err) throw err;
      console.log("Table deleted");
    });
  });
}

function csvtotable(){
let stream = fs.createReadStream("test.csv");
let csvData = [];
let csvStream = fastcsv
  .parse()
  .on("data", function(data) {
    csvData.push(data);
  })
  .on("end", function() {
    // remove the first line: header
    csvData.shift();
    // create a new connection to the database
    const connection = mysql.createConnection({
      host: "localhost",
      user: "root",
      password: "1234",
      database: "test"
    });
    // open the connection
    connection.connect(error => {
      if (error) {
        console.error(error);
      } else {
        let query =
          "INSERT INTO student (Name, Roll_no, Class) VALUES ?";
        connection.query(query, [csvData], (error, response) => {
          console.log(error || response);
        });
      }
    });
  });
stream.pipe(csvStream);
}


app.post('/table', function(request, response){
    fetchData(response);
    console.log('Done. Data displayed!');
});
var db = mysql.createConnection({
    host: 'localhost',
    user: 'root',
    password: '1234',
    database: 'test'
});


//now we have to create the connection to the database


db.connect(function(err){
    if(err){throw err;}
    console.log("Connected to the database!!");
});


function executeQuery(sql, cb){
    db.query(sql, function(error, result, fields){
        if(error){throw err;}
        cb(result);
    });
}


function fetchData(response){
    executeQuery("SELECT * FROM student", function(result){
        response.render('excel.ejs', {result: result})
    });
}

app.listen(process.env.PORT || 3000, () => console.log('Server started on port 3000'));
