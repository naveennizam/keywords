var express = require('express');
var router = express.Router();
var XLSX = require("xlsx");
/* GET home page. */
router.get('/', function(req, res, next) {
 var wb =  XLSX.readFile('kw.xlsx')
const columnA = [];
 var worksheet = (wb.Sheets["A"])
 for (let z in worksheet) {
  
  if(z.toString()[0] === 'A'){
    columnA.push(worksheet[z].v);
  }
}


  res.render('index', { title: columnA });
});

module.exports = router;
