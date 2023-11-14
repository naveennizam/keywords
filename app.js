var createError = require('http-errors');
var express = require('express');
var path = require('path');
var cookieParser = require('cookie-parser');
var logger = require('morgan');

var XLSX = require("xlsx");

var indexRouter = require('./routes/index');
var usersRouter = require('./routes/users');

var app = express();

// view engine setup
app.set('views', path.join(__dirname, 'views'));
app.set('view engine', 'pug');

app.use(logger('dev'));
app.use(express.json());
app.use(express.urlencoded({ extended: false }));
app.use(cookieParser());
app.use(express.static(path.join(__dirname, 'public')));

app.use('/', indexRouter);
app.use('/users', usersRouter);
app.use('/keywords', async function (req, res) {
  var wb = XLSX.readFile('kw.xlsx')
  const columnA = [];

  var worksheet = (wb.Sheets["A"])


  for (let z in worksheet) {
    if (z.toString()[0] === 'A') {
      columnA.push(worksheet[z].v);
    }
  }


  let text;
  text = "<ol>"
  for (let i = 0; i < columnA.length; i++) {
    text += "<li>" + columnA[i] + "</li>";
  }
  text += "</ol>";
  res.send(`<h1 style=" display: flex;
align-items: center;
justify-content: center;">KEYWORDS</h1>

<div  style=" display: flex;
align-items: center;
justify-content: center;">
${text}
</div>` )

})




// catch 404 and forward to error handler
app.use(function (req, res, next) {
  next(createError(404));
});

// error handler
app.use(function (err, req, res, next) {
  // set locals, only providing error in development
  res.locals.message = err.message;
  res.locals.error = req.app.get('env') === 'development' ? err : {};

  // render the error page
  res.status(err.status || 500);
  res.render('error');
});

module.exports = app;
