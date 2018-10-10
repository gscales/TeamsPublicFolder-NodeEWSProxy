'use strict';
var express = require('express');
var logger = require('connect-logger');
var cookieParser = require('cookie-parser');
var session = require('cookie-session');
var fs = require('fs');

var crypto = require('crypto');


var https = require('https');
var http = require('http');

var bodyParser = require('body-parser');
//var urlencodedParser = bodyParser.urlencoded({ extended: true });



//Start Express
var app = express();
app.use(logger());
app.use(bodyParser.urlencoded({
    extended: true
}));
app.use(bodyParser.json());
app.use(function(req, res, next) {
    res.header("Access-Control-Allow-Origin", "*");
    res.header("Access-Control-Allow-Credentials", "true");
    res.header("Access-Control-Allow-Methods", "GET,HEAD,OPTIONS,POST,PUT");
    res.header("Access-Control-Allow-Headers", "Access-Control-Allow-Headers, Origin,Accept, X-Requested-With, Content-Type, Access-Control-Request-Method, Access-Control-Request-Headers, Authorization");
    next();
  });
var routes = require(__dirname  + '/routes'); //importing route
routes(app); //register the route

console.log("Starting");
var server = app.listen(3007);
console.log('listening on 3007');


