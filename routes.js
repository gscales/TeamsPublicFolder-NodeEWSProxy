'use strict';
module.exports = function(app) {
  var ewsController = require(__dirname + '/controllers/ewsController');


  app.route('/FolderItems')
    .post(ewsController.getFolderItems)
  

};
