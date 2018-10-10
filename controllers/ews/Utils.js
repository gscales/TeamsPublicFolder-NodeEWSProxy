
(function () {

    const ParseServerFromResultXML = (XMLString) => {
        return new Promise(
            (resolve, reject) => {
                console.log("converting " + XMLString);
				var parseString = require('xml2js').parseString;               
                parseString(XMLString, function (err, result) {
                    //console.log(result);
                    resolve(result);
                    console.log("resolved");
                });
            }
        );
    }
  

    async function ParseServerFromResult(XMLString) {
        try {           
            let ServerVal = await ParseServerFromResultXML(XMLString);
            return ServerVal;
        }
        catch (error) {
            console.log(error.message);
        }
    }

    module.exports.ParseServerFromResult = function (XMLString) {
        return ParseServerFromResult(XMLString);
    }

}());
