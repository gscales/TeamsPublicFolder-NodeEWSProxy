"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
var fetch_1 = require("fetch");
var Promise_1 = require("./Promise");
/** @internal */
var XHRDefaults = (function () {
    function XHRDefaults() {
        try {
            var fetch_2 = require("fetch");
            XHRDefaults.FetchStream = fetch_2.FetchStream;
            XHRDefaults.fetchUrl = fetch_2.fetchUrl;
        }
        catch (e) { }
    }
    XHRDefaults.prototype.xhr = function (xhroptions, progressDelegate) {
        if (XHRDefaults.fetchUrl === null) {
            throw new Error("xhrApi - stub method, must be bootstrapped");
        }
        //setup xhr for github.com/andris9/fetch options
        var options = {
            url: xhroptions.url,
            payload: xhroptions.data,
            headers: xhroptions.headers,
            method: xhroptions.type
        };
        // xhroptions["payload"] = xhroptions.data;
        // delete xhroptions["data"];
        // xhroptions["method"] = xhroptions.type;
        // delete xhroptions["type"];
        return new Promise_1.Promise(function (resolve, reject) {
            XHRDefaults.fetchUrl(xhroptions.url, options, function (error, meta, body) {
                if (error) {
                    if (typeof error.status === 'undefined') {
                        error.status = 0;
                    }
                    reject(setupXhrResponse(error));
                }
                else {
                    var xhrResponse = {
                        response: body.toString(),
                        status: meta.status,
                        redirectCount: meta.redirectCount,
                        headers: meta.responseHeaders,
                        finalUrl: meta.finalUrl,
                        responseType: '',
                        statusText: undefined,
                    };
                    if (xhrResponse.status === 200) {
                        resolve(setupXhrResponse(xhrResponse));
                    }
                    else {
                        reject(setupXhrResponse(xhrResponse));
                    }
                }
            });
        });
    };
    XHRDefaults.prototype.xhrStream = function (xhroptions, progressDelegate) {
        var _this = this;
        if (XHRDefaults.FetchStream === null) {
            throw new Error("xhrApi - stub method, must be bootstrapped");
        }
        //setup xhr for github.com/andris9/fetch options
        var options = {
            payload: xhroptions.data,
            headers: xhroptions.headers,
            method: xhroptions.type
        };
        return new Promise_1.Promise(function (resolve, reject) {
            _this.stream = new XHRDefaults.FetchStream(xhroptions.url, options);
            _this.stream.on("data", function (chunk) {
                //console.log(chunk.toString());
                progressDelegate({ type: "data", data: chunk.toString() });
            });
            _this.stream.on("meta", function (meta) {
                progressDelegate({ type: "header", headers: meta["responseHeaders"] });
            });
            _this.stream.on("end", function () {
                progressDelegate({ type: "end" });
                resolve();
            });
            _this.stream.on('error', function (error) {
                progressDelegate({ type: "error", error: error });
                _this.disconnect();
                reject(error);
            });
        });
    };
    XHRDefaults.prototype.disconnect = function () {
        if (this.stream) {
            try {
                this.stream.destroy();
            }
            catch (e) { }
        }
    };
    Object.defineProperty(XHRDefaults.prototype, "apiName", {
        get: function () {
            return "default";
        },
        enumerable: true,
        configurable: true
    });
    return XHRDefaults;
}());
XHRDefaults.FetchStream = fetch_1.FetchStream;
XHRDefaults.fetchUrl = null;
exports.XHRDefaults = XHRDefaults;
/** @internal */
function setupXhrResponse(xhrResponse) {
    xhrResponse["responseText"] = xhrResponse.response;
    delete xhrResponse["response"];
    xhrResponse.getAllResponseHeaders = function () {
        var header = "";
        if (xhrResponse.headers) {
            for (var key in xhrResponse.headers) {
                header += key + " : " + xhrResponse.headers[key] + "\r\n";
            }
        }
        return header;
    };
    xhrResponse.getResponseHeader = function (header) {
        if (header) {
            if (xhrResponse.headers) {
                if (xhrResponse.headers[header]) {
                    return xhrResponse.headers[header];
                }
                if (xhrResponse.headers[header.toLocaleLowerCase()]) {
                    return xhrResponse.headers[header.toLocaleLowerCase()];
                }
            }
        }
        return null;
    };
    return xhrResponse;
}
