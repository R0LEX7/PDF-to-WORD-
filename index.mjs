// load the module
var GroupDocs = require('groupdocs-conversion-cloud');

// get your appSid and appKey at https://dashboard.groupdocs.cloud (free registration is required).
var appSid = "XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX";
var appKey = "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX";

// construct Api
var api = GroupDocs.InfoApi.fromKeys(appSid, appKey);
var request = new GroupDocs.GetSupportedConversionTypesRequest();
// retrieve supported conversion types
api.getSupportedConversionTypes(request)
    .then(function (response) {
        console.log("Supported file-formats:")
        response.forEach(function (format) {
            console.log(format.sourceFormat + ": [" + format.targetFormats.join(", ") + "]");
        });
    })
    .catch(function (error) {
        console.log("Error: " + error.message)
    });