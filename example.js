
var qs = require("./sheets.js");
qs.insert(
    {
        'name': 'Lorem ipsum',
        'contact': 'lorem.ipsum@gmail.com',
        'message': 'Test.Test.',
        'mathAuth': 14
    },
    '<<<SHEET_ID>>>',
    '<<<SHEET_NAME>>>').then(() => console.log('ok'), (err) => console.error(err));