const express = require('express');
const app = express();
const cors = require('cors');
const mind = require('./routes/mindmastery');
const excel = require('./helpers/excel');
const path = require('path');
const fs = require('fs');
app.use(cors());
app.use(mind)
app.use('/uploads', express.static(path.join(__dirname + '/uploads/')));

app.get('/uploads/:name', function (req, res) {
    var filePath = "/uploads/name";
    fs.readFile(__dirname + filePath, function (err, data) {
        res.contentType("application/pdf");
        res.send(data);
    });
});

app.get('/', function (req, res) {
    res.send("Working fine");
});

app.listen(3000, (err) => {
    if (err) {
        console.log(err);
    } else {
        console.log("Server is running on 3000");
    }
})