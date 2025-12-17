const express = require('express');
const multer = require('multer');
const bodyParser = require('body-parser');
const mindcontroller = require('../controllers/mindcontroller');
const updateXl = require('../controllers/updateXl');
const router = express.Router();
router.use(bodyParser.json());
router.use(bodyParser.urlencoded({
    extended: false
}));
const storage = multer.diskStorage({
    destination: function (req, file, cb) {
        return cb(null, './uploads')
    },
    filename: function (req, file, cb) {
        return cb(null, file.originalname)
    }
})
const uploader = multer({
    storage: storage,
});
router.post('/uploadpdf', uploader.single("pdfFile"), (req, res) => {
    console.log("we ate hrtr")
    mindcontroller.uploadPdf(req, res)
})
router.get('/testxl', (req, res) => {
    console.log("test xl");
    updateXl.updateXlsmSafe(req, res)
})
module.exports = router