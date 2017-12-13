let express = require('express');
const app = express();
let bodyParser = require('body-parser');
let multer = require('multer');
let exceltojson = require('xls-to-json-lc');
let xlstojson = require('xls-to-json-lc');
let xlsxtojson = require('xlsx-to-json-lc');
let fs = require('fs');

app.use(bodyParser.json());

const storage = multer.diskStorage({
  destination(req, file, cb) {
    cb(null, './uploads')
  },
  filename(req, file, cb) {
    const datetimestamp = Date.now();
    cb(null, `${file.fieldname}-${datetimestamp}.${file.originalname.split('.')[file.originalname.split('.').length - 1]}`)
  }
});

const upload = multer({
  storage,
  fileFilter(req, file, callback) {
    if(!['xls', 'xlsx'].includes(file.originalname.split('.')[file.originalname.split('.').length - 1])){
      return callback(new Error('Wrong extension type'));
    }
    callback(null, true);
  }
}).single('file');

app.post('/upload', (req, res) => {
  let exceltojson;
  upload(req, res, err => {
    if(err){
      res.json({error_code:1, err_desc:err});
      return;
    }
    if (!req.file) {
      res.json({error_code:1, err_desc:"No file passed"})
      return;
    }
    if (req.file.originalname.split('.')[req.file.originalname.split('.').length - 1] === 'xlsx') {
      exceltojson = xlsxtojson;
    } else {
      exceltojson = xlstojson
    }
    try {
      exceltojson({
        input: req.file.path,
        output: null,
        lowerCaseHeaders: true
      }, (err, result) => {
        if (err) {
          return res.json({error_code:1, err_desc:err, data: null});
        }
        res.json({error_code:0, err_desc:null, data: result});
        fs.unlinkSync(req.file.path); //deleting the uploaded file
      });
    } catch (e) {
      res.json({error_code: 1, err_desc:"Corupted excel file"});
    }
  });
});

app.get('/', (req, res) => {
  res.sendFile(`${__dirname}/index.html`);
});

app.listen('3000', () => {
  console.log('running on 3000');
});
